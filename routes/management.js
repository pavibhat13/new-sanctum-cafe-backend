const express = require('express');
const router = express.Router();
const ExcelJS = require('exceljs');
const DailySales = require('../models/DailySales');
const PurchaseHeader = require('../models/PurchaseHeader');
const PurchaseLine = require('../models/PurchaseLine');
const OnlineSettlement = require('../models/OnlineSettlement');
const Expense = require('../models/Expense');
const Salary = require('../models/Salary');
const Vendor = require('../models/Vendor');
const MasterValue = require('../models/MasterValue');
const ChecklistLog = require('../models/ChecklistLog');
const ManagementInventory = require('../models/ManagementInventory');
const Config = require('../models/Config');
const { authenticateToken, requireOwner } = require('../middleware/auth');

router.use(authenticateToken);

const normalizeItemName = (name) =>
  name.toLowerCase().replace(/[^a-z0-9]/g, '').replace(/(.)\1+/g, '$1');

const getDateRangeQuery = (fromDate, toDate) => {
  if (!fromDate && !toDate) return null;
  const q = {};
  if (fromDate) { const s = new Date(fromDate); s.setHours(0, 0, 0, 0); q.$gte = s; }
  if (toDate) { const e = new Date(toDate); e.setHours(23, 59, 59, 999); q.$lte = e; }
  return q;
};

const updateInventoryStock = async (itemName, quantityChange, role) => {
  const normInput = normalizeItemName(itemName);
  const allInventory = await ManagementInventory.find({});
  let inv = allInventory.find(i => normalizeItemName(i.item) === normInput);
  if (inv) {
    inv.purchasedQty = (inv.purchasedQty || 0) + quantityChange;
    inv.closingStock = (inv.openingStock || 0) + inv.purchasedQty - (inv.usedQty || 0);
    await inv.save();
  } else if (quantityChange > 0) {
    await new ManagementInventory({
      item: itemName, category: 'Other', unit: 'Pkt',
      openingStock: 0, purchasedQty: quantityChange, usedQty: 0,
      closingStock: quantityChange, createdBy: role,
    }).save();
  }
};

// ── Daily Sales ──────────────────────────────────────────────────────────────

router.get('/daily-sales', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate } = req.query;
    const query = {};
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    res.json(await DailySales.find(query).sort({ date: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/daily-sales', requireOwner, async (req, res) => {
  try {
    const { date, cash, upi, swiggy, zomato, notes } = req.body;
    const s = new Date(date); s.setHours(0, 0, 0, 0);
    const e = new Date(date); e.setHours(23, 59, 59, 999);
    const existing = await DailySales.findOne({ date: { $gte: s, $lte: e } });
    if (existing) return res.status(400).json({ message: 'Sales entry for this date already exists' });
    const entry = await new DailySales({ date, cash, upi, swiggy, zomato, notes, createdBy: req.user.role }).save();
    res.status(201).json(entry);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/daily-sales/:id', requireOwner, async (req, res) => {
  try {
    const entry = await DailySales.findById(req.params.id);
    if (!entry) return res.status(404).json({ message: 'Entry not found' });
    const { date, cash, upi, swiggy, zomato, notes } = req.body;
    if (date !== undefined) {
      const s = new Date(date); s.setHours(0, 0, 0, 0);
      const e = new Date(date); e.setHours(23, 59, 59, 999);
      const dup = await DailySales.findOne({ date: { $gte: s, $lte: e }, _id: { $ne: req.params.id } });
      if (dup) return res.status(400).json({ message: 'Sales entry for this date already exists' });
      entry.date = new Date(date);
    }
    if (cash !== undefined) entry.cash = cash;
    if (upi !== undefined) entry.upi = upi;
    if (swiggy !== undefined) entry.swiggy = swiggy;
    if (zomato !== undefined) entry.zomato = zomato;
    if (notes !== undefined) entry.notes = notes;
    res.json(await entry.save());
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/daily-sales/:id', requireOwner, async (req, res) => {
  try {
    const entry = await DailySales.findByIdAndDelete(req.params.id);
    if (!entry) return res.status(404).json({ message: 'Entry not found' });
    res.json({ message: 'Deleted' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Purchase Headers ─────────────────────────────────────────────────────────

router.get('/purchase-headers', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate, vendor } = req.query;
    const query = {};
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    if (vendor) query.vendor = vendor;
    res.json(await PurchaseHeader.find(query).sort({ date: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/purchase-headers', requireOwner, async (req, res) => {
  try {
    const header = await new PurchaseHeader({ ...req.body, createdBy: req.user.role }).save();
    res.status(201).json(header);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/purchase-headers/:id', requireOwner, async (req, res) => {
  try {
    const header = await PurchaseHeader.findByIdAndUpdate(req.params.id, req.body, { new: true, runValidators: true });
    if (!header) return res.status(404).json({ message: 'Not found' });
    res.json(header);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/purchase-headers/:id', requireOwner, async (req, res) => {
  try {
    const header = await PurchaseHeader.findByIdAndDelete(req.params.id);
    if (!header) return res.status(404).json({ message: 'Not found' });
    await PurchaseLine.deleteMany({ purchaseHeader: req.params.id });
    res.json({ message: 'Deleted with associated lines' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Purchase Lines ────────────────────────────────────────────────────────────

router.get('/purchase-lines', requireOwner, async (req, res) => {
  try {
    const { billNo, fromDate, toDate } = req.query;
    let query = {};
    if (billNo) {
      const h = await PurchaseHeader.findOne({ billNo });
      if (!h) return res.json([]);
      query.purchaseHeader = h._id;
    }
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) {
      const headers = await PurchaseHeader.find({ date: dq }).select('_id');
      const ids = headers.map(h => h._id);
      if (query.purchaseHeader) {
        if (!ids.some(id => id.equals(query.purchaseHeader))) return res.json([]);
      } else {
        query.purchaseHeader = { $in: ids };
      }
    }
    res.json(await PurchaseLine.find(query).populate('purchaseHeader', 'billNo vendor date').sort({ createdAt: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/purchase-lines', requireOwner, async (req, res) => {
  try {
    const newLines = Array.isArray(req.body) ? req.body : [req.body];
    const saved = [];
    for (const d of newLines) {
      const line = await new PurchaseLine({ ...d, createdBy: req.user.role }).save();
      await updateInventoryStock(line.item, line.quantity, req.user.role);
      saved.push(line);
    }
    res.status(201).json(saved);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/purchase-lines/:id', requireOwner, async (req, res) => {
  try {
    const line = await PurchaseLine.findById(req.params.id);
    if (!line) return res.status(404).json({ message: 'Not found' });
    const oldQty = line.quantity, oldItem = line.item;
    const { purchaseHeader, item, quantity, rate } = req.body;
    if (purchaseHeader !== undefined) line.purchaseHeader = purchaseHeader;
    if (item !== undefined) line.item = item;
    if (quantity !== undefined) line.quantity = quantity;
    if (rate !== undefined) line.rate = rate;
    const saved = await line.save();
    if (item !== undefined || quantity !== undefined) {
      await updateInventoryStock(oldItem, -oldQty, req.user.role);
      await updateInventoryStock(saved.item, saved.quantity, req.user.role);
    }
    res.json(saved);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/purchase-lines/:id', requireOwner, async (req, res) => {
  try {
    const line = await PurchaseLine.findById(req.params.id);
    if (!line) return res.status(404).json({ message: 'Not found' });
    await updateInventoryStock(line.item, -line.quantity, req.user.role);
    await PurchaseLine.findByIdAndDelete(req.params.id);
    res.json({ message: 'Deleted' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Online Settlements ────────────────────────────────────────────────────────

router.get('/online-settlements', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate, platform } = req.query;
    const query = {};
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.paymentDate = dq;
    if (platform) query.platform = platform;
    res.json(await OnlineSettlement.find(query).sort({ paymentDate: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/online-settlements', requireOwner, async (req, res) => {
  try {
    const s = await new OnlineSettlement({ ...req.body, createdBy: req.user.role }).save();
    res.status(201).json(s);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/online-settlements/:id', requireOwner, async (req, res) => {
  try {
    const s = await OnlineSettlement.findById(req.params.id);
    if (!s) return res.status(404).json({ message: 'Not found' });
    Object.assign(s, req.body);
    res.json(await s.save());
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/online-settlements/:id', requireOwner, async (req, res) => {
  try {
    const s = await OnlineSettlement.findByIdAndDelete(req.params.id);
    if (!s) return res.status(404).json({ message: 'Not found' });
    res.json({ message: 'Deleted' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Expenses ──────────────────────────────────────────────────────────────────
// Owner sees all; staff sees only their own (createdBy === 'staff')

router.get('/expenses', async (req, res) => {
  try {
    const { fromDate, toDate, category } = req.query;
    const query = {};
    if (req.user.role === 'staff') query.createdBy = 'staff';
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    if (category) query.category = category;
    res.json(await Expense.find(query).sort({ date: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/expenses', async (req, res) => {
  try {
    const expense = await new Expense({ ...req.body, createdBy: req.user.role }).save();
    res.status(201).json(expense);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/expenses/:id', requireOwner, async (req, res) => {
  try {
    const expense = await Expense.findByIdAndUpdate(req.params.id, req.body, { new: true, runValidators: true });
    if (!expense) return res.status(404).json({ message: 'Not found' });
    res.json(expense);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/expenses/:id', requireOwner, async (req, res) => {
  try {
    const expense = await Expense.findByIdAndDelete(req.params.id);
    if (!expense) return res.status(404).json({ message: 'Not found' });
    res.json({ message: 'Deleted' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Salaries ──────────────────────────────────────────────────────────────────

router.get('/salaries', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate } = req.query;
    const query = {};
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    res.json(await Salary.find(query).sort({ date: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/salaries', requireOwner, async (req, res) => {
  try {
    const salary = await new Salary({ ...req.body, createdBy: req.user.role }).save();
    res.status(201).json(salary);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/salaries/:id', requireOwner, async (req, res) => {
  try {
    const salary = await Salary.findByIdAndUpdate(req.params.id, req.body, { new: true, runValidators: true });
    if (!salary) return res.status(404).json({ message: 'Not found' });
    res.json(salary);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/salaries/:id', requireOwner, async (req, res) => {
  try {
    const salary = await Salary.findByIdAndDelete(req.params.id);
    if (!salary) return res.status(404).json({ message: 'Not found' });
    res.json({ message: 'Deleted' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Vendors ───────────────────────────────────────────────────────────────────

router.get('/vendors', async (req, res) => {
  try { res.json(await Vendor.find().sort({ name: 1 })); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/vendors', requireOwner, async (req, res) => {
  try { res.status(201).json(await new Vendor(req.body).save()); }
  catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/vendors/:id', requireOwner, async (req, res) => {
  try { res.json(await Vendor.findByIdAndUpdate(req.params.id, req.body, { new: true })); }
  catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/vendors/:id', requireOwner, async (req, res) => {
  try { await Vendor.findByIdAndDelete(req.params.id); res.json({ message: 'Deleted' }); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Master Values ─────────────────────────────────────────────────────────────

router.get('/master-values', async (req, res) => {
  try { res.json(await MasterValue.find().sort({ type: 1, value: 1 })); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/master-values', requireOwner, async (req, res) => {
  try { res.status(201).json(await new MasterValue(req.body).save()); }
  catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/master-values/:id', requireOwner, async (req, res) => {
  try { res.json(await MasterValue.findByIdAndUpdate(req.params.id, req.body, { new: true })); }
  catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/master-values/:id', requireOwner, async (req, res) => {
  try { await MasterValue.findByIdAndDelete(req.params.id); res.json({ message: 'Deleted' }); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Checklists ────────────────────────────────────────────────────────────────
// Both roles can create/view; only owner can delete

router.get('/checklists', async (req, res) => {
  try {
    const { fromDate, toDate, type } = req.query;
    const query = {};
    if (type) query.type = type;
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    res.json(await ChecklistLog.find(query).sort({ date: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/checklists', async (req, res) => {
  try {
    const log = await new ChecklistLog({ ...req.body, createdBy: req.user.role }).save();
    res.status(201).json(log);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/checklists/:id', requireOwner, async (req, res) => {
  try { await ChecklistLog.findByIdAndDelete(req.params.id); res.json({ message: 'Deleted' }); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Inventory ─────────────────────────────────────────────────────────────────

router.get('/inventory', async (req, res) => {
  try { res.json(await ManagementInventory.find().sort({ category: 1, item: 1 })); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/inventory', async (req, res) => {
  try {
    const { item, category, unit, openingStock, usedQty, closingStock, threshold } = req.body;
    let inv = await ManagementInventory.findOne({ item: { $regex: new RegExp(`^${item.trim()}$`, 'i') } });
    if (inv) {
      if (req.user.role === 'staff') {
        // Staff can only update usedQty
        if (usedQty !== undefined) {
          inv.usedQty = usedQty;
          inv.closingStock = (inv.openingStock || 0) + (inv.purchasedQty || 0) - inv.usedQty;
        }
      } else {
        if (category) inv.category = category;
        if (unit) inv.unit = unit;
        if (openingStock !== undefined) inv.openingStock = openingStock;
        if (usedQty !== undefined) inv.usedQty = usedQty;
        if (threshold !== undefined) inv.threshold = threshold;
        inv.closingStock = closingStock !== undefined ? closingStock
          : (inv.openingStock || 0) + (inv.purchasedQty || 0) - (inv.usedQty || 0);
      }
      res.json(await inv.save());
    } else {
      if (req.user.role === 'staff') return res.status(403).json({ message: 'Staff cannot create inventory items' });
      inv = await new ManagementInventory({
        item, category, unit, openingStock, usedQty,
        closingStock: closingStock !== undefined ? closingStock : (openingStock || 0) - (usedQty || 0),
        threshold: threshold || 0,
        createdBy: req.user.role,
      }).save();
      res.status(201).json(inv);
    }
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/inventory/threshold', requireOwner, async (req, res) => {
  try {
    const { item, threshold } = req.body;
    const inv = await ManagementInventory.findOneAndUpdate(
      { item: { $regex: new RegExp(`^${item.trim()}$`, 'i') } },
      { threshold },
      { new: true, upsert: true }
    );
    res.json(inv);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/inventory/:id', requireOwner, async (req, res) => {
  try { await ManagementInventory.findByIdAndDelete(req.params.id); res.json({ message: 'Deleted' }); }
  catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Dashboard Stats ───────────────────────────────────────────────────────────

router.get('/dashboard-stats', requireOwner, async (req, res) => {
  try {
    const salesData = await DailySales.aggregate([{ $group: { _id: null, totalCash: { $sum: '$cash' }, totalUpi: { $sum: '$upi' }, totalSwiggy: { $sum: '$swiggy' }, totalZomato: { $sum: '$zomato' } } }]);
    const offlineSales = salesData.length ? salesData[0].totalCash + salesData[0].totalUpi : 0;
    const onlineSales = salesData.length ? salesData[0].totalSwiggy + salesData[0].totalZomato : 0;
    const totalSales = offlineSales + onlineSales;

    const purchaseData = await PurchaseHeader.aggregate([{ $group: { _id: null, total: { $sum: '$totalAmount' } } }]);
    const totalPurchases = purchaseData.length ? purchaseData[0].total : 0;

    const expenseData = await Expense.aggregate([{ $group: { _id: null, total: { $sum: '$amount' } } }]);
    const totalExpenses = expenseData.length ? expenseData[0].total : 0;

    const settlementData = await OnlineSettlement.aggregate([{ $group: { _id: null, totalCharges: { $sum: '$charges' } } }]);
    const onlineCharges = settlementData.length ? settlementData[0].totalCharges : 0;

    const lowStockItems = await ManagementInventory.find({
      threshold: { $gt: 0 },
      $expr: { $lte: ['$closingStock', '$threshold'] },
    });

    const topVendors = await PurchaseHeader.aggregate([
      { $group: { _id: '$vendor', total: { $sum: '$totalAmount' } } },
      { $sort: { total: -1 } }, { $limit: 8 },
      { $project: { vendor: '$_id', total: 1, _id: 0 } },
    ]);

    const recentSales = await DailySales.find().sort({ date: -1 }).limit(7);

    res.json({
      totalSales, totalPurchases, totalExpenses, onlineCharges,
      offlineSales, onlineSales,
      simplePnL: totalSales - totalPurchases - totalExpenses,
      detailedPnL: totalSales - totalPurchases - totalExpenses - onlineCharges,
      lowStockItems,
      thresholdAlerts: lowStockItems.length,
      topVendors,
      recentSales,
    });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Reports ───────────────────────────────────────────────────────────────────

router.get('/reports', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate } = req.query;
    const start = new Date(fromDate); start.setHours(0, 0, 0, 0);
    const end = new Date(toDate); end.setHours(23, 59, 59, 999);
    const dq = { date: { $gte: start, $lte: end } };
    const dqPay = { paymentDate: { $gte: start, $lte: end } };

    const [salesAgg, purchaseAgg, expenseAgg, settlementAgg] = await Promise.all([
      DailySales.aggregate([{ $match: dq }, { $group: { _id: null, totalSales: { $sum: '$total' }, cashSales: { $sum: '$cash' }, upiSales: { $sum: '$upi' }, swiggySales: { $sum: '$swiggy' }, zomatoSales: { $sum: '$zomato' } } }]),
      PurchaseHeader.aggregate([{ $match: dq }, { $group: { _id: null, totalPurchases: { $sum: '$totalAmount' } } }]),
      Expense.aggregate([{ $match: dq }, { $group: { _id: null, totalExpenses: { $sum: '$amount' } } }]),
      OnlineSettlement.aggregate([{ $match: dqPay }, { $group: { _id: null, totalCharges: { $sum: '$charges' }, totalReceived: { $sum: '$payoutReceived' }, totalGrossSales: { $sum: '$grossSales' } } }]),
    ]);

    res.json({
      sales: salesAgg[0] || { totalSales: 0, cashSales: 0, upiSales: 0, swiggySales: 0, zomatoSales: 0 },
      purchases: purchaseAgg[0] || { totalPurchases: 0 },
      expenses: expenseAgg[0] || { totalExpenses: 0 },
      settlements: settlementAgg[0] || { totalCharges: 0, totalReceived: 0, totalGrossSales: 0 },
    });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Purchase Summary Report ───────────────────────────────────────────────────

router.get('/reports/purchase-summary', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate } = req.query;
    if (!fromDate || !toDate) return res.status(400).json({ message: 'fromDate and toDate are required' });
    const start = new Date(fromDate); start.setHours(0, 0, 0, 0);
    const end = new Date(toDate); end.setHours(23, 59, 59, 999);
    const lines = await PurchaseLine.aggregate([
      { $lookup: { from: 'purchaseheaders', localField: 'purchaseHeader', foreignField: '_id', as: 'header' } },
      { $unwind: '$header' },
      { $match: { 'header.date': { $gte: start, $lte: end } } },
      { $project: { date: '$header.date', item: '$item', quantity: '$quantity', unitPrice: { $cond: [{ $gt: ['$quantity', 0] }, { $divide: ['$rate', '$quantity'] }, 0] } } },
      { $sort: { date: -1, item: 1 } },
    ]);
    const itemPrices = {};
    lines.forEach(l => {
      if (!itemPrices[l.item]) itemPrices[l.item] = new Set();
      itemPrices[l.item].add(Number(l.unitPrice.toFixed(2)));
    });
    const data = lines.map(l => ({ ...l, priceChanged: itemPrices[l.item].size > 1 }));
    res.json({ success: true, data });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Config / Settings ─────────────────────────────────────────────────────────

router.get('/config', requireOwner, async (req, res) => {
  try {
    let config = await Config.findOne();
    if (!config) config = await new Config().save();
    res.json(config);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.put('/config', requireOwner, async (req, res) => {
  try {
    const { cafeName, ownerEmail, ownerPin, staffPin } = req.body;
    let config = await Config.findOne();
    if (!config) config = new Config();
    if (cafeName !== undefined) config.cafeName = cafeName;
    if (ownerEmail !== undefined) config.ownerEmail = ownerEmail;
    if (ownerPin !== undefined && ownerPin.trim()) config.ownerPin = ownerPin;
    if (staffPin !== undefined && staffPin.trim()) config.staffPin = staffPin;
    res.json(await config.save());
  } catch (e) { res.status(400).json({ message: e.message }); }
});

// ── Excel Export ──────────────────────────────────────────────────────────────

router.get('/export-excel', requireOwner, async (req, res) => {
  try {
    const { type, fromDate, toDate, vendor, platform } = req.query;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(type || 'Export');
    const query = {};

    if (fromDate && toDate) {
      const s = new Date(fromDate); const e = new Date(toDate); e.setHours(23, 59, 59, 999);
      if (type === 'online-settlements') query.paymentDate = { $gte: s, $lte: e };
      else if (type !== 'inventory') query.date = { $gte: s, $lte: e };
    }
    if (vendor && type === 'purchase-headers') query.vendor = vendor;
    if (platform && type === 'online-settlements') query.platform = platform;

    let data = [], columns = [];

    switch (type) {
      case 'daily-sales':
        data = await DailySales.find(query).sort({ date: -1 });
        columns = [{ header: 'Date', key: 'date', width: 15 }, { header: 'Cash', key: 'cash', width: 12 }, { header: 'UPI', key: 'upi', width: 12 }, { header: 'Swiggy', key: 'swiggy', width: 12 }, { header: 'Zomato', key: 'zomato', width: 12 }, { header: 'Total', key: 'total', width: 12 }, { header: 'Notes', key: 'notes', width: 30 }];
        break;
      case 'expenses':
        data = await Expense.find(query).sort({ date: -1 });
        columns = [{ header: 'Date', key: 'date', width: 15 }, { header: 'Category', key: 'category', width: 20 }, { header: 'Amount', key: 'amount', width: 12 }, { header: 'Notes', key: 'notes', width: 30 }];
        break;
      case 'purchase-headers':
        data = await PurchaseHeader.find(query).sort({ date: -1 });
        columns = [{ header: 'Date', key: 'date', width: 15 }, { header: 'Bill No', key: 'billNo', width: 15 }, { header: 'Vendor', key: 'vendor', width: 25 }, { header: 'Total Amount', key: 'totalAmount', width: 15 }, { header: 'Payment Method', key: 'paymentMethod', width: 20 }];
        break;
      case 'inventory':
        data = await ManagementInventory.find().sort({ category: 1, item: 1 });
        columns = [{ header: 'Item', key: 'item', width: 25 }, { header: 'Category', key: 'category', width: 20 }, { header: 'Unit', key: 'unit', width: 10 }, { header: 'Opening', key: 'openingStock', width: 10 }, { header: 'Purchased', key: 'purchasedQty', width: 10 }, { header: 'Used', key: 'usedQty', width: 10 }, { header: 'Closing', key: 'closingStock', width: 10 }, { header: 'Threshold', key: 'threshold', width: 10 }];
        break;
      case 'salaries':
        data = await Salary.find(query).sort({ date: -1 });
        columns = [{ header: 'Date', key: 'date', width: 15 }, { header: 'Employee', key: 'employeeName', width: 20 }, { header: 'Amount', key: 'amount', width: 12 }, { header: 'Type', key: 'type', width: 15 }, { header: 'Payment Method', key: 'paymentMethod', width: 15 }, { header: 'Notes', key: 'notes', width: 30 }];
        break;
      case 'purchase-lines':
        data = await PurchaseLine.find(query).populate('purchaseHeader', 'billNo vendor date').sort({ createdAt: -1 });
        columns = [{ header: 'Date', key: 'headerDate', width: 15 }, { header: 'Bill No', key: 'billNo', width: 15 }, { header: 'Vendor', key: 'vendor', width: 25 }, { header: 'Item', key: 'item', width: 30 }, { header: 'Quantity', key: 'quantity', width: 12 }, { header: 'Unit Price', key: 'unitPrice', width: 15 }, { header: 'Total', key: 'rate', width: 15 }];
        break;
      case 'online-settlements':
        data = await OnlineSettlement.find(query).sort({ paymentDate: -1 });
        columns = [{ header: 'Platform', key: 'platform', width: 15 }, { header: 'From', key: 'fromDate', width: 15 }, { header: 'To', key: 'toDate', width: 15 }, { header: 'Payment Date', key: 'paymentDate', width: 15 }, { header: 'Gross Sales', key: 'grossSales', width: 15 }, { header: 'Charges', key: 'charges', width: 12 }, { header: 'Expected (G-C)', key: 'expectedPayout', width: 15 }, { header: 'Received', key: 'payoutReceived', width: 15 }, { header: 'Difference', key: 'difference', width: 15 }, { header: 'Reference', key: 'reference', width: 20 }];
        break;
      case 'purchase-summary': {
        const s2 = new Date(fromDate); s2.setHours(0, 0, 0, 0);
        const e2 = new Date(toDate); e2.setHours(23, 59, 59, 999);
        const summaryLines = await PurchaseLine.aggregate([
          { $lookup: { from: 'purchaseheaders', localField: 'purchaseHeader', foreignField: '_id', as: 'header' } },
          { $unwind: '$header' },
          { $match: { 'header.date': { $gte: s2, $lte: e2 } } },
          { $project: { date: '$header.date', item: '$item', quantity: '$quantity', unitPrice: { $cond: [{ $gt: ['$quantity', 0] }, { $divide: ['$rate', '$quantity'] }, 0] } } },
          { $sort: { date: -1, item: 1 } },
        ]);
        const itemPrices = {};
        summaryLines.forEach(l => { if (!itemPrices[l.item]) itemPrices[l.item] = new Set(); itemPrices[l.item].add(Number(l.unitPrice.toFixed(2))); });
        data = summaryLines.map(l => ({ date: new Date(l.date).toLocaleDateString('en-IN'), item: l.item, quantity: l.quantity, unitPrice: Number(l.unitPrice.toFixed(2)), priceChanged: itemPrices[l.item].size > 1 ? 'Yes' : 'No' }));
        columns = [{ header: 'Date', key: 'date', width: 15 }, { header: 'Item', key: 'item', width: 30 }, { header: 'Quantity', key: 'quantity', width: 12 }, { header: 'Unit Price', key: 'unitPrice', width: 15 }, { header: 'Price Changed?', key: 'priceChanged', width: 15 }];
        ws.columns = columns;
        data.forEach(row => ws.addRow(row));
        ws.getRow(1).font = { bold: true };
        ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=purchase-summary-${new Date().toISOString().split('T')[0]}.xlsx`);
        await wb.xlsx.write(res);
        return res.end();
      }
      case 'pnl-simple':
      case 'pnl-detailed': {
        const s3 = new Date(fromDate); s3.setHours(0, 0, 0, 0);
        const e3 = new Date(toDate); e3.setHours(23, 59, 59, 999);
        const dq3 = { date: { $gte: s3, $lte: e3 } };
        const [sa, pa, ea, sta] = await Promise.all([
          DailySales.aggregate([{ $match: dq3 }, { $group: { _id: null, total: { $sum: '$total' }, cash: { $sum: '$cash' }, upi: { $sum: '$upi' }, swiggy: { $sum: '$swiggy' }, zomato: { $sum: '$zomato' } } }]),
          PurchaseHeader.aggregate([{ $match: dq3 }, { $group: { _id: null, total: { $sum: '$totalAmount' } } }]),
          Expense.aggregate([{ $match: dq3 }, { $group: { _id: null, total: { $sum: '$amount' } } }]),
          OnlineSettlement.aggregate([{ $match: { paymentDate: { $gte: s3, $lte: e3 } } }, { $group: { _id: null, charges: { $sum: '$charges' } } }]),
        ]);
        const sv = sa[0] || { total: 0, cash: 0, upi: 0, swiggy: 0, zomato: 0 };
        const pv = (pa[0] || {}).total || 0;
        const ev = (ea[0] || {}).total || 0;
        const stv = (sta[0] || {}).charges || 0;
        columns = [{ header: 'Metric', key: 'metric', width: 35 }, { header: 'Amount (₹)', key: 'amount', width: 20 }];
        data = type === 'pnl-simple'
          ? [{ metric: 'Total Sales', amount: sv.total }, { metric: 'Total Purchases', amount: pv }, { metric: 'Total Expenses', amount: ev }, { metric: 'Simple Profit / Loss', amount: sv.total - pv - ev }]
          : [{ metric: 'Offline Sales (Cash + UPI)', amount: sv.cash + sv.upi }, { metric: 'Online Sales (Swiggy + Zomato)', amount: sv.swiggy + sv.zomato }, { metric: 'Total Gross Sales', amount: sv.total }, { metric: 'Total Purchases', amount: pv }, { metric: 'Total Expenses', amount: ev }, { metric: 'Platform Charges (Swiggy/Zomato)', amount: stv }, { metric: 'Detailed Profit / Loss', amount: sv.total - pv - ev - stv }];
        ws.columns = columns;
        data.forEach(row => ws.addRow(row));
        ws.getRow(1).font = { bold: true };
        ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=${type}-${new Date().toISOString().split('T')[0]}.xlsx`);
        await wb.xlsx.write(res);
        return res.end();
      }
      default:
        return res.status(400).json({ message: 'Invalid export type' });
    }

    ws.columns = columns;
    data.forEach(item => {
      const row = { ...item.toObject() };
      if (type === 'purchase-lines' && row.purchaseHeader) {
        row.billNo = row.purchaseHeader.billNo;
        row.vendor = row.purchaseHeader.vendor;
        row.headerDate = new Date(row.purchaseHeader.date).toLocaleDateString('en-IN');
      }
      if (type === 'online-settlements') {
        row.expectedPayout = (row.grossSales || 0) - (row.charges || 0);
      }
      ['date', 'paymentDate', 'fromDate', 'toDate'].forEach(k => { if (row[k]) row[k] = new Date(row[k]).toLocaleDateString('en-IN'); });
      ws.addRow(row);
    });
    ws.getRow(1).font = { bold: true };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${type}-${new Date().toISOString().split('T')[0]}.xlsx`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Invoice Parsing (Excel) ───────────────────────────────────────────────────

router.post('/parse-invoice', requireOwner, async (req, res) => {
  try {
    const { fileData } = req.body;
    if (!fileData) return res.status(400).json({ message: 'No file data' });
    const buffer = Buffer.from(fileData.split(',')[1] || fileData, 'base64');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet(1);
    const items = [];
    ws.eachRow((row, n) => {
      if (n > 1) {
        const item = row.getCell(1).value;
        const qty = Number(row.getCell(2).value);
        const rate = Number(row.getCell(3).value);
        if (item && !isNaN(qty) && !isNaN(rate)) {
          items.push({ item: String(item).trim(), quantity: qty, rate, unitPrice: qty > 0 ? rate / qty : 0, total: rate });
        }
      }
    });
    res.json({ success: true, items });
  } catch (e) { res.status(500).json({ message: 'Failed to parse: ' + e.message }); }
});

router.post('/confirm-invoice', requireOwner, async (req, res) => {
  try {
    const { purchaseHeaderId, items } = req.body;
    if (!purchaseHeaderId || !Array.isArray(items)) return res.status(400).json({ message: 'Invalid data' });
    const header = await PurchaseHeader.findById(purchaseHeaderId);
    if (!header) return res.status(404).json({ message: 'Header not found' });
    const linesTotal = items.reduce((s, i) => s + (Number(i.rate) || 0), 0);
    if (Math.abs(header.totalAmount - linesTotal) > 0.01) {
      return res.status(400).json({ message: `Bill amount (${header.totalAmount}) ≠ items total (${linesTotal.toFixed(2)})` });
    }
    const created = [];
    for (const d of items) {
      const line = await new PurchaseLine({ purchaseHeader: purchaseHeaderId, ...d, createdBy: req.user.role }).save();
      await updateInventoryStock(d.item, d.quantity, req.user.role);
      created.push(line);
    }
    res.json({ success: true, message: `${created.length} items saved`, data: created });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Inventory Bulk Upload ─────────────────────────────────────────────────────

router.post('/inventory/bulk-upload', requireOwner, async (req, res) => {
  try {
    const { fileData } = req.body;
    if (!fileData) return res.status(400).json({ message: 'No file data' });
    const buffer = Buffer.from(fileData.split(',')[1] || fileData, 'base64');
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);
    const ws = wb.getWorksheet(1);
    const allowed = ['Food Raw Material', 'Vegetables', 'Flour/Other', 'Packaging', 'Other'];
    const existing = await ManagementInventory.find({});
    const normed = existing.map(i => normalizeItemName(i.item));
    const toAdd = [], skipped = [];
    ws.eachRow((row, n) => {
      if (n > 1) {
        const name = row.getCell(1).value;
        if (name) {
          const nm = String(name).trim();
          const cat = String(row.getCell(2).value || '').trim();
          if (!normed.includes(normalizeItemName(nm))) {
            toAdd.push({ item: nm, category: allowed.includes(cat) ? cat : 'Other', unit: String(row.getCell(3).value || 'Pkt').trim(), threshold: Number(row.getCell(4).value) || 0, createdBy: req.user.role });
            normed.push(normalizeItemName(nm));
          } else skipped.push(nm);
        }
      }
    });
    if (toAdd.length) await ManagementInventory.insertMany(toAdd);
    res.json({ success: true, message: `${toAdd.length} added, ${skipped.length} skipped`, addedCount: toAdd.length, skippedCount: skipped.length });
  } catch (e) { res.status(500).json({ message: 'Bulk upload failed: ' + e.message }); }
});

module.exports = router;
