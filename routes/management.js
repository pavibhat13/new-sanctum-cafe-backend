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
const StaffLeave = require('../models/StaffLeave');
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
    const { fromDate, toDate, vendor, source } = req.query;
    const query = {};
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    if (vendor) query.vendor = vendor;
    if (source) query.source = source;
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
    const { purchaseHeader, item, quantity, rate, note } = req.body;
    if (purchaseHeader !== undefined) line.purchaseHeader = purchaseHeader;
    if (item !== undefined) line.item = item;
    if (quantity !== undefined) line.quantity = quantity;
    if (rate !== undefined) line.rate = rate;
    if (note !== undefined) line.note = note;
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

router.put('/expenses/:id', async (req, res) => {
  try {
    const expense = await Expense.findById(req.params.id);
    if (!expense) return res.status(404).json({ message: 'Not found' });
    // Staff can only edit expenses they created
    if (req.user.role === 'staff' && expense.createdBy !== 'staff') {
      return res.status(403).json({ message: 'Not authorised to edit this expense' });
    }
    Object.assign(expense, req.body);
    res.json(await expense.save());
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/expenses/:id', async (req, res) => {
  try {
    const expense = await Expense.findById(req.params.id);
    if (!expense) return res.status(404).json({ message: 'Not found' });
    // Staff can only delete expenses they created
    if (req.user.role === 'staff' && expense.createdBy !== 'staff') {
      return res.status(403).json({ message: 'Not authorised to delete this expense' });
    }
    await expense.deleteOne();
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
    const { item, category, unit, openingStock, usedQty, closingStock, threshold, subCategory, physicalCount } = req.body;
    let inv = await ManagementInventory.findOne({ item: { $regex: new RegExp(`^${item.trim()}$`, 'i') } });
    if (inv) {
      if (req.user.role === 'staff') {
        if (physicalCount !== undefined) {
          const count = Number(physicalCount);
          const systemTotal = (inv.openingStock || 0) + (inv.purchasedQty || 0);
          if (count > systemTotal) {
            return res.status(400).json({
              message: `Count (${count}) exceeds system stock (${systemTotal} ${inv.unit || ''}). Record a purchase first.`,
            });
          }
          inv.usedQty = systemTotal - count;
          inv.closingStock = count;
        } else {
          // Legacy path (backward compat)
          if (openingStock !== undefined) inv.openingStock = openingStock;
          if (usedQty !== undefined) inv.usedQty = usedQty;
          inv.closingStock = (inv.openingStock || 0) + (inv.purchasedQty || 0) - (inv.usedQty || 0);
        }
      } else {
        if (category) inv.category = category;
        if (unit) inv.unit = unit;
        if (subCategory !== undefined) inv.subCategory = subCategory;
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
        subCategory: subCategory || '',
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

// ── Staff Quick Purchase ──────────────────────────────────────────────────────
// Both staff and owner can submit; creates a Purchase Header + Line atomically

router.post('/staff-purchase', async (req, res) => {
  try {
    const { item, quantity, amount, vendor, date, note } = req.body;
    if (!item || !item.trim()) return res.status(400).json({ message: 'Item is required' });
    if (!quantity || Number(quantity) <= 0) return res.status(400).json({ message: 'Quantity must be greater than 0' });
    if (amount === undefined || Number(amount) < 0) return res.status(400).json({ message: 'Amount is required' });

    const purchaseDate = date ? new Date(date) : new Date();
    const dateStr = purchaseDate.toISOString().slice(0, 10).replace(/-/g, '');
    const billNo = `STAFF-${dateStr}-${Date.now().toString().slice(-6)}`;

    const header = await new PurchaseHeader({
      billNo,
      vendor: (vendor && vendor.trim()) || 'Staff Purchase',
      date: purchaseDate,
      totalAmount: Number(amount),
      paymentMethod: 'Cash',
      notes: note || '',
      source: 'staff',
      reviewed: false,
      createdBy: req.user.role,
    }).save();

    const line = await new PurchaseLine({
      purchaseHeader: header._id,
      item: item.trim(),
      quantity: Number(quantity),
      rate: Number(amount),
      note: note || '',
      createdBy: req.user.role,
    }).save();

    await updateInventoryStock(line.item, line.quantity, req.user.role);

    res.status(201).json({ header, line });
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.put('/purchase-headers/:id/review', requireOwner, async (req, res) => {
  try {
    const header = await PurchaseHeader.findByIdAndUpdate(
      req.params.id, { reviewed: true }, { new: true }
    );
    if (!header) return res.status(404).json({ message: 'Not found' });
    res.json(header);
  } catch (e) { res.status(400).json({ message: e.message }); }
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

    const salaryData = await Salary.aggregate([{ $group: { _id: null, total: { $sum: '$amount' } } }]);
    const totalSalaries = salaryData.length ? salaryData[0].total : 0;

    const settlementData = await OnlineSettlement.aggregate([{ $group: { _id: null, totalCharges: { $sum: '$charges' } } }]);
    const onlineCharges = settlementData.length ? settlementData[0].totalCharges : 0;

    const lowStockItems = await ManagementInventory.find({
      threshold: { $gt: 0 },
      $expr: { $lte: ['$closingStock', '$threshold'] },
    });

    const unreviewedStaffPurchases = await PurchaseHeader.countDocuments({ source: 'staff', reviewed: false });

    const topVendors = await PurchaseHeader.aggregate([
      { $group: { _id: '$vendor', total: { $sum: '$totalAmount' } } },
      { $sort: { total: -1 } }, { $limit: 8 },
      { $project: { vendor: '$_id', total: 1, _id: 0 } },
    ]);

    const recentSales = await DailySales.find().sort({ date: -1 }).limit(7);

    res.json({
      totalSales, totalPurchases, totalExpenses, totalSalaries, onlineCharges,
      offlineSales, onlineSales,
      simplePnL: totalSales - totalPurchases - totalExpenses - totalSalaries,
      detailedPnL: totalSales - totalPurchases - totalExpenses - totalSalaries - onlineCharges,
      lowStockItems,
      thresholdAlerts: lowStockItems.length,
      unreviewedStaffPurchases,
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

    const [salesAgg, purchaseAgg, expenseAgg, settlementAgg, salaryAgg] = await Promise.all([
      DailySales.aggregate([{ $match: dq }, { $group: { _id: null, totalSales: { $sum: '$total' }, cashSales: { $sum: '$cash' }, upiSales: { $sum: '$upi' }, swiggySales: { $sum: '$swiggy' }, zomatoSales: { $sum: '$zomato' } } }]),
      PurchaseHeader.aggregate([{ $match: dq }, { $group: { _id: null, totalPurchases: { $sum: '$totalAmount' } } }]),
      Expense.aggregate([{ $match: dq }, { $group: { _id: null, totalExpenses: { $sum: '$amount' } } }]),
      OnlineSettlement.aggregate([{ $match: dqPay }, { $group: { _id: null, totalCharges: { $sum: '$charges' }, totalReceived: { $sum: '$payoutReceived' }, totalGrossSales: { $sum: '$grossSales' } } }]),
      Salary.aggregate([{ $match: dq }, { $group: { _id: null, totalSalaries: { $sum: '$amount' } } }]),
    ]);

    res.json({
      sales: salesAgg[0] || { totalSales: 0, cashSales: 0, upiSales: 0, swiggySales: 0, zomatoSales: 0 },
      purchases: purchaseAgg[0] || { totalPurchases: 0 },
      expenses: expenseAgg[0] || { totalExpenses: 0 },
      settlements: settlementAgg[0] || { totalCharges: 0, totalReceived: 0, totalGrossSales: 0 },
      salaries: salaryAgg[0] || { totalSalaries: 0 },
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

// ── Inventory Snapshot Report ─────────────────────────────────────────────────

router.get('/reports/inventory-snapshot', requireOwner, async (req, res) => {
  try {
    const items = await ManagementInventory.find().sort({ category: 1, item: 1 });
    res.json(items.map(i => ({
      _id: i._id,
      item: i.item,
      category: i.category,
      unit: i.unit,
      openingStock: i.openingStock,
      purchasedQty: i.purchasedQty,
      usedQty: i.usedQty,
      closingStock: i.closingStock,
      threshold: i.threshold,
      status: i.threshold > 0 && i.closingStock <= i.threshold ? 'Low' : 'OK',
    })));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Inventory Movement Report ─────────────────────────────────────────────────

router.get('/reports/inventory-movement', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate } = req.query;
    if (!fromDate || !toDate) return res.status(400).json({ message: 'fromDate and toDate required' });
    const start = new Date(fromDate); start.setHours(0, 0, 0, 0);
    const end = new Date(toDate); end.setHours(23, 59, 59, 999);

    const lines = await PurchaseLine.aggregate([
      { $lookup: { from: 'purchaseheaders', localField: 'purchaseHeader', foreignField: '_id', as: 'header' } },
      { $unwind: '$header' },
      { $match: { 'header.date': { $gte: start, $lte: end } } },
      { $group: {
        _id: '$item',
        timesOrdered: { $sum: 1 },
        totalQtyPurchased: { $sum: '$quantity' },
        totalSpend: { $sum: '$rate' },
      }},
      { $sort: { timesOrdered: -1 } },
    ]);

    const inventory = await ManagementInventory.find({});
    const invMap = {};
    inventory.forEach(i => { invMap[normalizeItemName(i.item)] = i.usedQty || 0; });

    const total = lines.length;
    const fastCut = Math.ceil(total / 3);
    const slowCut = total - Math.floor(total / 3);

    const data = lines.map((l, i) => ({
      item: l._id,
      timesOrdered: l.timesOrdered,
      totalQtyPurchased: Number(l.totalQtyPurchased.toFixed(2)),
      totalSpend: Number(l.totalSpend.toFixed(2)),
      usedQty: invMap[normalizeItemName(l._id)] || 0,
      movement: total < 3 ? 'Normal' : i < fastCut ? 'Fast' : i >= slowCut ? 'Slow' : 'Normal',
    }));

    res.json(data);
  } catch (e) { res.status(500).json({ message: e.message }); }
});

// ── Vendor Frequency Report ───────────────────────────────────────────────────

router.get('/reports/vendor-frequency', requireOwner, async (req, res) => {
  try {
    const { fromDate, toDate } = req.query;
    if (!fromDate || !toDate) return res.status(400).json({ message: 'fromDate and toDate required' });
    const start = new Date(fromDate); start.setHours(0, 0, 0, 0);
    const end = new Date(toDate); end.setHours(23, 59, 59, 999);

    const data = await PurchaseHeader.aggregate([
      { $match: { date: { $gte: start, $lte: end } } },
      { $group: {
        _id: '$vendor',
        invoiceCount: { $sum: 1 },
        totalAmount: { $sum: '$totalAmount' },
        lastInvoiceDate: { $max: '$date' },
      }},
      { $sort: { invoiceCount: -1 } },
      { $project: { vendor: '$_id', invoiceCount: 1, totalAmount: 1, lastInvoiceDate: 1, _id: 0 } },
    ]);

    res.json(data);
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

router.get('/download-template', requireOwner, async (req, res) => {
  try {
    const { type } = req.query;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Template');

    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD4A017' } };
    const headerFont = { bold: true, color: { argb: 'FF1A1815' }, size: 11 };
    const noteFill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEEEEEE' } };
    const noteFont  = { italic: true, color: { argb: 'FF888888' }, size: 9 };

    if (type === 'invoice') {
      ws.columns = [
        { header: 'Item', key: 'item', width: 30 },
        { header: 'Qty', key: 'qty', width: 12 },
        { header: 'Total Amount (₹)', key: 'total', width: 20 },
        { header: 'Note', key: 'note', width: 28 },
      ];
      ws.getRow(1).eachCell(c => { c.fill = headerFill; c.font = headerFont; });
      const samples = [['Onion', 5.5, 100, ''], ['Tomato', 3, 60, ''], ['Mozzarella Cheese', 2, 500, 'Exp: Dec 2026']];
      samples.forEach(r => ws.addRow(r));
      const noteRow = ws.addRow(['← Item name (text)', '← Quantity', '← Total cost for this line (not unit price). Grand total must match bill amount.', '← Optional note (expiry, batch no…)']);
      noteRow.eachCell(c => { c.fill = noteFill; c.font = noteFont; });

    } else if (type === 'inventory-master') {
      ws.columns = [
        { header: 'Item Name', key: 'item', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Unit', key: 'unit', width: 12 },
        { header: 'Low Stock Threshold', key: 'threshold', width: 22 },
      ];
      ws.getRow(1).eachCell(c => { c.fill = headerFill; c.font = headerFont; });
      [
        ['Mozzarella Cheese', 'Food Raw Material', 'Pkt', 5],
        ['Onion', 'Vegetables', 'Kg', 2],
        ['Maida', 'Flour/Other', 'Bag', 1],
        ['Pizza Box 7"', 'Packaging', 'Pkt', 10],
      ].forEach(r => ws.addRow(r));
      const noteRow = ws.addRow(['← Item name', '← Food Raw Material / Vegetables / Flour/Other / Packaging / Other', '← Pkt, Kg, Bag, Ltr…', '← Alert when closing stock ≤ this']);
      noteRow.eachCell(c => { c.fill = noteFill; c.font = noteFont; });

    } else if (type === 'inventory-stock') {
      ws.columns = [
        { header: 'Item Name', key: 'item', width: 30 },
        { header: 'Category', key: 'category', width: 25 },
        { header: 'Unit', key: 'unit', width: 12 },
        { header: 'Opening Stock', key: 'openingStock', width: 16 },
        { header: 'Used Qty', key: 'usedQty', width: 14 },
        { header: 'Low Stock Threshold', key: 'threshold', width: 22 },
      ];
      ws.getRow(1).eachCell(c => { c.fill = headerFill; c.font = headerFont; });
      [
        ['Mozzarella Cheese', 'Food Raw Material', 'Pkt', 10, 3, 5],
        ['Onion', 'Vegetables', 'Kg', 5, 2, 2],
        ['Maida', 'Flour/Other', 'Bag', 2, 1, 1],
        ['Pizza Box 7"', 'Packaging', 'Pkt', 50, 20, 10],
      ].forEach(r => ws.addRow(r));
      const noteRow = ws.addRow(['← Item name', '← Food Raw Material / Vegetables / Flour/Other / Packaging / Other', '← Pkt, Kg, Bag, Ltr…', '← Stock at start of period', '← How much was used', '← Alert threshold']);
      noteRow.eachCell(c => { c.fill = noteFill; c.font = noteFont; });
    } else {
      return res.status(400).json({ message: 'Unknown template type' });
    }

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${type}-template.xlsx"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.get('/export-excel', requireOwner, async (req, res) => {
  try {
    const { type, fromDate, toDate, vendor, platform, billNo } = req.query;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet(type || 'Export');
    const query = {};

    if (fromDate && toDate) {
      const s = new Date(fromDate); const e = new Date(toDate); e.setHours(23, 59, 59, 999);
      if (type === 'online-settlements') query.paymentDate = { $gte: s, $lte: e };
      else if (type !== 'inventory' && type !== 'purchase-lines') query.date = { $gte: s, $lte: e };
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
      case 'purchase-lines': {
        const headerQuery = {};
        if (fromDate && toDate) {
          const s = new Date(fromDate); const e = new Date(toDate); e.setHours(23, 59, 59, 999);
          headerQuery.date = { $gte: s, $lte: e };
        }
        if (billNo) headerQuery.billNo = { $regex: billNo, $options: 'i' };
        const matchingHeaders = await PurchaseHeader.find(headerQuery).select('_id');
        const lineQuery = matchingHeaders.length > 0 || fromDate || billNo
          ? { purchaseHeader: { $in: matchingHeaders.map(h => h._id) } }
          : {};
        data = await PurchaseLine.find(lineQuery).populate('purchaseHeader', 'billNo vendor date').sort({ createdAt: -1 });
        columns = [{ header: 'Date', key: 'headerDate', width: 15 }, { header: 'Bill No', key: 'billNo', width: 15 }, { header: 'Vendor', key: 'vendor', width: 25 }, { header: 'Item', key: 'item', width: 30 }, { header: 'Quantity', key: 'quantity', width: 12 }, { header: 'Unit Price', key: 'unitPrice', width: 15 }, { header: 'Total', key: 'rate', width: 15 }];
        break;
      }
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
      case 'vendor-frequency': {
        const svf = fromDate ? new Date(fromDate) : new Date(0); svf.setHours(0, 0, 0, 0);
        const evf = toDate ? new Date(toDate) : new Date(); evf.setHours(23, 59, 59, 999);
        const vfData = await PurchaseHeader.aggregate([
          { $match: { date: { $gte: svf, $lte: evf } } },
          { $group: { _id: '$vendor', invoiceCount: { $sum: 1 }, totalAmount: { $sum: '$totalAmount' }, lastInvoiceDate: { $max: '$date' } } },
          { $sort: { invoiceCount: -1 } },
          { $project: { vendor: '$_id', invoiceCount: 1, totalAmount: 1, lastInvoiceDate: 1, _id: 0 } },
        ]);
        const vfRows = vfData.map(r => ({
          vendor: r.vendor,
          invoiceCount: r.invoiceCount,
          totalAmount: r.totalAmount,
          lastInvoiceDate: new Date(r.lastInvoiceDate).toLocaleDateString('en-IN'),
        }));
        ws.columns = [
          { header: 'Vendor', key: 'vendor', width: 30 },
          { header: 'Invoice Count', key: 'invoiceCount', width: 15 },
          { header: 'Total Amount (₹)', key: 'totalAmount', width: 20 },
          { header: 'Last Invoice Date', key: 'lastInvoiceDate', width: 20 },
        ];
        vfRows.forEach(row => ws.addRow(row));
        ws.getRow(1).font = { bold: true };
        ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=vendor-frequency-${new Date().toISOString().split('T')[0]}.xlsx`);
        await wb.xlsx.write(res);
        return res.end();
      }
      case 'inventory-movement': {
        const sim = fromDate ? new Date(fromDate) : new Date(0); sim.setHours(0, 0, 0, 0);
        const eim = toDate ? new Date(toDate) : new Date(); eim.setHours(23, 59, 59, 999);
        const imLines = await PurchaseLine.aggregate([
          { $lookup: { from: 'purchaseheaders', localField: 'purchaseHeader', foreignField: '_id', as: 'header' } },
          { $unwind: '$header' },
          { $match: { 'header.date': { $gte: sim, $lte: eim } } },
          { $group: { _id: '$item', timesOrdered: { $sum: 1 }, totalQtyPurchased: { $sum: '$quantity' }, totalSpend: { $sum: '$rate' } } },
          { $sort: { timesOrdered: -1 } },
        ]);
        const imInventory = await ManagementInventory.find({});
        const imMap = {};
        imInventory.forEach(i => { imMap[normalizeItemName(i.item)] = i.usedQty || 0; });
        const imTotal = imLines.length;
        const imFastCut = Math.ceil(imTotal / 3);
        const imSlowCut = imTotal - Math.floor(imTotal / 3);
        const imRows = imLines.map((l, i) => ({
          item: l._id,
          timesOrdered: l.timesOrdered,
          totalQtyPurchased: Number(l.totalQtyPurchased.toFixed(2)),
          totalSpend: Number(l.totalSpend.toFixed(2)),
          usedQty: imMap[normalizeItemName(l._id)] || 0,
          movement: imTotal < 3 ? 'Normal' : i < imFastCut ? 'Fast' : i >= imSlowCut ? 'Slow' : 'Normal',
        }));
        ws.columns = [
          { header: 'Item', key: 'item', width: 30 },
          { header: 'Times Ordered', key: 'timesOrdered', width: 15 },
          { header: 'Total Qty Purchased', key: 'totalQtyPurchased', width: 20 },
          { header: 'Total Spend (₹)', key: 'totalSpend', width: 15 },
          { header: 'Used Qty', key: 'usedQty', width: 12 },
          { header: 'Movement', key: 'movement', width: 12 },
        ];
        imRows.forEach(row => ws.addRow(row));
        ws.getRow(1).font = { bold: true };
        ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename=inventory-movement-${new Date().toISOString().split('T')[0]}.xlsx`);
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
        const note = row.getCell(4).value ? String(row.getCell(4).value).trim() : '';
        if (item && !isNaN(qty) && !isNaN(rate)) {
          items.push({ item: String(item).trim(), quantity: qty, rate, unitPrice: qty > 0 ? rate / qty : 0, total: rate, note });
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
    const existingMap = {};
    existing.forEach(i => { existingMap[normalizeItemName(i.item)] = i; });
    const toAdd = [], toUpdate = [];
    ws.eachRow((row, n) => {
      if (n > 1) {
        const name = row.getCell(1).value;
        if (name) {
          const nm = String(name).trim();
          const cat = String(row.getCell(2).value || '').trim();
          const unit = String(row.getCell(3).value || 'Pkt').trim();
          const col4 = row.getCell(4).value;
          const col5 = row.getCell(5).value;
          const col6 = row.getCell(6).value;
          // Detect inventory-stock template (col5 present) vs inventory-master (col4=Threshold only)
          const hasStockCols = col5 !== null && col5 !== undefined && col5 !== '';
          const openingStock = hasStockCols ? Number(col4) || 0 : 0;
          const usedQty      = hasStockCols ? Number(col5) || 0 : 0;
          const threshold    = hasStockCols ? Number(col6) || 0 : Number(col4) || 0;
          const normKey = normalizeItemName(nm);
          if (existingMap[normKey]) {
            toUpdate.push({ doc: existingMap[normKey], cat, unit, openingStock, usedQty, threshold, hasStockCols });
          } else {
            const closingStock = openingStock - usedQty;
            toAdd.push({ item: nm, category: allowed.includes(cat) ? cat : 'Other', unit, openingStock, usedQty, closingStock, threshold, createdBy: req.user.role });
            existingMap[normKey] = true;
          }
        }
      }
    });
    if (toAdd.length) await ManagementInventory.insertMany(toAdd);
    for (const { doc, cat, unit, openingStock, usedQty, threshold, hasStockCols } of toUpdate) {
      if (cat && allowed.includes(cat)) doc.category = cat;
      if (unit) doc.unit = unit;
      if (threshold) doc.threshold = threshold;
      if (hasStockCols) {
        doc.openingStock = openingStock;
        doc.usedQty = usedQty;
        doc.closingStock = openingStock + (doc.purchasedQty || 0) - usedQty;
      }
      await doc.save();
    }
    res.json({ success: true, message: `${toAdd.length} added, ${toUpdate.length} updated`, addedCount: toAdd.length, updatedCount: toUpdate.length });
  } catch (e) { res.status(500).json({ message: 'Bulk upload failed: ' + e.message }); }
});

// ── Staff Leaves ──────────────────────────────────────────────────────────────

router.get('/staff-leaves', async (req, res) => {
  try {
    const { fromDate, toDate, employeeName } = req.query;
    const query = {};
    const dq = getDateRangeQuery(fromDate, toDate);
    if (dq) query.date = dq;
    if (employeeName) query.employeeName = new RegExp(employeeName.trim(), 'i');
    res.json(await StaffLeave.find(query).sort({ date: -1 }));
  } catch (e) { res.status(500).json({ message: e.message }); }
});

router.post('/staff-leaves', requireOwner, async (req, res) => {
  try {
    const { employeeName, date, leaveType, notes } = req.body;
    const leave = await new StaffLeave({ employeeName, date, leaveType, notes, createdBy: req.user.role }).save();
    res.status(201).json(leave);
  } catch (e) { res.status(400).json({ message: e.message }); }
});

router.delete('/staff-leaves/:id', requireOwner, async (req, res) => {
  try {
    await StaffLeave.findByIdAndDelete(req.params.id);
    res.json({ message: 'Deleted' });
  } catch (e) { res.status(500).json({ message: e.message }); }
});

module.exports = router;
