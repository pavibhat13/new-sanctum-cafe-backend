const mongoose = require('mongoose');

const inventoryPeriodItemSchema = new mongoose.Schema({
  periodId:     { type: mongoose.Schema.Types.ObjectId, ref: 'InventoryPeriod', required: true },
  item:         { type: String, required: true },
  openingStock: { type: Number, default: 0 },
  purchasedQty: { type: Number, default: 0 },
  usedQty:      { type: Number, default: 0 },
  closingStock: { type: Number, default: 0 },
}, { timestamps: true });

inventoryPeriodItemSchema.index({ periodId: 1, item: 1 }, { unique: true });

module.exports = mongoose.model('InventoryPeriodItem', inventoryPeriodItemSchema);
