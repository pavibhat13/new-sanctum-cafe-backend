const mongoose = require('mongoose');

const inventoryPeriodSchema = new mongoose.Schema({
  periodStart: { type: Date, required: true },
  periodEnd:   { type: Date, default: null },
  status:      { type: String, enum: ['open', 'closed'], default: 'open' },
  label:       { type: String, default: '' },
  closedAt:    { type: Date },
  closedBy:    { type: String },
}, { timestamps: true });

module.exports = mongoose.model('InventoryPeriod', inventoryPeriodSchema);
