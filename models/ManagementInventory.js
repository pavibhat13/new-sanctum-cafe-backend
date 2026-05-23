const mongoose = require('mongoose');

// Stock quantities (openingStock, purchasedQty, usedQty, closingStock) live in
// InventoryPeriodItem — this model is the item master only.
const managementInventorySchema = new mongoose.Schema({
  item: { type: String, required: true, unique: true, trim: true },
  category: {
    // Keep in sync with frontend/src/utils/constants.js INVENTORY_CATEGORIES
    type: String,
    required: true,
    enum: ['Food Raw Material', 'Vegetables', 'Flour/Other', 'Packaging', 'Other'],
    default: 'Other',
  },
  unit:        { type: String, required: true, default: 'Pkt' },
  threshold:   { type: Number, default: 0 },
  subCategory: { type: String, default: '', trim: true },
  createdBy:   { type: String, default: 'owner' },
}, { timestamps: true });

module.exports = mongoose.model('ManagementInventory', managementInventorySchema);
