const mongoose = require('mongoose');

const managementInventorySchema = new mongoose.Schema({
  item: { type: String, required: true, unique: true, trim: true },
  category: {
    // Keep in sync with frontend/src/utils/constants.js INVENTORY_CATEGORIES
    type: String,
    required: true,
    enum: ['Food Raw Material', 'Vegetables', 'Flour/Other', 'Packaging', 'Other'],
    default: 'Other',
  },
  unit: { type: String, required: true, default: 'Pkt' },
  openingStock: { type: Number, default: 0 },
  purchasedQty: { type: Number, default: 0 },
  usedQty: { type: Number, default: 0 },
  closingStock: { type: Number, default: 0 },
  threshold: { type: Number, default: 0 },
  subCategory: { type: String, default: '', trim: true },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true, toJSON: { virtuals: true }, toObject: { virtuals: true } });

managementInventorySchema.virtual('status').get(function () {
  return this.threshold > 0 && this.closingStock <= this.threshold ? 'Low Stock' : 'Normal';
});

module.exports = mongoose.model('ManagementInventory', managementInventorySchema);
