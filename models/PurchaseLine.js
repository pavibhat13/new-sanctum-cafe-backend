const mongoose = require('mongoose');

const purchaseLineSchema = new mongoose.Schema({
  purchaseHeader: { type: mongoose.Schema.Types.ObjectId, ref: 'PurchaseHeader', required: true },
  item: { type: String, required: true, trim: true },
  quantity: { type: Number, required: true, min: 0.01 },
  rate: { type: Number, required: true, min: 0 },
  unitPrice: { type: Number },
  total: { type: Number },
  note: { type: String, default: '', trim: true },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true });

purchaseLineSchema.pre('save', function (next) {
  this.total = this.rate;
  this.unitPrice = this.quantity > 0 ? this.rate / this.quantity : 0;
  next();
});

purchaseLineSchema.index({ purchaseHeader: 1 });
purchaseLineSchema.index({ item: 1 });

module.exports = mongoose.model('PurchaseLine', purchaseLineSchema);
