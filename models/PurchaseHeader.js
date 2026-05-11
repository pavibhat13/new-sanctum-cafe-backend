const mongoose = require('mongoose');

const purchaseHeaderSchema = new mongoose.Schema({
  billNo: { type: String, required: true, trim: true, unique: true },
  vendor: { type: String, required: true, trim: true },
  date: { type: Date, required: true, default: Date.now },
  totalAmount: { type: Number, required: true, min: 0 },
  paymentMethod: {
    type: String,
    required: true,
    enum: ['Cash', 'UPI', 'Bank Transfer', 'Credit', 'Other'],
    default: 'Cash',
  },
  notes: { type: String, maxlength: 500 },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true, toJSON: { virtuals: true }, toObject: { virtuals: true } });

purchaseHeaderSchema.virtual('lines', {
  ref: 'PurchaseLine',
  localField: '_id',
  foreignField: 'purchaseHeader',
});

purchaseHeaderSchema.index({ billNo: 1 });
purchaseHeaderSchema.index({ vendor: 1 });
purchaseHeaderSchema.index({ date: -1 });

module.exports = mongoose.model('PurchaseHeader', purchaseHeaderSchema);
