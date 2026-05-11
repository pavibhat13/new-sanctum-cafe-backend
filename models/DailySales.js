const mongoose = require('mongoose');

const dailySalesSchema = new mongoose.Schema({
  date: { type: Date, required: true, unique: true },
  cash: { type: Number, default: 0, min: 0 },
  upi: { type: Number, default: 0, min: 0 },
  swiggy: { type: Number, default: 0, min: 0 },
  zomato: { type: Number, default: 0, min: 0 },
  total: { type: Number },
  notes: { type: String, maxlength: 500 },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true });

dailySalesSchema.pre('save', function (next) {
  this.total = this.cash + this.upi + this.swiggy + this.zomato;
  next();
});

dailySalesSchema.index({ date: -1 });

module.exports = mongoose.model('DailySales', dailySalesSchema);
