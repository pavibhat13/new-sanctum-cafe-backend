const mongoose = require('mongoose');

const onlineSettlementSchema = new mongoose.Schema({
  platform: { type: String, required: true, enum: ['Swiggy', 'Zomato'] },
  fromDate: { type: Date, required: true },
  toDate: { type: Date, required: true },
  paymentDate: { type: Date, required: true },
  grossSales: { type: Number, default: 0 },
  charges: { type: Number, default: 0 },
  payoutReceived: { type: Number, default: 0 },
  difference: { type: Number, default: 0 },
  reference: { type: String, trim: true },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true });

onlineSettlementSchema.pre('save', function (next) {
  this.difference = (this.grossSales - this.charges) - this.payoutReceived;
  next();
});

module.exports = mongoose.model('OnlineSettlement', onlineSettlementSchema);
