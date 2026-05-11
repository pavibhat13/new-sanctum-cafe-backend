const mongoose = require('mongoose');

const salarySchema = new mongoose.Schema({
  employeeName: { type: String, required: true },
  date: { type: Date, required: true, default: Date.now },
  amount: { type: Number, required: true, min: 0 },
  type: { type: String, enum: ['Regular', 'Advance', 'Bonus'], default: 'Regular', required: true },
  paymentMethod: { type: String, required: true },
  notes: { type: String, trim: true },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true });

module.exports = mongoose.model('Salary', salarySchema);
