const mongoose = require('mongoose');

const salarySchema = new mongoose.Schema({
  type:          { type: String, enum: ['Salary', 'Advance', 'Bonus'], required: true, default: 'Salary' },
  employeeName:  { type: String, required: true },
  fromDate:      { type: Date, required: true },
  toDate:        { type: Date, required: true },

  // Advance / Bonus only
  amount:        { type: Number },

  // Salary only
  monthlySalary: { type: Number },
  totalDays:     { type: Number },
  dailyRate:     { type: Number },
  allowedLeaves: { type: Number, default: 2 },
  leavesTaken:   { type: Number, default: 0 },
  paidDays:      { type: Number },
  grossPay:      { type: Number },
  advance:       { type: Number, default: 0 },
  netPay:        { type: Number },

  paymentMethod: { type: String, default: 'Cash' },
  notes:         { type: String, trim: true },
  createdBy:     { type: String, default: 'owner' },
}, { timestamps: true });

module.exports = mongoose.model('Salary', salarySchema);
