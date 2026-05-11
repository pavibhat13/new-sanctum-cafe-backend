const mongoose = require('mongoose');

const expenseSchema = new mongoose.Schema({
  date: { type: Date, required: true, default: Date.now },
  category: { type: String, required: true },
  amount: { type: Number, required: true, min: 0 },
  notes: { type: String, trim: true, maxlength: 500 },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true });

expenseSchema.index({ date: -1 });

module.exports = mongoose.model('Expense', expenseSchema);
