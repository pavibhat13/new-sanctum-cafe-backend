const mongoose = require('mongoose');

const masterValueSchema = new mongoose.Schema({
  type: {
    type: String,
    enum: ['Expense Category', 'Payment Method', 'Cleaning Checklist', 'Hygiene Checklist', 'Employee'],
    required: true,
  },
  value: { type: String, required: true, trim: true },
}, { timestamps: true });

masterValueSchema.index({ type: 1, value: 1 }, { unique: true });

module.exports = mongoose.model('MasterValue', masterValueSchema);
