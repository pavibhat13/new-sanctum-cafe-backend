const mongoose = require('mongoose');

const masterValueSchema = new mongoose.Schema({
  type: {
    type: String,
    enum: ['Expense Category', 'Payment Method', 'Cleaning Checklist', 'Mandatory Checklist', 'Employee', 'Inventory Category', 'Inventory Sub Category'],
    required: true,
  },
  value:          { type: String, required: true, trim: true },
  salary:         { type: Number },  // Employee only: fixed monthly salary
  allowedLeaves:  { type: Number },  // Employee only: free leaves per period
}, { timestamps: true });

masterValueSchema.index({ type: 1, value: 1 }, { unique: true });

module.exports = mongoose.model('MasterValue', masterValueSchema);
