const mongoose = require('mongoose');

const checklistLogSchema = new mongoose.Schema({
  date: { type: Date, required: true, default: Date.now },
  type: { type: String, enum: ['Cleaning', 'Mandatory'], required: true },
  items: [{ name: String, checked: Boolean }],
  remarks: { type: String, trim: true },
  createdBy: { type: String, default: 'staff' },
}, { timestamps: true });

module.exports = mongoose.model('ChecklistLog', checklistLogSchema);
