const mongoose = require('mongoose');

const staffLeaveSchema = new mongoose.Schema({
  employeeName: { type: String, required: true, trim: true },
  fromDate: { type: Date, required: true },
  toDate: { type: Date, required: true },
  leaveType: { type: String, enum: ['Full Day', 'Half Day'], default: 'Full Day' },
  notes: { type: String, trim: true },
  createdBy: { type: String, default: 'owner' },
}, { timestamps: true });

module.exports = mongoose.model('StaffLeave', staffLeaveSchema);
