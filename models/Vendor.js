const mongoose = require('mongoose');

const vendorSchema = new mongoose.Schema({
  name: { type: String, required: true, unique: true, trim: true },
  address: { type: String, trim: true },
  phone: { type: String, trim: true },
  bankName: { type: String, trim: true },
  accountNo: { type: String, trim: true },
  ifsc: { type: String, trim: true },
}, { timestamps: true });

module.exports = mongoose.model('Vendor', vendorSchema);
