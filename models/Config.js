const mongoose = require('mongoose');

const configSchema = new mongoose.Schema({
  cafeName:             { type: String, default: 'The Sanctum Cafe' },
  ownerEmail:           { type: String, default: '' },
  ownerPin:             { type: String, default: '1121' },
  staffPin:             { type: String, default: '1234' },
  inventoryPeriodDays:  { type: Number, default: 7 },
}, { timestamps: true });

module.exports = mongoose.model('Config', configSchema);
