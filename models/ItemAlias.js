const mongoose = require('mongoose');

const itemAliasSchema = new mongoose.Schema({
  rawItem:     { type: String, required: true, unique: true, trim: true },
  generalItem: { type: String, required: true, trim: true },
}, { timestamps: true });

module.exports = mongoose.model('ItemAlias', itemAliasSchema);
