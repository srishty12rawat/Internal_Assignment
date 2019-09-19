const mongoose = require('mongoose');

const DeathDataSchema = mongoose.Schema({
    columns: Array,
    index: Array,
    data: Array
});

module.exports = mongoose.model('death_datas', DeathDataSchema);