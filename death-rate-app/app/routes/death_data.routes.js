module.exports = (app) => {
    const death_datas = require('../controllers/death_data.controller.js');

    // Retrieve all data
    app.get('/data', death_datas.findAll);

    // Retrieve columns
    app.get('data/columns', death_datas.find)

}