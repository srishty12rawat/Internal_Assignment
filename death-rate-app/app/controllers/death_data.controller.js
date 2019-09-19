const death_datas = require('../models/death_data.model.js');


exports.find = (req, res) => {
    death_datas.find({},{columns:1})
    .then(notes => {
        res.send(notes);
    }).catch(err => {
        res.status(500).send({
            message: err.message || "Some error occurred while retrieving columns."
        });
    });

};

exports.findAll = (req, res) => {
    death_datas.find()
    .then(notes => {
        res.send(notes);
    }).catch(err => {
        res.status(500).send({
            message: err.message || "Some error occurred while retrieving data."
        });
    });

};

