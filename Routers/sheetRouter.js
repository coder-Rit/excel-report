const express = require('express');
const { getDummySheet, addSheet } = require('../controller/sheetController');

const Router   =  express.Router();

Router.route('/getDummySheet/:reportId').get(getDummySheet)
Router.route('/postReport').post(addSheet)
// Router.route('').get(getDummySheet)



module.exports = Router