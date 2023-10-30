const express = require('express');
const { getDummySheet, addSheet,getplantDetail } = require('../controller/sheetController');

const Router   =  express.Router();

Router.route('/getDummySheet/:reportId').get(getDummySheet)
Router.route('/postReport').post(addSheet)
Router.route('/plant/all').get(getplantDetail)
// Router.route('').get(getDummySheet)



module.exports = Router