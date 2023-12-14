const express = require('express');
const { getDummySheet, addSheet,getplantDetail,createExcelSheet } = require('../controller/sheetController');

const Router   =  express.Router();

Router.route('/getDummySheet/:reportId').get(getDummySheet)
Router.route('/postReport').post(addSheet)
Router.route('/plant/all').get(getplantDetail)


//task after exams
Router.route('/getSheet').get(createExcelSheet)
// Router.route('').get(getDummySheet)



module.exports = Router