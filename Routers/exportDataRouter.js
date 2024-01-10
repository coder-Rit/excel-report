const express = require('express');
const { justRun } = require("../controller/exportJSON");

 
const Router = express.Router(); 
//task after exams
Router.route("/exportData").get(justRun);
// Router.route('').get(getDummySheet)

module.exports = Router;
