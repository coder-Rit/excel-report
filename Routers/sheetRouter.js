const express = require("express");
const {
  getDummySheet,
  addSheet,
  getplantDetail,
  abc,
} = require("../controller/sheetController");
const {
 
    getSheet,
} = require("../controller/deviceController");

const Router = express.Router();

Router.route("/getDummySheet/:reportId").get(getDummySheet);
Router.route("/postReport").post(addSheet);
Router.route("/plant/all").get(getplantDetail);

//task after exams
Router.route("/getSheet").get(getSheet);
// Router.route('').get(getDummySheet)

module.exports = Router;
