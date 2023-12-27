const express = require("express");
const {
    watchFileAction
} = require("../controller/watchfileController");
const {
    getEReport
} = require("../controller/DailyPowerReport");
 
const Router = express.Router();

 
Router.route("/sendData").post(watchFileAction);
Router.route("/eReport").get(getEReport);
 
module.exports = Router; 
 