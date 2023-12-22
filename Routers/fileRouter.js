const express = require("express");
const {
    watchFileAction
} = require("../controller/watchfileController");
 
const Router = express.Router();

 
Router.route("/sendData").post(watchFileAction);
 
module.exports = Router;
