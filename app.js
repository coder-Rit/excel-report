const bodyParser = require('body-parser');
const express = require('express');
const connectDB = require('./config/connectDB.js');
const app  = express()


app.use(bodyParser.json())


require('dotenv').config({path:"./config/config.env"})



connectDB()

const sheetRouter = require("./Routers/sheetRouter.js");
const error = require('./middelware/error.js'); 

app.use('/api/v1',sheetRouter)
// app.use('',sheetRouter)


app.use(error)

module.exports = app; 
