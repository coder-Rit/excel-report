const mongoose = require("mongoose");

const ReportSchema = new mongoose.Schema({
 
    report:{
        type:"object"
    }

   
});

module.exports = mongoose.model("report", ReportSchema);
