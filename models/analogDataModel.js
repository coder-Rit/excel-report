const mongoose  = require("mongoose");


const analogData = mongoose.Schema({
    deviceId: String,
    analog:Object,
    createdAt:{
        type: Date,
        default:Date.now(),
    } 
});

module.exports  =  mongoose.model("analogdata", analogData)