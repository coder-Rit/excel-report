const mongoose  = require("mongoose");


const deviceSchema = mongoose.Schema({
    deviceId: String,
    machineName:String,
});

module.exports  =  mongoose.model("device", deviceSchema)