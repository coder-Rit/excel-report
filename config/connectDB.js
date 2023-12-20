


const mongoose = require('mongoose');

const connectDB  =()=>{

    
    mongoose.connect(process.env.DB_URL).then(()=>{
        console.log("db successfull");
    }).catch((err)=>console.log("oops error ", err))
    
}

module.exports = connectDB; 