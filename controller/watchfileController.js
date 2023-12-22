 
const catchAsynch = require("../middelware/catchAsynch");


exports.watchFileAction = catchAsynch(async (req, res, next) => {

    console.log(req.body);
  
    res.status(200).json({
      msg: "got the data",
      res: req.body,
    });
  });