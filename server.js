 
const app = require('./app.js');
const InitializeFile = require('./controller/fileWatcher.js');
const path = require('path');

const tankfilePath = path.resolve(__dirname,`./${process.env.WATCH_FILE_NAME}`)


app.listen(process.env.PORT,()=>{
    console.log("wlc to ",process.env.PORT);
})  


 
console.log("file waching started");
const funcObj = InitializeFile(tankfilePath,process.env.WATCH_FILE_URL)
funcObj.watchFile()
