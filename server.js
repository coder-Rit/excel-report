
const schedule = require('node-schedule');

const app = require('./app.js');





app.listen(process.env.PORT,()=>{
    console.log("wlc to ",process.env.PORT);
})  

const job = schedule.scheduleJob(' */13 * * * *', function async(){
    axios.post(" ").then(data=>console.log(data.data.msg));
  });
  