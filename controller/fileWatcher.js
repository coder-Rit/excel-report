const fs = require('fs');

const axios = require('axios');




function InitializeFile(filePath,watchFileUrl) {

    var data;
    var headerArray;
    var lastEntryArray;
    var JSONobj ={};

    function readfile() {

        try { 
            data =   fs.readFileSync(filePath,'utf8');
            proccessData()

        } catch (error) {
            console.log("here is error for you ", error);
        }
        
    }

    const postData = async ()=>{
        try {
            await axios.post(watchFileUrl, { JSONobj });
            console.log('Data posted successfully.');
        } catch (error) {
            console.error('Error while posting data:', error.message);
        }
    }

    function  proccessData () {
        const lines = data.split('\n')
       
         headerArray = lines[0].replace(/[., ]/g, '');  // output -> DateTime|Level1m|Level2m 
         headerArray = headerArray.split("|")
       
         lastEntryArray = lines[lines.length-2].replace(/[ ]/g, '');  // output -> DateTime|Level1m|Level2m 
         lastEntryArray = lastEntryArray.split("|")

         
         
         for (let i = 0; i < headerArray.length; i++) {
            JSONobj = {...JSONobj,[headerArray[i]]:lastEntryArray[i]}
          }

          postData()
        

    }
 

    function watchFile() {
        console.log("waching file...");
        fs.watchFile(filePath, (curr, prev) => {
            if (curr.mtime > prev.mtime) {
                readfile(filePath)
            }
          });
    }
 
    return{
        readfile,proccessData,watchFile

    } 
}
 


module.exports =  InitializeFile;
    