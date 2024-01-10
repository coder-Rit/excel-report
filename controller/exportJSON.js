const analogDataModel = require("../models/analogDataModel");
const fs = require("fs");

async function exportDataByYear(year, pageSize) {
  try {
    let page = 1;
    let dataToExport = [];
    let moreData = true;
    while (moreData) {
      const query = {
        createdAt: {
          $gte: new Date(`${year}-01-01T00:00:00.000Z`),
          $lt: new Date(`${year + 1}-01-01T00:00:00.000Z`),
        },
      };

      const result = await analogDataModel
        .find(query)
        .sort({ createdAt: 1 }) // Adjust the sorting 
        .skip((page - 1) * pageSize)
        .limit(pageSize)
        .lean();

      if (result.length === 0) {
        moreData = false;
      } else {
        dataToExport = dataToExport.concat(result);
        page++;
      }
    }

    // Export the paginated data to a JSON file
    const outputFilePath = `AnalogJSON_${year}.json`;
    fs.writeFileSync(outputFilePath, JSON.stringify(dataToExport, null, 2));

    console.log(`Data exported successfully to ${outputFilePath}`);
  } catch (err) {
    console.log(err);
  }
} 

const inputYear = 2023; //  user input
const pageSize = 1000; // Adjust the page size 



exports.justRun =()=>{

    exportDataByYear(inputYear, pageSize);

    
}
 