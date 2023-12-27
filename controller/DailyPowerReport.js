const ExcelJS = require("exceljs");
const analogDataModel = require("../models/analogDataModel");
const catchAsynch = require("../middelware/catchAsynch");

const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Sheet 1");
const filename = "energyReport.xlsx";

const headerFill = {
  type: "pattern",
  pattern: "solid",
  fgColor: { argb: "305496" },
};
const headerFont = {
  color: { argb: "FFFFFF" },
  size: 14, // Font color (e.g., black)
};
const FontBoldness = {
  bold: true,
};
const textAlignment = {
  horizontal: "center",
  vertical: "middle",
};
const ColumnList = [
  0,
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "K",
  "L",
  "M",
  "N",
  "O",
  "P",
  "Q",
  "R",
  "S",
  "T",
  "U",
  "V",
  "W",
  "X",
  "Y",
  "Z",
];
const headerName =["Particulars","shift1","shift2","shift3","To day","To Month","Till Date",];
const formattedData = [
  {
    id: "n1701797057997",
    text: "Power Vision1",
    children: [
      {
        id: "n1701797066456",
        text: "Power Visoion G1",
        children: [
          {
            id: "n1701797076335",
            text: "Grid Incomer",
            children: [
              {
                id: "n1701797085464",
                text: "Main I/C",
                children: [
                  {
                    id: "D1",
                    text: "TG Meter",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                ],
                depth: 3,
                maxChildDepth: 1,
              },
            ],
            depth: 2,
            maxChildDepth: 2,
          },
        ],
        depth: 1,
        maxChildDepth: 3,
      },
      {
        id: "n1701797140175",
        text: "Power Visoion G2",
        children: [
          {
            id: "n1701797160551",
            text: "Generation",
            children: [
              {
                id: "n1701797169495",
                text: "Generator",
                children: [
                  {
                    id: "D2",
                    text: "GRID",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                ],
                depth: 3,
                maxChildDepth: 1,
              },
            ],
            depth: 2,
            maxChildDepth: 2,
          },
        ],
        depth: 1,
        maxChildDepth: 3,
      },
      {
        id: "n1701797221024",
        text: "Power Vision G3",
        children: [
          {
            id: "n1701797230112",
            text: "Bus Coupler",
            children: [
              {
                id: "n1701797235951",
                text: "B/C",
                children: [
                  {
                    id: "864180050392229",
                    text: "METER 3",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "D1",
                    text: "METER 7",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                ],
                depth: 3,
                maxChildDepth: 1,
              },
            ],
            depth: 2,
            maxChildDepth: 2,
          },
        ],
        depth: 1,
        maxChildDepth: 3,
      },
    ],
    depth: 0,
    maxChildDepth: 4,
  },
];
var startRow = 1;
//data for headers

const headerDetails = {
  name: "READBULL DISTILARY",
  timeDetails: {
    startDate: "08/06/5260",
    startTime: "06:22",
    reportDate: "08/06/5260",
    tillDate: "06",
    toMonths: "22",
  },
  reportName: "DAILY POWER REPORT",
};

// common functions

const fillCell = (address, value,center,fill,fontFill,bold,size,tables) => {
  const cell = worksheet.getCell(address); //get cell using address
  cell.value = value; // assign value

  if (center) {
    cell.alignment = textAlignment;
  }
  if(fill){
    cell.style.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: fill },
    };

  }
  if (fontFill) {
    cell.style.font = {
      color: { argb: fontFill },
      size: size, // Font color (e.g., black)
      bold: bold,

    };
  }

  if (tables) {
    cell.style.border ={
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    }
  }
  
};

const mergeArea =(t,l,b,r)=>{ 
  worksheet.mergeCells(t,l,b,r) 

}

// important functions
function stylesInitializer() {
  function namePrinter() {
    let address = `${ColumnList[1]}${startRow}`;
    console.log(address);
    fillCell(address, headerDetails.name,true,"ffc000","ffffff",true,20,true);
    mergeArea(1,1,2,13)
    startRow += 2;
  }

 

  function headerDetailsPrinter() {

    function fillTwinCell(key, value, C_idx, R_idx) {
      let address = `${ColumnList[C_idx]}${R_idx}`;
      fillCell(address, key,true,"00b050","ffffff",false,10,true);;
      mergeArea(R_idx,C_idx,R_idx,C_idx+1)
      
      address = `${ColumnList[C_idx + 2]}${R_idx}`;
      fillCell(address, value,true,"00b050","ffffff",false,10,true);
      mergeArea(R_idx,C_idx + 2,R_idx,C_idx + 3)
      
  
    }

    const { tillDate, startDate, toMonths, reportDate, startTime } =
      headerDetails.timeDetails;

    fillTwinCell("Start Date", startDate, 3, 4);
    fillTwinCell("Start Time", startTime, 3, 5);
    fillTwinCell(" ", " ", 3, 6);
    
    fillTwinCell("reportDate", reportDate, 8, 4);
    fillTwinCell("toMonths", toMonths, 8, 5);
    fillTwinCell("tillDate", tillDate, 8, 6);

    startRow += 5;
  }

  function reportNamePrinter() {
    let address = `${ColumnList[1]}${startRow}`;
    console.log(address);

    fillCell(address, headerDetails.reportName,true,"ffffff","182361",true,16,true);
    mergeArea(startRow,1,startRow,13)
    startRow +=2;
  }

  return {
    namePrinter,
    reportNamePrinter,
    headerDetailsPrinter,
  };
}

function PowerReportInitializer(date ) {
   

  var deviceList = [];
  var arr = [];
  var subMap = [];

  function filterLeafNodes(data) {
    for (let i = 0; i < data.length; i++) {
      if (data[i].maxChildDepth === 0) {
        deviceList.push({
          id: data[i].id,
          text: data[i].text,
        });
      }
      filterLeafNodes(data[i].children);
    }
  }

  function getArrayOfshifts_Helper() {
    function add24HoursToDate(originalDate) {
      // Convert the string to a Date object
      var dateObject = new Date(originalDate);

      // Add 24 hours to the date
      dateObject.setHours(dateObject.getHours() + 24);

      // Convert the updated date object back to a string
      var updatedDateString = dateObject.toUTCString();

      return updatedDateString;
    }
    function calculateMonthStartEnd(inputDate) {
      // Convert input date string to a Date object
      const currentDate = new Date(inputDate);

      // Get the current month and year
      const currentMonth = currentDate.getMonth();
      const currentYear = currentDate.getFullYear();

      // Calculate the first day of the month
      const startOfMonth = new Date(currentYear, currentMonth, 1, 0, 0, 0, 0);

      // Calculate the last day of the month
      const lastDayOfMonth = new Date(
        currentYear,
        currentMonth + 1,
        0,
        23,
        59,
        59,
        999
      );

      // Format the output in the specified format
      const outputFormat = {
        weekday: "short",
        day: "2-digit",
        month: "short",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        timeZoneName: "short",
      };
      const startOfMonthString = startOfMonth.toLocaleString(
        "en-US",
        outputFormat
      );
      const lastDayOfMonthString = lastDayOfMonth.toLocaleString(
        "en-US",
        outputFormat
      );

      // Return the result
      return { start: startOfMonthString, end: lastDayOfMonthString };
    }

    arr.push({
      shift: "shift 1",
      query: {
        createdAt: {
          $gte: new Date(`${date} 06:00:00 GMT`),
          $lt: new Date(`${date} 14:00:00 GMT`),
        },
      },
    });
    arr.push({
      shift: "shift 1",
      query: {
        createdAt: {
          $gte: new Date(`${date} 14:00:00 GMT`),
          $lt: new Date(`${date} 22:00:00 GMT`),
        },
      },
    });
    arr.push({
      shift: "shift 1",
      query: {
        $and: [
          {
            createdAt: {
              $gte: new Date(`${date} 22:00:00 GMT`),
            },
          },
          {
            createdAt: {
              $lt: new Date(add24HoursToDate(`${date} 06:00:00 GMT`)),
            },
          },
        ],
      },
    });

    // for the day we will do maths

    // for  Month

    const Dobj = calculateMonthStartEnd(`${date} 14:00:00 GMT`);
    arr.push({
      shift: "To Month",
      query: {
        createdAt: {
          $gte: new Date(Dobj.start),
          $lt: new Date(Dobj.end),
        },
      },
    });

    // for till date
    arr.push({
      shift: "Till Date",
      query: {
        createdAt: {
          // two years before date
          $gte: new Date(`Fri, 1 Aug 2021 00:00:00 GMT`),
          $lt: new Date(add24HoursToDate(`${date} 00:00:00 GMT`)),
        },
      },
    });

    return arr;
  }

  function getPipeLine_Helper(deviceId, ...shift) {
    return [
      {
        $match: {
          $and: [
            {
              deviceId: deviceId,
            },
            {
              $or: [...shift],
            },
          ],
        },
      },
      {
        $group: {
          _id: "$analog",
        },
      },
      {
        $project: {
          A1: {
            $toDouble: "$_id.A1",
          },
        },
      },
      {
        $group: {
          _id: null,
          sumA1: {
            $sum: "$A1",
          },
        },
      },
    ];
  }

  async function genrateAggegation() {
    getArrayOfshifts_Helper();

    for (let i = 0; i < deviceList.length; i++) {
      let newEntrie = [];

      for (let J = 0; J < arr.length; J++) {
        const element = arr[J];

        // console.log(deviceList[i].id, element.query);
        let result = await analogDataModel.aggregate(
          getPipeLine_Helper(deviceList[i].id, element.query)
        );

        if (result.length === 0 || !result) {
          result = [
            {
              sumA1: "No Data",
              shift: element.shift,
            },
          ];
        } else {
          result[0].shift = element.shift;
        }
        delete result[0]._id;
        newEntrie.push(result[0]);
      }

      // console.log(newEntrie);

      const entrie = {
        paritcular: deviceList[i].text,
        id: deviceList[i].id,
        shift: newEntrie,
      };
      subMap.push(entrie);
    }
  }

  function fillHeader( ) {

    function fillHeader_helper(value, C_idx_H,R_idx_H) {
      address = `${ColumnList[C_idx_H + 1]}${R_idx_H}`;
      fillCell(address, value,true,"ffff00","black",true,10,true);
    } 
    let colm = 1

    fillHeader_helper(headerName[0],0,startRow)
       mergeArea(12,0,12,0)

    for (let i = 1; i < headerName.length; i++) {
        fillHeader_helper(headerName[i],colm,startRow)
        mergeArea(startRow,colm-1,startRow+1,colm)
        fillHeader_helper("KWH",colm,startRow+2)
        fillHeader_helper("MW",colm+1,startRow+2)
        colm += 2;
      }
      //last remaing cell
      mergeArea(startRow,colm-1,startRow+1,colm)
    
    startRow += 2;


  }

  function mapDataToExcle() {
    for (let i = 0; i < subMap.length; i++) {
      const element = subMap[i];

      //paricatular fill
      let address = `${ColumnList[1]}:${i + 1 + startRow}`;
      fillCell(address, element.paritcular, false,"","",false,12,true);

      // shifts fill

      const currentShift = element.shift;
      let colNumber = 2;
      let dayTotal = 0;
      let onceUserd = false
      for (let j = 0; j < currentShift.length; j++) {
        const shiftData = currentShift[j];

        if (j <= 2) {
          dayTotal += shiftData.sumA1;
        }

        // for day total

        if(j<=2) {
          // for KWh
          const KWhaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
          fillCell(KWhaddress, Math.round(shiftData.sumA1),true,"ddd9c3","black",false,12,true);
          colNumber++;

          //for MW
          const MWaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
          fillCell(MWaddress, Math.round(shiftData.sumA1 / 1000),true,"ddd9c3","black",false,12,true);
          colNumber++;
        }
        else if (j === 3 && !onceUserd) {
          onceUserd = true;
          // for KWh
          const KWhaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
          fillCell(KWhaddress, Math.round(dayTotal),true,"8eb4e3","black",false,12,true);
          colNumber++;

          //for MW
          const MWaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
          fillCell(MWaddress, Math.round(dayTotal / 1000),true,"8eb4e3","black",false,12,true);
          colNumber++;
          j--;

          // for other cases
        }
        else if(j===4){
          // for KWh
          const KWhaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
          fillCell(KWhaddress, Math.round(shiftData.sumA1),true,"e6b9b8","black",false,12,true);
          colNumber++;

          //for MW
          const MWaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
          fillCell(MWaddress, Math.round(shiftData.sumA1 / 1000),true,"e6b9b8","black",false,12,true);
          colNumber++;

        }else{
           // for KWh
           const KWhaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
           fillCell(KWhaddress, Math.round(shiftData.sumA1),true,"d7e4bd","black",false,12,true);
           colNumber++;
 
           //for MW
           const MWaddress = `${ColumnList[colNumber]}:${i + 1 + startRow}`;
           fillCell(MWaddress, Math.round(shiftData.sumA1 / 1000),true,"d7e4bd","black",false,12,true);
           colNumber++;
        }
      }

    }
  } 

  function saveFile() {
    workbook.xlsx
      .writeFile("electricReport.xlsx")
      .then(() => {
        console.log("Excel file created successfully!");
      })
      .catch((err) => {
        console.error("Error:", err);
      });
  }

  function sendFileInRes(req, res, next) {
    res.status(200).json({
      msg: "Check the files",
    });
  }

  return {
    filterLeafNodes,
    genrateAggegation,
    mapDataToExcle,
    saveFile,
    sendFileInRes,
    fillHeader
  };
}

exports.getEReport = catchAsynch(async (req, res, next) => {
  const styles = stylesInitializer(headerDetails);
  styles.namePrinter();
  styles.headerDetailsPrinter();
  styles.reportNamePrinter();

  const Report = PowerReportInitializer("Sat, 12 Aug 2023", 2);

  Report.filterLeafNodes(formattedData);
  await Report.genrateAggegation();
  Report.fillHeader();
  Report.mapDataToExcle();
  Report.saveFile();
  Report.sendFileInRes(req, res, next);

  // Save the workbook to a file
});
