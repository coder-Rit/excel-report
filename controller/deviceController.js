const ExcelJS = require("exceljs");
const analogdatas = require("../analogdatas.json");

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
const colName = [
  "Group",
  "Sub Group 1",
  "Sub Group 2",
  "Sub Group 3",
  "Meters",
];
let map = [];

const data =[
  {
      "id": "n1701797057997",
      "text": "Power Vision1",
      "children": [
          {
              "id": "n1701797066456",
              "text": "Power Visoion G1",
              "children": [
                  {
                      "id": "n1701797076335",
                      "text": "Grid Incomer",
                      "children": [
                          {
                              "id": "n1701797085464",
                              "text": "Main I/C",
                              "children": [
                                  {
                                      "id": "D1",
                                      "text": "TG Meter",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  }
                              ],
                              "depth": 3,
                              "maxChildDepth": 1
                          }
                      ],
                      "depth": 2,
                      "maxChildDepth": 2
                  }
              ],
              "depth": 1,
              "maxChildDepth": 3
          },
          {
              "id": "n1701797140175",
              "text": "Power Visoion G2",
              "children": [
                  {
                      "id": "n1701797160551",
                      "text": "Generation",
                      "children": [
                          {
                              "id": "n1701797169495",
                              "text": "Generator",
                              "children": [
                                  {
                                      "id": "D2",
                                      "text": "GRID",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  }
                              ],
                              "depth": 3,
                              "maxChildDepth": 1
                          }
                      ],
                      "depth": 2,
                      "maxChildDepth": 2
                  }
              ],
              "depth": 1,
              "maxChildDepth": 3
          },
          {
              "id": "n1701797221024",
              "text": "Power Vision G3",
              "children": [
                  {
                      "id": "n1701797230112",
                      "text": "Bus Coupler",
                      "children": [
                          {
                              "id": "n1701797235951",
                              "text": "B/C",
                              "children": [
                                  {
                                      "id": "864180050392229",
                                      "text": "METER 3",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  },
                                  {
                                      "id": "D2",
                                      "text": "METER 7",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  }
                              ],
                              "depth": 3,
                              "maxChildDepth": 1
                          }
                      ],
                      "depth": 2,
                      "maxChildDepth": 2
                  }
              ],
              "depth": 1,
              "maxChildDepth": 3
          }
      ],
      "depth": 0,
      "maxChildDepth": 4
  },
  {
      "id": "n1701797302944",
      "text": "Power Vision2",
      "children": [
          {
              "id": "n1701797315216",
              "text": "Power Vision GA1",
              "children": [
                  {
                      "id": "n1701797321575",
                      "text": "Auxillary Consump",
                      "children": [
                          {
                              "id": "n1701797329471",
                              "text": "Aux",
                              "children": [
                                  {
                                      "id": "864180050392229",
                                      "text": "METER 4",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  },
                                  {
                                      "id": "D2",
                                      "text": "METER 5",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  }
                              ],
                              "depth": 3,
                              "maxChildDepth": 1
                          }
                      ],
                      "depth": 2,
                      "maxChildDepth": 2
                  }
              ],
              "depth": 1,
              "maxChildDepth": 3
          },
          {
            "id": "n1701797379424",
            "text": "Power Vision GA2",
            "children": [
                {
                    "id": "n1701797398295",
                    "text": "HT Motor",
                    "children": [
                        {
                            "id": "n1701797406703",
                            "text": "Sinter",
                            "children": [
                                {
                                    "id": "D1",
                                    "text": "METER 6",
                                    "type": "sys",
                                    "children": [],
                                    "depth": 4,
                                    "maxChildDepth": 0
                                },
                                {
                                    "id": "864180050392229",
                                    "text": "METER 8",
                                    "type": "sys",
                                    "children": [],
                                    "depth": 4,
                                    "maxChildDepth": 0
                                },
                                {
                                    "id": "D2",
                                    "text": "METER 9",
                                    "type": "sys",
                                    "children": [],
                                    "depth": 4,
                                    "maxChildDepth": 0
                                }
                            ],
                            "depth": 3,
                            "maxChildDepth": 1
                        }
                    ],
                    "depth": 2,
                    "maxChildDepth": 2
                }
            ],
            "depth": 1,
            "maxChildDepth": 3
        },
          {
              "id": "n1701797386122",
              "text": "Power Vision GA3",
              "children": [
                  {
                      "id": "n1701797428327",
                      "text": "Sinter 1",
                      "children": [
                          {
                              "id": "n1701797438271",
                              "text": "Sinter1 &3",
                              "children": [
                                  {
                                      "id": "864180050392229",
                                      "text": "METER 10",
                                      "type": "sys",
                                      "children": [],
                                      "depth": 4,
                                      "maxChildDepth": 0
                                  }
                              ],
                              "depth": 3,
                              "maxChildDepth": 1
                          }
                      ],
                      "depth": 2,
                      "maxChildDepth": 2
                  }
              ],
              "depth": 1,
              "maxChildDepth": 3
          },
          
      ],
      "depth": 0,
      "maxChildDepth": 4
  }
]


// fiding max depth of childern
const findMaxDepth = () => {
  
  let maxChildDepth = 0;
  for (let I = 0; I < data.length; I++) {
    const element = data[I];
    if (element.maxChildDepth>maxChildDepth) {
      maxChildDepth = element.maxChildDepth
    }
  }

   
  return maxChildDepth+1;
};

// some imp globle variables
const maxDepth = findMaxDepth();
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Sheet 1");
const filename = "output.xlsx"

const fillCell = (address, value) => {
  const cell = worksheet.getCell(address); //get cell using address
  cell.value = value; // assign value
};

//fulling paramter  into excel sheet called by mapGroup_subGroup
const MapParamsForDeviceId = (
  deviceFreqList,
  listOfDevices,
  start_col,
  start_row
 ) => {
  
  let cellAddress=""
  let currentCol = start_col;
  let currentRow = start_row;
  const comingRow = start_row;

  // storing group total
  let totalV1 = 0;
  let totalV5 = 0;
  let totalV13 = 0;
  let totalV30 = 0;


  // filling total sum value
  for (let I = 0; I < listOfDevices.length; I++) {
    const element = listOfDevices[I];
    if (element.t === "grp") {
      return;
    }
    console.log(element);
    currentRow = start_row + I;

    const deviceFreq = deviceFreqList.filter((data) => {
      return data.deviceId === element.id;
    });
    if (deviceFreq.length === 0) {
      console.log("device id missmatch");
      return;
    }

    const analogList = deviceFreq[0].analog;

    totalV1 = analogList.A1 + totalV1;
    totalV5 = analogList.A5 + totalV5;
    totalV13 = analogList.A13 + totalV13;
    totalV30 = analogList.A30 + totalV30;

    cellAddress = ColumnList[currentCol+1] + currentRow;
    fillCell(cellAddress, analogList.A1);

    cellAddress = ColumnList[currentCol + 3] + currentRow;
    fillCell(cellAddress, analogList.A5);

    cellAddress = ColumnList[currentCol + 5] + currentRow;
    fillCell(cellAddress, analogList.A13);

    cellAddress = ColumnList[currentCol + 7] + currentRow;
    fillCell(cellAddress, analogList.A30);
  }

  // filling gropu value
  cellAddress = ColumnList[currentCol + 2] + comingRow;
  fillCell(cellAddress, totalV1);
  cellAddress = ColumnList[currentCol + 4] + comingRow;
  fillCell(cellAddress, totalV5);
  cellAddress = ColumnList[currentCol + 6] + comingRow;
  fillCell(cellAddress, totalV13);
  cellAddress = ColumnList[currentCol + 8] + comingRow;
  fillCell(cellAddress, totalV30);


  //merging cells
  cellAddress = ColumnList[currentCol+2] + comingRow;
  console.log(`${cellAddress}:${ColumnList[currentCol+2] + currentRow}`);
  worksheet.mergeCells([`${cellAddress}:${ColumnList[currentCol+2] + currentRow}`]);
  cellAddress = ColumnList[currentCol+4] + comingRow;
  console.log(`${cellAddress}:${ColumnList[currentCol+2] + currentRow}`);
  worksheet.mergeCells([`${cellAddress}:${ColumnList[currentCol+4] + currentRow}`]);
  cellAddress = ColumnList[currentCol+6] + comingRow;
  console.log(`${cellAddress}:${ColumnList[currentCol+2] + currentRow}`);
  worksheet.mergeCells([`${cellAddress}:${ColumnList[currentCol+6] + currentRow}`]);
  cellAddress = ColumnList[currentCol+8] + comingRow;
  console.log(`${cellAddress}:${ColumnList[currentCol+2] + currentRow}`);
  worksheet.mergeCells([`${cellAddress}:${ColumnList[currentCol+8] + currentRow}`]);


};
 
// getting frequency of devices and sum of A1, A13 ...
const FreqAndSum = () => {
  for (let I = 0; I < analogdatas.length; I++) {
    let element = analogdatas[I];
    const idx = map.findIndex((data) => {
      return data.deviceId === element.deviceId;
    });

    const undefinedFIx = (value) => {
      if (typeof value === "undefined" || value === "N") {
        return 0;
      } else if (typeof value === "number") {
        return value;
      } else {
        let a = parseFloat(value);
        return a;
      }
    };

    if (idx === -1) {
      let myObj = {
        deviceId: element.deviceId,
        analog: {
          A1: undefinedFIx(element.analog.A1),
          A5: undefinedFIx(element.analog.A5),
          A13: undefinedFIx(element.analog.A13),
          A30: undefinedFIx(element.analog.A30),
        },
      };
      map.push(myObj);
    } else {
      let tempObj = map[idx];

      tempObj.analog.A1 = tempObj.analog.A1 + undefinedFIx(element.analog.A1);
      tempObj.analog.A5 = tempObj.analog.A5 + undefinedFIx(element.analog.A5);
      tempObj.analog.A13 =
        tempObj.analog.A13 + undefinedFIx(element.analog.A13);
      tempObj.analog.A30 =
        tempObj.analog.A30 + undefinedFIx(element.analog.A30);
      map[idx] = tempObj;
    }
  }
};

// // maping grups , subgruops and meter in excel sheet
// const mapGroup_subGroup =()=>{
//   let rowOffset = 2;
//   let maxRow = 0;

//   //calling for each child indivisualy
//   const assignCell = (Child, level) => {
//     level++;
//     let currentRow = rowOffset;
  
//     if (maxRow < currentRow) {
//       maxRow = currentRow;
//     }
//     if (level === 1 && Child.t === "d") {
//       MapParamsForDeviceId(map, [Child], maxDepth , currentRow, worksheet);
//     }
  
//     // console.log(Child.text , level);
//     let cellAddress;
//     if (Child.t === "d") {
//       rowOffset++;
//       cellAddress = ColumnList[maxDepth] + currentRow;
//     }
    
//     if (Child.t === "grp") {
//       cellAddress = ColumnList[level] + currentRow;
//       if (Child.children === null||Child.children.length ===0) {
//         //call for null
//        const final = ColumnList[maxDepth] + currentRow;
//        // call for null
//        fillCell(final,"NMF")
//        rowOffset++;
//        return;
//       }
//     }

    
    
//     fillCell(cellAddress,Child.text)
   

//     if (Child.children == null) {
//       return;
//     } else {
//       if (Child.children.length !== 0) {
//         MapParamsForDeviceId(
//           map,
//           Child.children,
//           maxDepth ,
//           currentRow,
//           worksheet
//         );
//       }
//     } 

    
  
//     for (let i = 0; i < Child.children.length; i++) {
//       assignCell(Child.children[i], level);
//     }
//     cellAddress = ColumnList[level] + currentRow;
//     worksheet.mergeCells([`${cellAddress}:${ColumnList[level] + maxRow}`]);
//   };

//   //caller
//   for (let I = 0; I < data.length; I++) {
//     assignCell(data[I], 0);
//   }

// }

const createExcelSheet =   () => {
   
  // row offset of groups or pinters
  let groupOffset_row = 2; // start without header
  let maxRow = 0;

  // recursive fuction for filling from right
  const fillFromRight = (data0, Column) => {
    //updating rows and columns
    let currentColumn = Column + 1;
    let currentRow = groupOffset_row;

    if (maxRow < currentRow) {
      maxRow = currentRow;
    }

    

    //counting max heigth of group to calculate next groups starting row point
    if (data0.children.length === 0) {
      groupOffset_row++;
    }

    if (data0.maxChildDepth ===1) {
      MapParamsForDeviceId(map,data0.children,currentColumn+1,currentRow)
    }

    // calling recusion for each subgroup items
    for (let index = 0; index < data0.children.length; index++) {
      fillFromRight(data0.children[index], currentColumn,);
    }

    // current cell address
    const cellAddress = ColumnList[currentColumn] + currentRow;
    const cell = worksheet.getCell(cellAddress); //get cell using address
    cell.value = data0.text; // assign value

    // merging cells
    worksheet.mergeCells([
      `${cellAddress}:${ColumnList[currentColumn] + maxRow}`,
    ]);
    //center alingment
    cell.alignment = textAlignment;

     
  };

  // calling the fuction for each group
  for (let index = 0; index < data.length; index++) {
    fillFromRight(data[index], 0, groupOffset_row);
  }

    
};


// setting dynamic header groups and static header for paramter headers 
const DynamicHeaderSetup = () => {
  let cellAddress = "";

  // helper func
  const fillCell = (address, value, col) => {
    worksheet.getColumn(ColumnList[col]).width = 13; //get cell using address
    const cell = worksheet.getCell(address); //get cell using address
    cell.value = value; // assign value

    cell.fill = headerFill;
    cell.font = { ...headerFont, size: 10 };
    cell.alignment = textAlignment;
  };



  //for groups
  for (let i = 1; i <= maxDepth; i++) {
    cellAddress = ColumnList[i] + 1;
    const cell = worksheet.getCell(cellAddress);
    worksheet.getColumn(ColumnList[i]).width = 20; //get cell using address
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = textAlignment;

    if (i === 1) {
      cell.value = "Group";
    } else if (i === maxDepth) {
      cell.value = "Meters";
    } else {
      cell.value = "Sub Group " + (i - 1);
    }
  }




  //for static paramters
  const parameterList = [
    "Total KW",
    "Group Total KW",
    "Total KVR",
    "Group Total KVR",
    "Total KVA",
    "Group Total KVA",
    "Total KWH",
    "Group Total KWH",
  ];
  for (let I = 0; I < parameterList.length; I++) {
    const element = parameterList[I];
    cellAddress = ColumnList[maxDepth + I + 1] + 1;
    fillCell(cellAddress, element, maxDepth + I + 1);
  }
};

// excel sheet genrator
const getSheet = async () => {
  const filename = "output.xlsx";

  await workbook.xlsx.writeFile(filename);
  console.log(`Excel sheet created and saved as ${filename}`);
};




// getting frequency of devices and sum of A1, A13 ...
FreqAndSum();

// maping grups , subgruops and meter in excel sheet
createExcelSheet() // calling  MapParamsForDeviceId from inside

// setting dynamic header groups and static header for paramter headers 
 DynamicHeaderSetup();
 
// excel sheet genrator
getSheet();
