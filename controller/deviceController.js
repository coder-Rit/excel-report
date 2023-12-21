const ExcelJS = require("exceljs");
let analogdatas = require("../analogdatas.json");
const fs = require("fs");
const path = require("path");
const catchAsynch = require("../middelware/catchAsynch");
const analogDataModel = require("../models/analogDataModel");

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
 
let map = []; 
     
const data = [
  { id: "D1", text: "Device 1 Demo", t: "d", children: [] },
  { id: "D2", text: "Demo-Device 2", t: "d", children: [] },
  {
    id: "7",
    text: "Node 7",
    t: "grp",
    children: [
      { id: "D2", text: "OEE System", t: "d", children: null },
      {
        id: "123",
        text: "PLC-1D-04J",
        t: "grp",
        children: [{ id: "D2", text: "Demo-Device 2", t: "d", children: [] }],
      },
    ],
  },
  {
    id: "n1702571622537",
    text: "Plant 2",
    t: "grp",
    children: [
      {
        id: "D1",
        text: "Clean Room Monitoring",
        t: "d",
        children: null,
      },
      { id: "D2", text: "test", t: "d", children: null },
    ],
  },
];

const groupsNsubGroups = [
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
                    id: "D2",
                    text: "METER 3",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "D2",
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
  {
    id: "n1701797302944",
    text: "Power Vision2",
    children: [
      {
        id: "n1701797315216",
        text: "Power Vision GA1",
        children: [
          {
            id: "n1701797321575",
            text: "Auxillary Consump",
            children: [
              {
                id: "n1701797329471",
                text: "Aux",
                children: [
                  {
                    id: "D2",
                    text: "METER 4",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "D1",
                    text: "METER 5",
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
        id: "n1701797386122",
        text: "Power Vision GA3",
        children: [
          {
            id: "n1701797428327",
            text: "Sinter 1",
            children: [
              {
                id: "n1701797438271",
                text: "Sinter1 &3",
                children: [
                  {
                    id: "D1",
                    text: "METER 10",
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
        id: "n1701797379424",
        text: "Power Vision GA2",
        children: [
          {
            id: "n1701797398295",
            text: "HT Motor",
            children: [
              {
                id: "n1701797406703",
                text: "Sinter",
                children: [
                  {
                    id: "D2",
                    text: "METER 6",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "D1",
                    text: "METER 8",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "D2",
                    text: "METER 9",
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

// fiding max depth of childern
const findMaxDepth = () => {
  let maxChildDepth = 0;
  for (let I = 0; I < groupsNsubGroups.length; I++) {
    const element = groupsNsubGroups[I];
    if (element.maxChildDepth > maxChildDepth) {
      maxChildDepth = element.maxChildDepth;
    }
  }

  return maxChildDepth + 1;
};

// some imp globle variables
const maxDepth = findMaxDepth();
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Sheet 1");
const filename = "output23.xlsx";
const tableContentStart = 17;
let idList = [];

const fillCell = (address, value) => {
  const cell = worksheet.getCell(address); //get cell using address
  cell.value = value; // assign value
};

//fulling paramter  into excel sheet called by mapGroup_subGroup
const MapParamsForDeviceId = async(deviceFreqList, start_col, start_row) => {
  let cellAddress = "";
  let currentCol = start_col ;
  let currentRow = start_row;
  const comingRow = start_row;

  // storing group total
  let totalV1 = 0;
  let totalV5 = 0;      
  let totalV13 = 0;         
  let totalV30 = 0;   
  

  // filling total sum value
  for (let I = 0; I < deviceFreqList.length; I++) {
    const element = deviceFreqList[I];
    
    currentRow = start_row + I;

    if (element.status !="not found") {
      totalV1 = element.sumA1 + totalV1;     
      totalV5 = element.sumA5 + totalV5;
      totalV13 = element.sumA13 + totalV13;    
      totalV30 = element.sumA30 + totalV30;
    } 

   

    cellAddress = ColumnList[currentCol + 1] + currentRow;
    console.log(cellAddress,element.sumA1);
    await fillCell(cellAddress, element.sumA1);

    cellAddress = ColumnList[currentCol + 3] + currentRow;
    await fillCell(cellAddress, element.sumA5);

    cellAddress = ColumnList[currentCol + 5] + currentRow;
    await fillCell(cellAddress, element.sumA13);

    cellAddress = ColumnList[currentCol + 7] + currentRow;
    await  fillCell(cellAddress, element.sumA30);
  }
 

  // filling gropu value
  cellAddress = ColumnList[currentCol + 2] + comingRow;
  await fillCell(cellAddress, totalV1);
  cellAddress = ColumnList[currentCol + 4] + comingRow;
  await fillCell(cellAddress, totalV5);
  cellAddress = ColumnList[currentCol + 6] + comingRow;
  await fillCell(cellAddress, totalV13);
  cellAddress = ColumnList[currentCol + 8] + comingRow;
 await fillCell(cellAddress, totalV30);

  //merging cells
  cellAddress = ColumnList[currentCol + 2] + comingRow;
  worksheet.mergeCells([
    `${cellAddress}:${ColumnList[currentCol + 2] + currentRow}`,
  ]); 
  cellAddress = ColumnList[currentCol + 4] + comingRow;
  worksheet.mergeCells([
    `${cellAddress}:${ColumnList[currentCol + 4] + currentRow}`,
  ]);
  cellAddress = ColumnList[currentCol + 6] + comingRow;
  worksheet.mergeCells([
    `${cellAddress}:${ColumnList[currentCol + 6] + currentRow}`,
  ]);
  cellAddress = ColumnList[currentCol + 8] + comingRow;
  worksheet.mergeCells([
    `${cellAddress}:${ColumnList[currentCol + 8] + currentRow}`,
  ]);


};

const idSelector = async() => {
  function getLeafNodesWithTypeD(node) {
    if (node.children && node.children.length > 0) {
      return node.children.reduce(
        (acc, child) => acc.concat(getLeafNodesWithTypeD(child)),
        []
      );
    } else if (node.t === "d") {
      return [node.id];
    } else {
      return [];
    }
  }

  const leafNodesWithTypeD = data.reduce(
    (acc, node) => acc.concat(getLeafNodesWithTypeD(node)),
    []
  );
  idList = leafNodesWithTypeD;
};

// getting frequency of devices and sum of A1, A13 ...
const FreqAndSum = async (ChildesList, currentColumn, currentRow) => {

  const getCommand = (id) => {
    return [
      {
        $match: {
          deviceId: id,
        },
      },
      {
        $group: {
          _id: "$analog",
        },
      },
      {
        $project: {
          A1: { $toDouble: "$_id.A1" },
          A5: { $toDouble: "$_id.A5" },
          A13: { $toDouble: "$_id.A5" },
          A30: { $toDouble: "$_id.A5" },
        },
      },
      {
        $group: {
          _id: null,
          sumA1: { $sum: "$A1" },
          sumA5: { $sum: "$A5" },
          sumA13: { $sum: "$A13" },
          sumA30: { $sum: "$A30" },
        },
      },
    ];
  };

  let subMap = [];

  for (let i = 0; i < (ChildesList.length ); i++) {
    let result = await analogDataModel.aggregate(getCommand(ChildesList[i].id));
    console.log("result",result);
    if (result.length ===0 || !result) {
      result =[{
        status:"not found",
        _id:ChildesList[i].id,
        sumA1:"No Data",
        sumA5:"No Data",
        sumA13:"No Data",
        sumA30:"No Data",
      }]
    }else{
      result[0]._id = ChildesList[i].id;
    }
    console.log(result); 
    subMap.push(result[0]);
  }
  console.log("aggregate ",subMap);
  await  MapParamsForDeviceId(subMap, currentColumn, currentRow);
 
};
     
const logoMaping = async () => {
  const galanfi = path.resolve(__dirname, "../images/galanfi.png");
  const murgesh = path.resolve(__dirname, "../images/murgesh.png");

  const galanfiBase64 = fs.readFileSync(galanfi, { encoding: "base64" });
  const murgeshBase64 = fs.readFileSync(murgesh, { encoding: "base64" });

  const galanfiId = workbook.addImage({
    base64: galanfiBase64,
    extension: "png",
  });
  const murgeshId = workbook.addImage({
    base64: murgeshBase64,
    extension: "png",
  });

  // for murgesh
  worksheet.mergeCells(0,0,11,3);
  // for galanfi
  worksheet.mergeCells(0,8,11,14);



  worksheet.addImage(galanfiId, {
    tl: { col: 8, row: 3 },
    ext: { width: 372.28, height: 60.47 },
  });
  worksheet.addImage(murgeshId, {
    tl: { col: 1, row: 1 },
    ext: { width: 134.55, height: 123.21 }, 
  });
};

const createExcelSheet =  async(tableContentStart) => {
  // row offset of groups or pinters
  let groupOffset_row = tableContentStart; // start without header
  let maxRow = 0;

  // recursive fuction for filling from right
  const fillFromRight =  async(data0, Column) => {
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

    if (data0.maxChildDepth === 1) {   
      console.log("finding aggrgation for ",data0.children);
       await FreqAndSum(data0.children, currentColumn + 1, currentRow);
    }

    // calling recusion for each subgroup items
    for (let index = 0; index < data0.children.length; index++) {
      await  fillFromRight(data0.children[index], currentColumn);
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
  for (let index = 0; index < groupsNsubGroups.length; index++) {
    await fillFromRight(groupsNsubGroups[index], 0);
  }
};

// setting dynamic header groups and static header for paramter headers
const DynamicHeaderSetup =  async(headerStart) => {
  let cellAddress = "";

  // helper func
  const fillCell =  async(address, value, col) => {
    worksheet.getColumn(ColumnList[col]).width = 13; //get cell using address
    const cell = worksheet.getCell(address); //get cell using address
    cell.value = value; // assign value

    cell.fill = headerFill;
    cell.font = { ...headerFont, size: 10 };
    cell.alignment = textAlignment;
  };

  //for groups
  for (let i = 1; i <= maxDepth; i++) {
    cellAddress = ColumnList[i] + headerStart;
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
    cellAddress = ColumnList[maxDepth + I + 1] + headerStart;
    await fillCell(cellAddress, element, maxDepth + I + 1);
  }
};

const reportDetailsHeader = async()=>{
  
  
  const data = {
    name:"NIRANI SUGAR LIMITED",
    address:"sdfas kl skldfja fklasfasoa  adsdfaso",
    reportName:"DISTELLARY ENERGY REPORT",
    details:{
      DateFrom:"10/9/2023",
      DateTo:"10/9/2023",
      Parameter:"Energy",
    },
    reportHeadingLine:"abc"
  }
  

  let companyName = worksheet.getCell("D1"); //getcompanyName using address
 companyName.value = data.name;
 companyName.style.font = {
    color: { argb: "003366" },
    size: 30, // Font color (e.g., black)
    bold:true
  };
 companyName.style.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF" },
  };
 companyName.style.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
};
 companyName.alignment =textAlignment; 
  worksheet.mergeCells(0,4,3,7);


  //address

  
  let companyAddress  = worksheet.getCell("D4"); //get companyAddress using address

 companyAddress.value = data.address;
 companyAddress.style.font = {
    color: { argb: "black" }, 
    size: 13, // Font color (e.g., black)
    bold:true
  };
 companyAddress.style.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF" },
  };
 companyAddress.style.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
};
 companyAddress.alignment =textAlignment; 
  worksheet.mergeCells(4,4,4,7);




// reportname 
 
 let reportName  = worksheet.getCell("D5"); //get reportName using address

 reportName.value = data.reportName;
 reportName.style.font = {
    color: { argb: "048204" }, 
    size: 20, // Font color (e.g., black)
    bold:true
  };
 reportName.style.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFFFF" },
  };
 reportName.style.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
};
 reportName.alignment =textAlignment; 
  worksheet.mergeCells(5,4,6,7);

 
  let DateFrom  = worksheet.getCell("D7");
  let DateTo  = worksheet.getCell("D8");
  let parameter  = worksheet.getCell("D9");

  let DateFromValue  = worksheet.getCell("E7");
  let DateToValue  = worksheet.getCell("E8");
  let parameterValue  = worksheet.getCell("E9");

  const styleMaping =(cell,value)=>{
    cell.value = value;
    cell.style.font = {
       color: { argb: "3e658b" }, 
       size: 12, // Font color (e.g., black)
     };
    cell.style.fill = {
       type: "pattern",
       pattern: "solid",
       fgColor: { argb: "FFFFFF" },
     };
    cell.style.border = { 
       top: { style: 'thin' }, 
       left: { style: 'thin' },
       bottom: { style: 'thin' },
       right: { style: 'thin' },
   };
    cell.alignment =textAlignment; 
  }

  styleMaping(DateFrom,"Date From");
  styleMaping(DateTo,"Date To");
  styleMaping(parameter,"parameter");

  styleMaping(DateFromValue,data.details.DateFrom); 
  styleMaping(DateToValue,data.details.DateTo);
  styleMaping(parameterValue,data.details.Parameter);
 
  worksheet.mergeCells(7,5,7,7); 
  worksheet.mergeCells(8,5,8,7);
  worksheet.mergeCells(9,5,9,7);

  //blank space

let blankSpace  = worksheet.getCell("D10"); //get reportName using address
  
blankSpace.style.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "00ccff" },
  };
  blankSpace.style.border = { 
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
};
 
  worksheet.mergeCells(10,4,11,7);


  //purple line 

  let purpleLine  = worksheet.getCell("A12"); //get reportName using address
 
  purpleLine.style.fill = {
    type: "pattern",
    pattern: "solid", 
    fgColor: { argb: "333399" },
  };
  purpleLine.style.border = { 
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
};
 
  worksheet.mergeCells(12,0,12,maxDepth+8);

 


  //yellow reportHeadingLine 

  let reportHeadingLine  = worksheet.getCell("A13"); 
 
  
 reportHeadingLine.value = data.reportHeadingLine;
 reportHeadingLine.style.font = {
    color: { argb: "black" }, 
    size: 25, // Font color (e.g., black)
    bold:true 
  };
 reportHeadingLine.style.fill = {
    type: "pattern",
    pattern: "solid",  
    fgColor: { argb: "ffff00" },
  };
 reportHeadingLine.style.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' },
};
 reportHeadingLine.alignment =textAlignment; 
 
  worksheet.mergeCells(13,0,15,maxDepth+8);








} 

// excel sheet genrator
const getSheetFile = async () => {
  

  await workbook.xlsx.writeFile(filename);
  console.log(`Excel sheet created and saved as ${filename}`);
};

// // maping grups , subgruops and meter in excel sheet
// // createExcelSheet(tableContentStart) // calling  MapParamsForDeviceId from inside
// createExcelSheettd(tableContentStart);
// // setting dynamic header groups and static header for paramter headers
// DynamicHeaderSetup(tableContentStart - 1);

// // excel sheet genrator
// getSheet();

 const callme = async ()=>{
  await idSelector();
  //logo setup
  await logoMaping();

  // getting frequency of devices and sum of A1, A13 ...
  
  await createExcelSheet(tableContentStart); // calling  MapParamsForDeviceId from inside

  // setting dynamic header groups and static header for paramter headers
  await DynamicHeaderSetup(tableContentStart - 1);
  await reportDetailsHeader();

  // excel sheet genrator
  await getSheetFile();

 }

 callme()
exports.getSheet = catchAsynch(async (req, res, next) => {
 
  
});
