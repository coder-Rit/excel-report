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
      { id: "D1", text: "OEE System", t: "d", children: null },
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
const tableContentStart = 10;
let idList = [];

const fillCell = (address, value) => {
  const cell = worksheet.getCell(address); //get cell using address
  cell.value = value; // assign value
};

//fulling paramter  into excel sheet called by mapGroup_subGroup
const MapParamsForDeviceId = (deviceFreqList, start_col, start_row) => {
  let cellAddress = "";
  let currentCol = start_col ;
  let currentRow = start_row;
  const comingRow = start_row;

  // storing group total
  let totalV1 = 0;
  let totalV5 = 0;      
  let totalV13 = 0;         
  let totalV30 = 0;   
 

  fillCell("F12", "lk");


  // // filling total sum value
  // for (let I = 0; I < deviceFreqList.length; I++) {
  //   const element = deviceFreqList[I];
  //   console.log(element, currentCol, start_row);
  //   // if (element.t === "grp") {
  //   //   return;
  //   // }
  //   currentRow = start_row + I;

  //   // const deviceFreq = deviceFreqList.filter((data) => {
  //   //   return data.deviceId === element.id;
  //   // });
   
  //   totalV1 = element.sumA1 + totalV1;
  //   totalV5 = element.sumA5 + totalV5;
  //   totalV13 = element.sumA13 + totalV13;
  //   totalV30 = element.sumA30 + totalV30;

  //   cellAddress = ColumnList[currentCol + 1] + currentRow;
  //   fillCell(cellAddress, element.sumA1);

  //   cellAddress = ColumnList[currentCol + 3] + currentRow;
  //   fillCell(cellAddress, element.sumA5);

  //   cellAddress = ColumnList[currentCol + 5] + currentRow;
  //   fillCell(cellAddress, element.sumA13);

  //   cellAddress = ColumnList[currentCol + 7] + currentRow;
  //   fillCell(cellAddress, element.sumA30);
  // }

  // console.log(totalV1);
  // console.log(totalV5);
  // console.log(totalV13);
  // console.log(totalV30);

  // // filling gropu value
  // cellAddress = ColumnList[currentCol + 2] + comingRow;
  // fillCell(cellAddress, totalV1);
  // cellAddress = ColumnList[currentCol + 4] + comingRow;
  // fillCell(cellAddress, totalV5);
  // cellAddress = ColumnList[currentCol + 6] + comingRow;
  // fillCell(cellAddress, totalV13);
  // cellAddress = ColumnList[currentCol + 8] + comingRow;
  // fillCell(cellAddress, totalV30);

  // //merging cells
  // cellAddress = ColumnList[currentCol + 2] + comingRow;
  // console.log(`${cellAddress}:${ColumnList[currentCol + 2] + currentRow}`);
  // worksheet.mergeCells([
  //   `${cellAddress}:${ColumnList[currentCol + 2] + currentRow}`,
  // ]);
  // cellAddress = ColumnList[currentCol + 4] + comingRow;
  // console.log(`${cellAddress}:${ColumnList[currentCol + 4] + currentRow}`);
  // worksheet.mergeCells([
  //   `${cellAddress}:${ColumnList[currentCol + 4] + currentRow}`,
  // ]);
  // cellAddress = ColumnList[currentCol + 6] + comingRow;
  // console.log(`${cellAddress}:${ColumnList[currentCol + 6] + currentRow}`);
  // worksheet.mergeCells([
  //   `${cellAddress}:${ColumnList[currentCol + 6] + currentRow}`,
  // ]);
  // cellAddress = ColumnList[currentCol + 8] + comingRow;
  // console.log(`${cellAddress}:${ColumnList[currentCol + 8] + currentRow}`);
  // worksheet.mergeCells([
  //   `${cellAddress}:${ColumnList[currentCol + 8] + currentRow}`,
  // ]);


};

const idSelector = () => {
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
  // const setAnalogData = async()=>{
  //   let perfectData = []
  //   analogdatas.map((ele)=>{
  //     let newEle = ele;

  //     //change the according to you
  //     delete newEle._id;
  //     delete newEle.__v;
  //     delete newEle.createdAt;
  //     perfectData.push(newEle)
  //   })

  //  await analogDataModel.insertMany(perfectData)

  // }

  // const getAnalogData = async()=>{

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

  for (let i = 0; i < ChildesList.length; i++) {
    let result = await analogDataModel.aggregate(getCommand(ChildesList[i].id));
    result[0]._id = ChildesList[i].id;

    subMap.push(result[0]);
  }

  MapParamsForDeviceId(subMap, currentColumn, currentRow);

  // setAnalogData()
  // getAnalogData()
};

const logoMaping = () => {
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

  worksheet.addImage(galanfiId, {
    tl: { col: 7, row: 2 },
    ext: { width: 372.28, height: 60.47 },
  });
  worksheet.addImage(murgeshId, {
    tl: { col: 1, row: 1 },
    ext: { width: 134.55, height: 123.21 },
  });
};

const createExcelSheet = (tableContentStart) => {
  // row offset of groups or pinters
  let groupOffset_row = tableContentStart; // start without header
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

    if (data0.maxChildDepth === 1) {
      FreqAndSum(data0.children, currentColumn + 1, currentRow);
    }

    // calling recusion for each subgroup items
    for (let index = 0; index < data0.children.length; index++) {
      fillFromRight(data0.children[index], currentColumn);
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
    fillFromRight(groupsNsubGroups[index], 0);
  }
};

// setting dynamic header groups and static header for paramter headers
const DynamicHeaderSetup = (headerStart) => {
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
    fillCell(cellAddress, element, maxDepth + I + 1);
  }
};

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

exports.getSheet = catchAsynch(async (req, res, next) => {
  idSelector();
  //logo setup
  logoMaping();

  // getting frequency of devices and sum of A1, A13 ...
  
  createExcelSheet(tableContentStart); // calling  MapParamsForDeviceId from inside

  // setting dynamic header groups and static header for paramter headers
  DynamicHeaderSetup(tableContentStart - 1);

  // excel sheet genrator
  getSheetFile();

  
});
