const ExcelJS = require("exceljs");

const data = [
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
                    id: "860987057798875_1",
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
                    id: "860987057798875_2",
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
                    id: "860987057798875_3",
                    text: "METER 3",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "860987057798875_7",
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
                    id: "860987057798875_4",
                    text: "METER 4",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "860987057798875_5",
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
                    id: "860987057798875_10",
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
                    id: "860987057798875_6",
                    text: "METER 6",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "860987057798875_8",
                    text: "METER 8",
                    type: "sys",
                    children: [],
                    depth: 4,
                    maxChildDepth: 0,
                  },
                  {
                    id: "860987057798875_9",
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

const analogdatas = require("../analogdatas.json");
let map = [];

const filename = "output2.xlsx";
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

const header_settings = (worksheet, textAlignment) => {
  //coloum settings
  worksheet.getColumn("A").width = 20;
  worksheet.getColumn("B").width = 20;
  worksheet.getColumn("C").width = 20;
  worksheet.getColumn("D").width = 20;
  worksheet.getColumn("E").width = 20;

  //Table header settings
  const headerFill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "305496" },
  };
  const headerFont = {
    color: { argb: "FFFFFF" },
    size: 14, // Font color (e.g., black)
  };

  const colName = [
    "Group",
    "Sub Group 1",
    "Sub Group 2",
    "Sub Group 3",
    "Meters",
  ];
  worksheet.addRow(colName);
  const headerRow = worksheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = headerFill;
    cell.font = headerFont;
    cell.alignment = textAlignment;
  });
};
const MapParamsForDeviceId = (
    deviceFreqList,
    listOfDevices,
    start_col,
    start_row,
    worksheet
  ) => {
    const fillCell = (address, value) => {
      const cell = worksheet.getCell(address); //get cell using address
      cell.value = value; // assign value
    };
  
    let currentCol = start_col;
    let currentRow = start_row;
  
    let totalV1 = 0;
    let totalV5 = 0;
    let totalV13 = 0;
    let totalV30 = 0;
  
    for (let I = 0; I < listOfDevices.length; I++) {
      const element = listOfDevices[I];
  
      const deviceFreq = deviceFreqList.filter((data) => {
        return data.deviceId === element.id;
      });
  
      const analogList = deviceFreq[0].analog;
  
      totalV1 = analogList.A1 + totalV1;
      totalV5 = analogList.A5 + totalV5;
      totalV13 = analogList.A13 + totalV13;
      totalV30 = analogList.A30 + totalV30;
  
      let cellAddress = ColumnList[currentCol] + currentRow;
      fillCell(cellAddress, analogList.A1);
  
      cellAddress = ColumnList[currentCol + 2] + currentRow;
      fillCell(cellAddress, analogList.A5);
  
      cellAddress = ColumnList[currentCol + 4] + currentRow;
      fillCell(cellAddress, analogList.A13);
  
      cellAddress = ColumnList[currentCol + 6] + currentRow;
      fillCell(cellAddress, analogList.A13);
    }
    cellAddress = ColumnList[currentCol + 1] + currentRow;
    fillCell(cellAddress, totalV1);
    cellAddress = ColumnList[currentCol + 3] + currentRow;
    fillCell(cellAddress, totalV5);
    cellAddress = ColumnList[currentCol + 5] + currentRow;
    fillCell(cellAddress, totalV13);
    cellAddress = ColumnList[currentCol + 7] + currentRow;
    fillCell(cellAddress, totalV13);
  };
const createExcelSheet = async (data, filename) => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  // globle styles
  const FontBoldness = {
    bold: true,
  };
  const textAlignment = {
    horizontal: "center",
    vertical: "middle",
  };

  // header style hanlder
  header_settings(worksheet, textAlignment);

  // list of cols for dynamic assingment of cells

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

    //making last col E bold
    if (ColumnList[currentColumn] === "E") {
      cell.font = FontBoldness;
    }
  };

  // calling the fuction for each group
  for (let index = 0; index < data.length; index++) {
    fillFromRight(data[index], 0, groupOffset_row);
  }

  //storing file
  await workbook.xlsx.writeFile(filename);
  console.log(`Excel sheet created and saved as ${filename}`);
};



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

      if (element.analog.A5 === "N") {
        console.log(element.analog.A5);
      }

      tempObj.analog.A1 = tempObj.analog.A1 + undefinedFIx(element.analog.A1);
      tempObj.analog.A5 = tempObj.analog.A5 + undefinedFIx(element.analog.A5);
      tempObj.analog.A13 =
        tempObj.analog.A13 + undefinedFIx(element.analog.A13);
      tempObj.analog.A30 =
        tempObj.analog.A30 + undefinedFIx(element.analog.A30);
      map[idx] = tempObj;
    }
  }
  console.log(map);
};

FreqAndSum();
//   createExcelSheet(data, filename);
