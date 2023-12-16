const ExcelJS = require("exceljs");
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

const colName = [
  "Group",
  "Sub Group 1",
  "Sub Group 2",
  "Sub Group 3",
  "Meters",
];


const data = [
  { id: "CNC0284CCA864EEE8", text: "Device 1 Demo", t: "d", children: [
    {
      id: "CRM083AF2521DB0",
      text: "Clean Room Monitoring",
      t: "d",
      children:  null
    },
    { id: "134445566667777", text: "test", t: "d", children: null },
  ] },
  { id: "CNC0284CCA864EEE9", text: "Demo-Device 2", t: "d", children: [] },
  {
    id: "7",
    text: "Node 7",
    t: "grp",
    children: [
      { id: "wh43ekkw", text: "OEE System", t: "d", children: null },
      { id: "123", text: "PLC-1D-04J", t: "d", children: null },
    ],
  },
  {
    id: "n1702571622537",
    text: "Plant 2",
    t: "grp",
    children: [
      {
        id: "CRM083AF2521DB0",
        text: "Clean Room Monitoring",
        t: "d",
        children:  [
          {
            id: "CRM083AF2521DB0",
            text: "Clean Room Monitoring",
            t: "d",
            children:  [
              {
                id: "CRM083AF2521DB0",
                text: "Clean Room Monitoring",
                t: "d",
                children:  null
              },
              { id: "134445566667777", text: "test", t: "d", children: null },
            ]
          },
          { id: "134445566667777", text: "test", t: "d", children: null },
        ]
      },
      { id: "134445566667777", text: "test", t: "d", children: null },
    ],
  },
];

const findMaxDepth = (node) => {
  if (!node || !node.children) {
    return 0;
  }

  let maxChildDepth = 0;

  for (const child of node.children) {
    const childDepth = findMaxDepth(child);
    maxChildDepth = Math.max(maxChildDepth, childDepth);
  }

  return 1 + maxChildDepth;
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

  // list of cols for dynamic assingment of cells
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

  // row offset of groups or pinters
  let groupOffset_row = 2; // start without header
  let maxRow = 0;

  // recursive fuction for filling from right
  const fillFromRight = (data0, Column) => {
    if (!data0) {
      groupOffset_row++;
      return 0;
    }

    console.log(data0);

    //updating rows and columns
    let currentColumn = Column + 1;
    let currentRow = groupOffset_row;

    if (maxRow < currentRow) {
      maxRow = currentRow;
    }

    //counting max heigth of group to calculate next groups starting row point

    // calling recusion for each subgroup items

    for (let index = 0; index < data0.children.length; index++) {
      fillFromRight(data0.children[index], currentColumn);
    }

    // current cell address
    const cellAddress = ColumnList[currentColumn] + currentRow;
    const cell = worksheet.getCell(cellAddress); //get cell using address
    cell.value = data0.text; // assign value

    // merging cells
    //   worksheet.mergeCells([
    //     `${cellAddress}:${ColumnList[currentColumn] + maxRow}`
    //   ]);
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

const maxDepth = findMaxDepth({ children: data });

let rowOffset = 2;
let maxRow = 0;
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
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet("Sheet 1");

const assignCell = (Child, level) => {
  level++;
  let currentRow = rowOffset;

  if (maxRow < currentRow) {
    maxRow = currentRow;
  }

  // console.log(Child.text , level);
  console.log(Child.text, rowOffset);
  let cellAddress;
  if (Child.t === "d") {
    rowOffset++;
    cellAddress = ColumnList[maxDepth] + currentRow;
  }

  if (Child.t === "grp") {
    cellAddress = ColumnList[level] + currentRow;
  }

  const cell = worksheet.getCell(cellAddress); //get cell using address
  cell.value = Child.text;

  if (Child.children == null) {
    return;
  }

  for (let i = 0; i < Child.children.length; i++) {
    assignCell(Child.children[i], level);
  }
  cellAddress = ColumnList[level] + currentRow
  console.log( `${cellAddress}:${ColumnList[level] + maxRow}`);
  worksheet.mergeCells([
    `${cellAddress}:${ColumnList[level] + maxRow}`,
  ]);
};

console.log("Max depth of children:", maxDepth);






for (let I = 0; I < data.length; I++) {
  assignCell(data[I], 0);
}

 for (let i = 1; i <= maxDepth; i++) {

  let cellAddress = ColumnList[i] + 1;
  const cell = worksheet.getCell(cellAddress);
  worksheet.getColumn( ColumnList[i]).width = 20; //get cell using address
  cell.fill = headerFill;
  cell.font = headerFont;
  cell.alignment = textAlignment;

  if (i===1) {
    cell.value = "Group";
  }else if (i===maxDepth) {
    cell.value = "Meters";
  }else{
    cell.value = "Sub Group "+ (i-1);
  }
  
 }







const getSheet = async () => {
  const filename = "output.xlsx";

  await workbook.xlsx.writeFile(filename);
  console.log(`Excel sheet created and saved as ${filename}`);
};
getSheet();
