const catchAsynch = require("../middelware/catchAsynch");
const ErrorHandler = require("../utils/ErrorHandler");
const Excel = require("excel4node");
const path = require("path");
const reports = require("../report.json");
const { type } = require("os");
const ReportSchema = {};


exports.getDummySheet = catchAsynch(async (req, res, next) => {
  let data;

  if (req.params.reportId === "all") {
    data = await ReportSchema.find({});
  } else {
    data = await ReportSchema.findById(req.params.reportId);
    data = [data];
  }

  const workbook = new Excel.Workbook();
  console.log(data);

  for (let reportNumber = 0; reportNumber < data.length; reportNumber++) {
    const {
      mainHeading,
      address,
      reportName,
      fromDate,
      toDate,
      parameter,
      reportHeading,
      records,
    } = data[reportNumber].report;
    const worksheet = workbook.addWorksheet(`Sheet ${reportNumber + 1}`);

    let groupBorder = {
      left: {
        style: "medium", // Border style (medium, medium, etc.)
        color: "black", // Border color
      },
      right: {
        style: "medium",
        color: "black",
      },
      top: {
        style: "medium",
        color: "black",
      },
      bottom: {
        style: "medium",
        color: "black",
      },
    };
    let indiBorder = {
      ...groupBorder,
      top: {
        style: "thin",
        color: "black",
      },
      bottom: {
        style: "thin",
        color: "black",
      },
    };

    //style for main heading
    const mainHeadingStyle = workbook.createStyle({
      font: {
        size: 20,
        color: "#002060",
        bold: true,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "white",
      },
      border: groupBorder,
    });

    //style for main heading
    const subHeading = workbook.createStyle({
      font: {
        size: 15,
        color: "#00B050",
        bold: true,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "white",
      },
      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },
      border: groupBorder,
    });

    //style for main heading
    const tableHeading = workbook.createStyle({
      font: {
        size: 22,
        color: "#002060",
        bold: true,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "#00b0f0",
      },
      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },
      border: groupBorder,
    });

    //style for main heading
    const voidEmptyCell = workbook.createStyle({
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "white",
      },
    });

    //style for report header
    const colHeader = workbook.createStyle({
      font: {
        size: 9,
        color: "white",
        bold: true,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "#305496",
      },
      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },
      border: groupBorder,
    });
    const style_srNo = workbook.createStyle({
      font: {
        size: 7,
        color: "black",
      },
      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },
      border: indiBorder,
    });

    //style for report header
    const style_megaBold = workbook.createStyle({
      font: {
        size: 12,
        color: "black",
        bold: true,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "white",
      },
      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },

      border: groupBorder,
    });

    //style for report header
    const style_simpleBold = workbook.createStyle({
      font: {
        size: 8,
        color: "black",
        bold: true,
      },
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: "white",
      },
      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },
      border: groupBorder,
    });

    //style for report header
    const idesStyle = workbook.createStyle({
      font: {
        size: 8,
        color: "black",
        bold: true,
      },

      alignment: {
        wrapText: true, // Enable text wrapping
        vertical: ["center"], // Vertical alignment (you can use 'middle' or 'bottom' as well)
        horizontal: ["center"], // Horizontal alignment (you can use 'center' or 'right' as well)
      },
    });

    //void cells for images nearby space filling
    //for murgesh
    worksheet.cell(1, 1, 7, 3).string("").style(voidEmptyCell);
    //for gal
    worksheet.cell(1, 7, 7, 11).string("").style(voidEmptyCell);

    //logo 1
    const murgesh = worksheet.addImage({
      path: path.resolve(__dirname, "../images/murgesh.png"),
      type: "picture",
      position: {
        type: "twoCellAnchor",
        from: {
          col: 2,
          row: 2,
        },
        to: {
          col: 3,
          row: 6,
        },
      },
    });
    //logo 2
    const galanfi = worksheet.addImage({
      path: path.resolve(__dirname, "../images/galanfi.png"),
      type: "picture",
      position: {
        type: "twoCellAnchor",
        from: {
          col: 8,
          row: 3,
        },
        to: {
          col: 11,
          row: 5,
        },
      },
    });

    //setting colum width
    worksheet.column(1).setWidth(8);
    worksheet.column(2).setWidth(15);
    worksheet.column(3).setWidth(15);
    worksheet.column(6).setWidth(15);
    worksheet.column(7).setWidth(8);
    worksheet.column(8).setWidth(8);

    //data filling
    //        cell(row,col,row,col)

    worksheet
      .cell(1, 4, 1, 6, true)
      .string(mainHeading)
      .style(mainHeadingStyle);
    worksheet.cell(2, 4, 2, 6, true).string(address).style(style_srNo);
    worksheet.cell(3, 4, 3, 6, true).string(reportName).style(subHeading);

    worksheet.cell(4, 4).string("FROM DATE").style(style_srNo);
    worksheet.cell(4, 5, 4, 6, true).string(fromDate).style(style_srNo);
    worksheet.cell(5, 4).string("TO DATE").style(style_srNo);
    worksheet.cell(5, 5, 5, 6, true).string(toDate).style(style_srNo);
    worksheet.cell(6, 4).string("PARAMETER").style(style_srNo);
    worksheet.cell(6, 5, 6, 6, true).string(parameter).style(style_srNo);

    worksheet
      .cell(7, 4, 7, 6, true)
      .string("")
      .style({
        fill: {
          type: "pattern",
          patternType: "solid",
          fgColor: "#00b0f0",
        },
        border: groupBorder,
      });

    worksheet.cell(8, 1, 8, 10, true).string(reportHeading).style(tableHeading);

    const startrow = 9;

    //makeing columns
    worksheet
      .cell(startrow, 1, startrow + 1, 1, true)
      .string("SR. NO.")
      .style(colHeader);
    worksheet
      .cell(startrow, 2, startrow + 1, 2, true)
      .string("GROUP NAME")
      .style(colHeader);
    worksheet
      .cell(startrow, 3, startrow + 1, 3, true)
      .string("SUB GROUP")
      .style(colHeader);
    worksheet
      .cell(startrow, 4, startrow + 1, 4, true)
      .string("SUB GROUP 1")
      .style(colHeader);
    worksheet
      .cell(startrow, 5, startrow + 1, 5, true)
      .string("SUB GROUP 2")
      .style(colHeader);
    worksheet
      .cell(startrow, 6, startrow + 1, 6, true)
      .string("METERS")
      .style(colHeader);
    worksheet
      .cell(startrow, 7, startrow + 1, 7, true)
      .string("KWH IMP")
      .style(colHeader);
    worksheet
      .cell(startrow, 8, startrow + 1, 8, true)
      .string("KWH EXP")
      .style(colHeader);

    worksheet
      .cell(startrow, 9, startrow, 10, true)
      .string("GROUP TOTAL")
      .style(colHeader);

    worksheet
      .cell(startrow + 1, 9)
      .string("KWH IMPORT")
      .style(colHeader);
    worksheet
      .cell(startrow + 1, 10)
      .string("KWH EXPORT")
      .style(colHeader);

    worksheet
      .cell(startrow + 1, 11)
      .string("METER SID")
      .style(idesStyle);
    worksheet
      .cell(startrow + 1, 12)
      .string("DEVICE IMEI")
      .style(idesStyle);

    let subGroup_count = 0;

    let cellSettings = {
      group_name: {
        start: startrow + 1,
        end: startrow + 1,
        value: "",
      },
      subGroup: {
        start: startrow + 1,
        end: startrow + 1,
        value: "",
      },
      subGroup1: {
        start: startrow + 1,
        end: startrow + 1,
        value: "",
        init: false,
      },
      subGroup2: {
        start: startrow + 1,
        end: startrow + 1,
        value: "",
        init: false,
      },
      kwhImport: {
        start: startrow + 1,
        end: startrow + 1,
        value: 0,
        init: false,
      },
      kwhExport: {
        start: startrow + 1,
        end: startrow + 1,
        value: 0,
        init: false,
      },
    };

    const { group_name, subGroup, subGroup1, subGroup2, kwhImport, kwhExport } =
      cellSettings;

    for (let index = 0; index < records.length; index++) {
      if (typeof records[index].Srno != "undefined") {
        worksheet
          .cell(index + startrow + 2, 1)
          .number(records[index].Srno)
          .style(style_srNo);
      }

      if (typeof records[index].groupNo === "undefined") {
        group_name.end++;
      } else {
        group_name.value = records[index].groupNo;
      }

      if (typeof records[index].subGroup != "undefined") {
        if (subGroup_count > 0) {
          worksheet
            .cell(subGroup.start, 3, subGroup.end, 3, true)
            .string(subGroup.value)
            .style(style_simpleBold);
        }
        subGroup.start = startrow + 1 + index + 1;

        subGroup.value = records[index].subGroup;
        subGroup_count++;
      }
      subGroup.end++;

      if (typeof records[index].subGroup1 != "undefined") {
        if (subGroup1.init) {
          worksheet
            .cell(subGroup1.start, 4, subGroup1.end, 4, true)
            .string(subGroup1.value)
            .style(style_simpleBold);
        }

        subGroup1.start = startrow + 1 + index + 1;
        subGroup1.value = records[index].subGroup1;
        subGroup1.init = true;
      }
      subGroup1.end++;

      if (typeof records[index].subGroup2 != "undefined") {
        if (subGroup2.init) {
          worksheet
            .cell(subGroup2.start, 5, subGroup2.end, 5, true)
            .string(subGroup2.value)
            .style(style_simpleBold);
        }

        subGroup2.start = startrow + 1 + index + 1;
        subGroup2.value = records[index].subGroup2;
        subGroup2.init = true;
      }
      subGroup2.end++;

      worksheet
        .cell(index + startrow + 2, 6)
        .string(records[index].meters)
        .style(style_srNo);

      if (typeof records[index].kwhImp != "undefined") {
        worksheet
          .cell(index + startrow + 2, 7)
          .number(records[index].kwhImp)
          .style(style_srNo);
      } else {
        worksheet
          .cell(index + startrow + 2, 7)
          .string("")
          .style(style_srNo);
      }

      if (typeof records[index].kwhExp != "undefined") {
        worksheet
          .cell(index + startrow + 2, 8)
          .number(records[index].kwhExp)
          .style(style_srNo);
      } else {
        worksheet
          .cell(index + startrow + 2, 8)
          .string("")
          .style(style_srNo);
      }

      if (typeof records[index].kwhImport != "undefined") {
        if (kwhImport.init) {
          worksheet
            .cell(kwhImport.start, 9, kwhImport.end, 9, true)
            .number(kwhImport.value)
            .style(style_srNo);
        }

        kwhImport.start = startrow + 1 + index + 1;
        kwhImport.value = records[index].kwhImport;
        kwhImport.init = true;
      }
      kwhImport.end++;

      if (typeof records[index].kwhExport != "undefined") {
        if (kwhExport.init) {
          worksheet
            .cell(kwhExport.start, 10, kwhExport.end, 10, true)
            .number(kwhExport.value)
            .style(style_srNo);
        }

        kwhExport.start = startrow + 1 + index + 1;
        kwhExport.value = records[index].kwhExport;
        kwhExport.init = true;
      }
      kwhExport.end++;

      if (typeof records[index].meterSid != "undefined") {
        worksheet
          .cell(index + startrow + 2, 11)
          .number(records[index].meterSid)
          .style({ ...style_srNo, border: false });
      }

      worksheet
        .cell(index + startrow + 2, 12)
        .string(records[index].deviceIMEI)
        .style({ ...style_srNo, border: false });

      if (index === records.length - 1) {
        worksheet
          .cell(group_name.start + 1, 2, group_name.end + 1, 2, true)
          .string(group_name.value)
          .style(style_megaBold);

        worksheet
          .cell(subGroup.start, 3, subGroup.end, 3, true)
          .string(subGroup.value)
          .style(style_simpleBold);

        worksheet
          .cell(subGroup1.start, 4, subGroup1.end, 4, true)
          .string(subGroup1.value)
          .style(style_simpleBold);
        worksheet
          .cell(subGroup2.start, 5, subGroup2.end, 5, true)
          .string(subGroup2.value)
          .style(style_simpleBold);

        worksheet
          .cell(kwhImport.start, 9, kwhImport.end, 9, true)
          .string(kwhImport.value)
          .style(style_simpleBold);
        worksheet
          .cell(kwhExport.start, 10, kwhExport.end, 10, true)
          .string(kwhExport.value)
          .style(style_simpleBold);
      }
    }
  }

  res.setHeader("Content-Disposition", "attachment; filename=example.xlsx");

  const buffer = await workbook.writeToBuffer();

  res.status(200).send(buffer);
});

exports.addSheet = catchAsynch(async (req, res, next) => {
  const data = await ReportSchema.create(req.body);

  res.status(200).json({
    msg: "added",
    data,
  });
});

exports.getplantDetail = catchAsynch(async (req, res, next) => {
  const data = await ReportSchema.find(
    {},
    {
      _id: 1,
      "report.mainHeading": 1,
      "report.reportName": 1,
      "report.fromDate": 1,
      "report.toDate": 1,
    }
  );

  res.status(200).json({
    msg: "data loaded",
    data,
  });
});

const fs = require("fs");
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
                    children: [{
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
                          children: [  {
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
                                        children: [{
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
                                        },],
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
                          },],
                          depth: 4,
                          maxChildDepth: 0,
                        },
                      ],
                      depth: 3,
                      maxChildDepth: 1,
                    },],
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





  



];



const filename = "output2.xlsx";

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

exports.abc =  (req, res, next) => {
  createExcelSheet(data, filename);
}
