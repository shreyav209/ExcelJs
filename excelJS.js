const Excel = require("exceljs");

async function writeExcelSheet(searchText, replaceText, filePath) {
  const workbook = new Excel.Workbook(); //object of workbook

  //File path
  await workbook.xlsx.readFile(filePath);
  //targte sheet on workbook
  const worksheet = workbook.getWorksheet("Sheet1");

  const output=await readExcel(worksheet, searchText);

  //To write the value on sheet
  const cell = worksheet.getCell(output.row, output.column);
  cell.value = replaceText;
  await workbook.xlsx.writeFile(filePath);
}

async function readExcel(worksheet, searchText) {
  let output = { row: -1, column: -1 };

  worksheet.eachRow((row, rowNumber) => {
    //cell from worksheet
    row.eachCell((cell, colNumber) => {
      if (cell.value == searchText) {
        output.row = rowNumber;
        output.column = colNumber;
      }
    });
  });
  return output
}
writeExcelSheet(
  "Banana",
  "XYZZ",
  "D:/shreya vethekar/playwright/elements/tests/bulkupload/download.xlsx"
);
