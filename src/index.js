const Excel = require("exceljs/modern.nodejs");

(async () => {
  let workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("meuponto_2019.xlsx");

  workbook.eachSheet(function(worksheet, sheetId) {
    if (sheetId == 14) {
      // console.log("worksheet", worksheet);
      console.log(worksheet.getColumn(8).values);
    }
  });

  // console.log("workbook", workbook);
})();
