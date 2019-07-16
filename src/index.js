const Excel = require("exceljs/modern.nodejs");
const moment = require("moment");
const dateNotTyped = "- - : - -";

(async () => {
  let month = "JUL";
  let workbook = new Excel.Workbook();
  let daysCurrentWeek = currentWeek();
  let daysLastWeek = lastWeek();
  await workbook.xlsx.readFile("meuponto_2019.xlsx");

  let daysToInput = [];
  workbook.eachSheet(function(worksheet, sheetId) {
    if (worksheet.name == month) {
      for (
        let pos = 2;
        pos <
        moment()
          .endOf("month")
          .format("D");
        pos++
      ) {
        let date = formatDate(worksheet.getColumn(1).values[pos]);
        if (daysLastWeek.indexOf(date) >= 0) {
          let entrance1 = formatHour(worksheet.getColumn(2).values[pos]);
          let exit1 = formatHour(worksheet.getColumn(3).values[pos]);
          let entrance2 = formatHour(worksheet.getColumn(4).values[pos]);
          let exit2 = formatHour(worksheet.getColumn(8).values[pos].result);
          if (entrance1 != undefined) {
            daysToInput.push({
              date: date,
              entrance1: entrance1,
              exit1: exit1,
              entrance2: entrance2,
              exit2: exit2
            });
          }
        }
      }
    }
  });

  console.log(daysToInput);

  function formatHour(columnValues) {
    if (columnValues == dateNotTyped) return undefined;
    return moment(columnValues)
      .utc()
      .format("HH:mm");
  }

  function formatDate(columnsValues) {
    return moment(columnsValues)
      .add(1, "days")
      .format("YYYY-MM-DD");
  }

  function lastWeek() {
    let days = [];
    for (let index = 1; index <= 5; index++) {
      days.push(
        moment()
          .add(-7, "days")
          .day(index)
          .format("YYYY-MM-DD")
      );
    }
    return days;
  }

  function currentWeek() {
    let days = [];
    for (let index = 1; index <= 5; index++) {
      days.push(
        moment()
          .day(index)
          .format("YYYY-MM-DD")
      );
    }
    return days;
  }

  // console.log("workbook", workbook);
})();
