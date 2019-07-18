const Excel     = require("exceljs/modern.nodejs");
const moment    = require("moment");
const puppeteer = require("puppeteer");

const dateNotTyped = "- - : - -";
const formatHour   = "HH:mm";
const formatDate   = "DD-MM-YYYY";

(async () => {
  let month = "JUL";
  let fileName = "meuponto_2019.xlsx";
  
  let daysToInput = await readTimetableFromExcel(fileName, month);
  console.log(daysToInput);

  async function readTimetableFromExcel(fileName, month){
    let workbook = new Excel.Workbook();
    let daysCurrentWeek = weekdays();
    let daysLastWeek = weekdays(true);
    await workbook.xlsx.readFile(fileName);

    let daysToInput = [];
    let lastDayOfMonth = moment()
      .endOf("month")
      .format("D");

    workbook.eachSheet(function(worksheet, sheetId) {
      if (worksheet.name == month) {
        for (let pos = 2; pos < lastDayOfMonth; pos++) {
          let date = worksheet.getColumn(1).values[pos];
          let dateFormatted = formatToDate(date);
          if (daysLastWeek.indexOf(dateFormatted) >= 0) {
            let entrance1 = formatToHour(worksheet.getColumn(2).values[pos]);
            let exit1 = formatToHour(worksheet.getColumn(3).values[pos]);
            let entrance2 = formatToHour(worksheet.getColumn(4).values[pos]);
            let exit2 = formatToHour(worksheet.getColumn(8).values[pos].result);
            if (entrance1 != undefined && exit2 != "Invalid date") {
              daysToInput.push({
                date: date,
                dateFormatted: dateFormatted,
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

    return daysToInput;
  }

  function formatToHour(columnValues) {
    if (columnValues == dateNotTyped) return undefined;
    return moment(columnValues)
      .utc()
      .format(formatHour);
  }

  function formatToDate(columnsValues) {
    return moment(columnsValues)
      .add(1, "days")
      .format(formatDate);
  }

  function weekdays(isLastWeek) {
    let days = [];
    for (let index = 1; index <= 5; index++) {
      let day = moment().day(index);
      if(isLastWeek) day.add(-1, "weeks");
      days.push(day.format(formatDate));
    }
    return days;
  }
})();
