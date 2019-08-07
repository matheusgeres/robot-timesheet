const Excel = require("exceljs/modern.nodejs");
const moment = require("moment");

const dateNotTyped = "- - : - -";
const formatHour   = "HH:mm";
const formatDate   = "DD/MM/YYYY";

exports.readTimetableFromExcel = async function (fileName, month){
    let workbook = new Excel.Workbook();
    let daysCurrentWeek = weekdays();
    let daysLastWeek = weekdays(true);
    await workbook.xlsx.readFile(fileName);

    let daysToInput = [];
    let lastDayOfMonth = parseInt(
      moment()
      .endOf("month")
      .format("D"));
    lastDayOfMonth = lastDayOfMonth+2;

    workbook.eachSheet(function(worksheet, sheetId) {
      if (worksheet.name == month) {
        for (let pos = 2; pos < lastDayOfMonth; pos++) {
          let date = worksheet.getColumn(1).values[pos];
          let dateFormatted = formatToDate(date);
          // if (daysCurrentWeek.indexOf(dateFormatted) >= 0) {
            let entrance1 = formatToHour(worksheet.getColumn(2).values[pos]);
            let exit1 = formatToHour(worksheet.getColumn(3).values[pos]);
            let entrance2 = formatToHour(worksheet.getColumn(4).values[pos]);
            let exit2 = formatToHour(worksheet.getColumn(8).values[pos].result);
            let narrative = worksheet.getColumn(14).values[pos];
            let clientCode = worksheet.getColumn(15).values[pos];
            let projectCode = worksheet.getColumn(16).values[pos];
            if (entrance1 != undefined && exit2 != "Invalid date" && narrative!=undefined && clientCode!=undefined && projectCode!=undefined) {
              daysToInput.push({
                date: date,
                dateFormatted: dateFormatted,
                entrance1: entrance1,
                exit1: exit1,
                entrance2: entrance2,
                exit2: exit2,
                narrative: narrative,
                clientCode: clientCode.toString().padStart(4, '0'),
                projectCode: projectCode
              });
            } else {
              console.log("Date", dateFormatted);
              if(entrance1 == undefined){
                console.log("Input 1 was not entered.");
              }

              if(exit2 == "Invalid date"){
                console.log("Exit 2 has invalid date");
              }

              if(narrative==undefined){
                console.log("Narrative has not entered");
              }

              if(clientCode==undefined){
                console.log("Client Code has not entered");
              }

              if(projectCode==undefined){
                console.log("Project Code has not entered");
              }
            }
          // }
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