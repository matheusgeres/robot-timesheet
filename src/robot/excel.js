const Excel = require("exceljs/modern.nodejs");
const moment = require("moment");

const dateNotTyped = "- - : - -";
const formatHour = "HH:mm";
const formatDate = "DD/MM/YYYY";
const twoHours = 120;
const hourZero = "00:00";

exports.readTimetableFromExcel = async function (fileName, month, periodRead, overTimeOption){
  let workbook = new Excel.Workbook();
  let daysCurrentWeek = weekdays();
  let daysLastWeek = weekdays(true);
  let daysFilter = [];

  switch (periodRead) {
    case 1:
      break;
    case 2: 
      daysFilter = daysCurrentWeek;
      break;
    case 3:
      daysFilter = daysLastWeek;
      break;
  }

  await workbook.xlsx.readFile(fileName);

  let daysToInput = [];
  let lastDayOfMonth = parseInt(
    moment()
    .endOf("month")
    .format("D"));
  lastDayOfMonth = lastDayOfMonth+2;

  workbook.eachSheet(function(worksheet, sheetId) {
    if (worksheet.name == month) {
      let dateInput = {};
      let dateWithErrors = [];
      for (let pos = 2; pos < lastDayOfMonth; pos++) {
        let date = worksheet.getColumn(1).values[pos];
        let dateFormatted = formatToDate(date);
        if (daysFilter.indexOf(dateFormatted) >= 0 || periodRead == 1) {
          dateInput = {
            date: date,
            dateFormatted: formatToDate(date),
            entrance1: formatToHour(worksheet.getColumn(2).values[pos]),
            exit1: formatToHour(worksheet.getColumn(3).values[pos]),
            entrance2: formatToHour(worksheet.getColumn(4).values[pos]),
            exit2: formatToHour(worksheet.getColumn(8).values[pos]),
            narrative: worksheet.getColumn(14).values[pos],
            clientCode: worksheet.getColumn(15).values[pos],
            projectCode: worksheet.getColumn(16).values[pos]
          }
          addDaysToInput(daysToInput, dateInput, dateWithErrors);

          if (overTimeOption == 1) {
            dateInput.overTime = formatToExtraHour(worksheet.getColumn(12).values[pos]);
            dateInput.suggestedExit = worksheet.getColumn(8).values[pos];
            dateInput.dateExit2 = worksheet.getColumn(5).values[pos];
            addOvertime(daysToInput, dateInput, dateWithErrors);
          }
        }
      }
      if(dateWithErrors.length>0){
        console.log("\nDates with errors");
        console.table(dateWithErrors);
      }
    }
  });

  return daysToInput;
}

function addOvertime(daysToInput, dateInput, dateWithErrors) {
  if (dateInput.overTime > 0) {
    if (dateInput.overTime <= twoHours) {
      let exitOver = formatToHour(dateInput.dateExit2);
      let dateInputOvertime = changeEntranceAndExit(dateInput, dateInput.exit2, exitOver);
      addDaysToInput(daysToInput, dateInputOvertime, dateWithErrors);
    } else if (dateInput.overTime > twoHours) {
      let exitOver1 = formatToHourAddMinute(dateInput.suggestedExit.result, twoHours);
      let dateInputOvertime1 = changeEntranceAndExit(dateInput, dateInput.exit2, exitOver1);
      addDaysToInput(daysToInput, dateInputOvertime1, dateWithErrors);

      let exitOver2 = formatToHour(dateInput.dateExit2);
      let dateInputOvertime2 = changeEntranceAndExit(dateInput, exitOver1, exitOver2);
      addDaysToInput(daysToInput, dateInputOvertime2, dateWithErrors);
    }
  }
}

function changeEntranceAndExit(dateInput, entrance, exit){
  let dateInputOvertime = Object.assign({}, dateInput);
  dateInputOvertime.entrance1 = entrance;
  dateInputOvertime.exit1 = hourZero;
  dateInputOvertime.entrance2 = hourZero;
  dateInputOvertime.exit2 = exit;
  return dateInputOvertime;
}

function addDaysToInput(daysToInput, dateInput, dateWithErrors) {
  if (dateInput.entrance1 != undefined && dateInput.exit2 != "Invalid date" && dateInput.narrative != undefined 
    && dateInput.clientCode != undefined && dateInput.projectCode != undefined) {
    daysToInput.push({
      date: dateInput.date,
      dateFormatted: dateInput.dateFormatted,
      entrance1: dateInput.entrance1,
      exit1: dateInput.exit1,
      entrance2: dateInput.entrance2,
      exit2: dateInput.exit2,
      narrative: dateInput.narrative,
      clientCode: dateInput.clientCode.toString().padStart(4, '0'),
      projectCode: dateInput.projectCode.toString()
    });
  }
  else {
    let dateError = { date: dateInput.dateFormatted };
    if (dateInput.entrance1 == undefined) {
      dateError.entrance1 = "Input 1 was not entered.";
    }
    if (dateInput.exit2 == "Invalid date") {
      dateError.exit2 = "Exit 2 has invalid date";
    }
    if (dateInput.narrative == undefined) {
      dateError.narrative = "Narrative has not entered";
    }
    if (dateInput.clientCode == undefined) {
      dateError.clientCode = "Client Code has not entered";
    }
    if (dateInput.projectCode == undefined) {
      dateError.projectCode = "Project Code has not entered";
    }
    dateWithErrors.push(dateError);
  }
}

function formatToHourAddMinute(columnValues, minutes) {
  if (columnValues == dateNotTyped) return undefined;
  return moment(columnValues)
    .add(minutes, 'minutes')
    .utc()
    .format(formatHour);
}

function formatToHour(columnValues) {
  if (columnValues == dateNotTyped || columnValues == undefined) return undefined;
  if (columnValues.hasOwnProperty('result')) columnValues = columnValues.result;

  return moment(columnValues)
    .utc()
    .format(formatHour);
}

function formatToExtraHour(columnValues) {
  if (columnValues == dateNotTyped){
    return undefined;
  }

  if(columnValues == undefined){
    return 0;
  }

  let hour = formatToHour(columnValues);
  let minute = moment.duration(hour).asMinutes();

  return minute>480?0:minute;
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