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
      let dateWithErrors = [];
      for (let pos = 2; pos < lastDayOfMonth; pos++) {
        let date = worksheet.getColumn(1).values[pos];
        let dateFormatted = formatToDate(date);
        if (daysFilter.indexOf(dateFormatted) >= 0 || periodRead == 1) {
          let entrance1 = formatToHour(worksheet.getColumn(2).values[pos]);
          let exit1 = formatToHour(worksheet.getColumn(3).values[pos]);
          let entrance2 = formatToHour(worksheet.getColumn(4).values[pos]);
          let exit2 = formatToHour(worksheet.getColumn(8).values[pos].result);
          let narrative = worksheet.getColumn(14).values[pos];
          let clientCode = worksheet.getColumn(15).values[pos];
          let projectCode = worksheet.getColumn(16).values[pos];
          let overTime = formatToExtraHour(worksheet.getColumn(12).values[pos]);
          addDaysToInput(date, dateFormatted, entrance1 , exit1, entrance2, exit2, narrative, clientCode, projectCode, daysToInput, dateWithErrors);

          if(overTimeOption==1){
            if(overTime>0){
              if(overTime<=twoHours){
                let exitOver = formatToHour(worksheet.getColumn(5).values[pos]);
                addDaysToInput(date, dateFormatted, exit2 , hourZero, hourZero, exitOver, narrative, clientCode, projectCode, daysToInput, dateWithErrors);
              }else if(overTime>twoHours){
                let exitOver1 = formatToHourAddMinute(worksheet.getColumn(8).values[pos].result, twoHours);
                let exitOver2 = formatToHour(worksheet.getColumn(5).values[pos]);
  
                addDaysToInput(date, dateFormatted, exit2 , hourZero, hourZero, exitOver1, narrative, clientCode, projectCode, daysToInput, dateWithErrors);
                addDaysToInput(date, dateFormatted, exitOver1 , hourZero, hourZero, exitOver2, narrative, clientCode, projectCode, daysToInput, dateWithErrors);
              }
            }
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

function addDaysToInput(date, dateFormatted, entrance1 , exit1, entrance2, exit2, narrative, clientCode, projectCode, daysToInput, dateWithErrors) {
  if (entrance1 != undefined && exit2 != "Invalid date" && narrative != undefined && clientCode != undefined && projectCode != undefined) {
    daysToInput.push({
      date: date,
      dateFormatted: dateFormatted,
      entrance1: entrance1,
      exit1: exit1,
      entrance2: entrance2,
      exit2: exit2,
      narrative: narrative,
      clientCode: clientCode.toString().padStart(4, '0'),
      projectCode: projectCode.toString()
    });
  }
  else {
    let dateError = { date: dateFormatted };
    if (entrance1 == undefined) {
      dateError.entrance1 = "Input 1 was not entered.";
    }
    if (exit2 == "Invalid date") {
      dateError.exit2 = "Exit 2 has invalid date";
    }
    if (narrative == undefined) {
      dateError.narrative = "Narrative has not entered";
    }
    if (clientCode == undefined) {
      dateError.clientCode = "Client Code has not entered";
    }
    if (projectCode == undefined) {
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
  if (columnValues == dateNotTyped) return undefined;
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