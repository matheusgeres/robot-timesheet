const readline = require("readline-sync");
const excel    = require('./robot/excel');
const chrome   = require('./robot/chrome');

(async () => {
  let month = "AGO";
  let fileName = "meuponto_2019.xlsx";

  let periodRead = parseInt(readline.question(
    `Do you want read entire timesheet, current week or last week? \nUse one of these options:
    1 = Entire Timesheet
    2 = Current week
    3 = Last week\n`));

  let periodsToReadValid = [1,2,3];
  if(periodsToReadValid.indexOf(periodRead)<0){
    console.log("You enter not valid period. Try again, please! :D");
    return;
  }
  
  let daysToInput = await excel.readTimetableFromExcel(fileName, month, periodRead);
  console.log("\nDate with hours to input")
  console.table(daysToInput);

  let doInput = readline.question("Do you want continue with input of hours? [y/N]");
  if(doInput.toLowerCase()=="y"){
    await chrome.inputHoursOnTimesheet(daysToInput);
  }
})();
