const readline    = require("readline-sync");
const excel = require('./robot/excel');
const chrome = require('./robot/chrome');

(async () => {
  let month = "AGO";
  let fileName = "meuponto_2019.xlsx";
  
  let daysToInput = await excel.readTimetableFromExcel(fileName, month);
  console.log(daysToInput);

  let doInput = readline.question("Do you want continue with input of hours? [y/N]");
  if(doInput.toLowerCase()=="y"){
    chrome.inputHoursOnTimesheet(daysToInput);
  }
})();
