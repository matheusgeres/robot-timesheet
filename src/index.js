const Excel       = require("exceljs/modern.nodejs");
const moment      = require("moment");
const puppeteer   = require("puppeteer");
const env         = require("../local.env")
const credentials = require("../credentials");

const dateNotTyped = "- - : - -";
const formatHour   = "HH:mm";
const formatDate   = "DD/MM/YYYY";

(async () => {
  let month = "JUL";
  let fileName = "meuponto_2019.xlsx";
  
  let daysToInput = await readTimetableFromExcel(fileName, month);
  console.log(daysToInput);

  inputHoursOnTimesheet(daysToInput);

  async function inputHoursOnTimesheet(daysToInput){
    const browser      = await puppeteer.launch({
      headless         : env.puppeteer.headless,
      ignoreHTTPSErrors: env.puppeteer.ignoreHTTPSErrors
    });

    const page = await browser.newPage();
    await page.setViewport({width: env.puppeteer.viewPort.width, height: env.puppeteer.viewPort.height});

    await page.goto(`${env.baseUrl}/login.php`);
    await page.type("#login", credentials.username);
    await page.type("#password_sem_md5", credentials.password);
    await Promise.all([
      page.click("#submit"),
      page.waitForNavigation({waitUntil: env.goto.waitUntil})
    ]);

    await page.evaluate((daysToInput) => { editHora('08:00', '', daysToInput[0].dateFormatted, '') }, daysToInput);
    await page.waitFor(".ui-dialog");

    await page.on("dialog", (dialog) => { dialog.accept(); });

    for(let di of daysToInput){
      let selectorClient = "#codcliente_form_lanctos";
      await page.click(selectorClient, {clickCount: 3});
      await page.type(selectorClient, di.clientCode);
      await page.evaluate((clientCode) => { getCodClientePrj(clientCode,'','cadastro_time_despesa') }, di.clientCode);
  
      let selectProject = "#codprojeto_form_lanctos";
      await page.click(selectProject, {clickCount: 3});
      await page.type(selectProject, di.projectCode);
      await page.evaluate((projectCode) => { getCodCliProjeto(projectCode,'set_dados_lanctos','cadastro_time_despesa') }, di.projectCode);
      await page.waitForResponse(`${env.baseUrl}/includes/ajax_calls/get_dadosAtividades.ajax.php`);

      let selectorDate = "#f_data_b";
      await page.click(selectorDate, {clickCount: 3});
      await page.type(selectorDate, di.dateFormatted);
  
      await page.type("#hora", di.entrance1);
      await page.type("#intervalo_hr_inicial", di.exit1);
      await page.type("#intervalo_hr_final", di.entrance2);
      await page.type("#hora_fim", di.exit2);
  
      let selectorNarrative = "#narrativa_principal";
      await page.click(selectorNarrative, {clickCount: 3});
      await page.type(selectorNarrative, di.narrative);

      await page.click("div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(2)");
      await page.waitForResponse(`${env.baseUrl}/includes/ajax_calls/saveLanctos.ajax.php`);
    }

    await browser.close();
  }
  
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
