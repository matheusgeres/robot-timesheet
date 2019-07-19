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

    await page.goto(env.baseUrl);
    await page.type("#login", credentials.username);
    await page.type("#password_sem_md5", credentials.password);
    await Promise.all([
      page.click("#submit"),
      page.waitForNavigation({waitUntil: env.goto.waitUntil})
    ]);

    await page.evaluate((daysToInput) => { editHora('08:00', '', daysToInput[0].dateFormatted, '') }, daysToInput);
    await page.waitFor(".ui-dialog");

    let clientCode = "0033";
    await page.type("#codcliente_form_lanctos", clientCode);
    await page.evaluate((clientCode) => { getCodClientePrj(clientCode,'','cadastro_time_despesa') }, clientCode);

    let projectCode = "KC2068";
    await page.type("#codprojeto_form_lanctos", projectCode);
    await page.evaluate((projectCode) => { getCodCliProjeto(projectCode,'set_dados_lanctos','cadastro_time_despesa') }, projectCode);
    await page.waitForResponse("https://timesheet.keyrus.com.br/includes/ajax_calls/get_dadosAtividades.ajax.php");

    await page.on("dialog", (dialog) => { dialog.accept(); });

    for(let di of daysToInput){
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
      await page.waitForResponse("https://timesheet.keyrus.com.br/includes/ajax_calls/saveLanctos.ajax.php");
    }

    // await browser.close();
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
          if (daysCurrentWeek.indexOf(dateFormatted) >= 0) {
            let entrance1 = formatToHour(worksheet.getColumn(2).values[pos]);
            let exit1 = formatToHour(worksheet.getColumn(3).values[pos]);
            let entrance2 = formatToHour(worksheet.getColumn(4).values[pos]);
            let exit2 = formatToHour(worksheet.getColumn(8).values[pos].result);
            let narrative = worksheet.getColumn(14).values[pos]
            if (entrance1 != undefined && exit2 != "Invalid date") {
              daysToInput.push({
                date: date,
                dateFormatted: dateFormatted,
                entrance1: entrance1,
                exit1: exit1,
                entrance2: entrance2,
                exit2: exit2,
                narrative: narrative
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
