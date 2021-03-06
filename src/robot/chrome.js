const puppeteer = require("puppeteer");
const env = require("../../local.env.json")
const credentials = require("../../credentials/timesheet.json");
const moment = require("moment");

const currentTimeFormat = 'YYYYMMDDHHmm';
const formatTimesheet = 'YYYY-MM-DD';

exports.inputHoursOnTimesheet = async function(daysToInput, doPrintscreen){
  const browser      = await puppeteer.launch({
    headless         : env.puppeteer.headless,
    ignoreHTTPSErrors: env.puppeteer.ignoreHTTPSErrors,
    args             : env.puppeteer.args
  });

  const pages = await browser.pages();
  const page  = pages[0];
  await page.setViewport({width: env.puppeteer.viewPort.width, height: env.puppeteer.viewPort.height});

  await page.goto(`${env.baseUrl}/login.php`);
  await page.type("#login", credentials.username);
  await page.type("#password_sem_md5", credentials.password);
  await Promise.all([
    page.click("#submit"),
    page.waitForNavigation({waitUntil: env.goto.waitUntil})
  ]);

  await page.goto(`${env.baseUrl}/timesheet/multidados/module/calendario/index.php`, { waitUntil: env.goto.waitUntil });
  await page.evaluate((daysToInput) => { editHora('08:00', '', daysToInput[0].dateFormatted, '') }, daysToInput);
  await page.waitFor(".ui-dialog");

  await page.on("dialog", (dialog) => { console.log(dialog.message()); dialog.accept(); });

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

  if(doPrintscreen){
    let firstDate = getFormattedTimesheet(daysToInput[0].date);
    let currentTime = getCurrentTimeFormatted(moment());

    await page.goto(`${env.baseUrl}/timesheet/multidados/module/calendario/index.php?DATA=${firstDate}&cmd=listar&SHOW=mes`, { waitUntil: env.goto.waitUntil });
    await page.screenshot({path: `inputHoursTimesheet_${currentTime}.png`});
  }

  await browser.close();
}

function getCurrentTimeFormatted(moment){
  return moment.format(currentTimeFormat);
}

function getFormattedTimesheet(date){
  return moment(date)
    .add(1, "days")
    .format(formatTimesheet);
}