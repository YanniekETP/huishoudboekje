/** ET WebApp Starter. Configureer dit per nieuwe app. */
const SPREADSHEET_ID = 'REPLACE_ME_SPREADSHEET_ID';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ET WebApp')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function healthcheck() {
  return { ok: true, version: 'starter-1', timestamp: new Date().toISOString() };
}

function getDashboardData() {
  if (SPREADSHEET_ID === 'REPLACE_ME_SPREADSHEET_ID') {
    throw new Error('Configureer eerst de spreadsheet-ID.');
  }
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return { title: spreadsheet.getName(), sheets: spreadsheet.getSheets().map(s => s.getName()) };
}
