/* global getConfig_, getSpreadsheetList_, LocalizedMessage */

function getData() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  var config = getConfig_();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);

}