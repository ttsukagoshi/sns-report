// MIT License
// 
// Copyright (c) 2020 Taro TSUKAGOSHI
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

// Dependencies on the manifest file 'appsscript.json'
// "dependencies": {
//   "libraries": [{
//     "userSymbol": "OAuth2",
//     "libraryId": "1B7FSrk5Zi6L1rSxxTDgDEUsPzlukDsi4KGuTMorsTQHhGBzBkMun4iDF",
//     "version": "38"
//   }]
// }
// See https://github.com/gsuitedevs/apps-script-oauth2

const LOG_SHEET_NAME = '99_Log';

/**
 * onOpen()
 * Add menu to spreadsheet
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('YouTube')
    .addItem('Authorize', 'showSidebarYouTubeApi')
    .addItem('Logout/Reset', 'logoutYouTube')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Update Channel/Video List')
        .addItem('Update All', 'updateAllYouTubeList')
        .addSeparator()
        .addItem('Update Channel List', 'updateYouTubeSummaryChannelList')
        .addItem('Update Video List', 'updateYouTubeSummaryVideoList')
    )
    .addSubMenu(
      ui.createMenu('Analytics')
        .addItem('Get Latest Data', 'updateYouTubeAnalyticsData')
    )
    .addSeparator()
    .addItem('Create Analytics Summary', 'createYouTubeAnalyticsSummary')
    .addToUi();
  ui.createMenu('Facebook')
    .addItem('Authorize', 'showSidebarFacebookApi')
    .addItem('Logout/Reset', 'logoutFacebook')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('Update Page List')
        .addItem('Update All', 'updateAllFbList')
        .addSeparator()
        .addItem('Update Page List', 'updateFbSummaryPageList')
    )
    .addToUi();
  ui.createMenu('Settings')
    .addItem('Setup', 'initialSettings')
    .addItem('Check Settings', 'checkSettings')
    .addToUi();
  ui.createMenu('test')/////////////////////////////////////////////////////
    .addItem('test', 'test')
    .addItem('catchup', 'archiveCatchup')
    .addToUi()
}

/////////////////////////////
// Configurations and Misc //
/////////////////////////////

/**
 * Standarized error message
 * @param {Object} e Error object
 * @return {string} Standarized error message
 */
function errorMessage_(e) {
  let message = `Error: line - ${e.lineNumber}\n${e.stack}`
  return message;
}

/**
 * Returns an object of configurations from the current spreadsheet.
 * @param {string} configSheetName Name of sheet with configurations. Defaults to 'Config'.
 * @return {Object}
 * The sheet should have a first row of headers, with the keys (properties) in its first column and values in the second.
 */
function getConfig_(configSheetName = 'Config') {
  // Get values from spreadsheet
  var configValues = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(configSheetName).getDataRange().getValues();
  configValues.shift();

  // Convert the 2d array values into a Javascript object
  var configObj = {};
  configValues.forEach(element => configObj[element[0]] = element[1]);
  return configObj;
}

/**
 * Get the list of year-by-year Spreadsheet URLs in an array of Javascript objects
 * @param {string} sheetName Name of sheet that the URLs are listed on.
 * @returns {array}
 */
function getSpreadsheetList_(sheetName) {
  var spreadsheetList = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange().getValues();
  var header = spreadsheetList.shift();
  // Convert the 2d array into an array of row-by-row Javascript objects
  var spreadsheetListObj = spreadsheetList.map(function (value) {
    return header.reduce(function (obj, key, ind) {
      obj[key] = value[ind];
      return obj;
    }, {});
  });
  return spreadsheetListObj;
}

/**
 * Returns a formatted date
 * @param {Date} date Date object to format
 * @param {string} timeZone [Optional] Time zone; defaults to the script's time zone. https://developers.google.com/apps-script/reference/base/session#getScriptTimeZone()
 * @returns {string} Formatted date string
 */
function formattedDate_(date, timeZone = Session.getScriptTimeZone()) {
  var dateString = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd HH:mm:ss Z');
  return dateString;
}

/**
 * Checks whether the given date belongs to the given year. Timezone can be designated.
 * @param {string} dateString Date string in ISO8601 format to be checked.
 * @param {number} checkYear The year to check whether the dateString belongs to.
 * @param {string} timeZone [Optional] Timezone in which to conduct the check. Defaults to the scripts' timezone.
 * @returns {boolean} Returns true if the given date belongs to the given year.
 */
function checkYear_(dateString, checkYear, timeZone = Session.getScriptTimeZone()) {
  var dateObj = new Date(dateString);
  var year = parseInt(formattedDate_(dateObj, timeZone).slice(0, 4));
  return year == checkYear;
}

/**
 * Gets the URL of the target spreadsheet from the list of spreadsheet names and their URLs.
 * @param {string} spreadsheetListName Sheet name of the spreadsheet list
 * @param {number} targetYear Year of the spreadsheet
 * @param {string} platform SNS platform
 * @param {Object} options [Optional] Parameters to automatically create new file(s).
 * @param {boolean} options.createNewFile [Optional] Automatically creates a new file if no existing file matches the conditions; defaults to false.
 * @param {string} options.driveFolderId [Optional] The Google Drive folder ID in which to create the new file.
 * If this property is not supplied, the new file will be created under the root folder.
 * @param {string} options.templateFileId [Optional; required if options.createNewFile = true] The ID of the Google Spreadsheet to use as template for the new file.
 * @param {string} options.newFileName [Optional] Name of the new file. If not supplied, the new file will have the same name with the template file.
 * @param {string} options.newFileNamePrefix [Optional] Prefix for the new file name.
 * @param {string} options.newFileNameSuffix [Optional] Suffix for the new file name.
 * @returns {object} JavaScript object with the following properties:
 * {string} url - URL of the spreadsheet. Null if no matching spreadsheet is available and options.createNewFile is set to false (which is the default value).
 * {boolean} created - True when a new spreadsheet is created.
 */
function spreadsheetUrl_(spreadsheetListName, targetYear, platform, options = {}) {
  var optionsDefault = {
    createNewFile: false,
    driveFolderId: '',
    templateFileId: '',
    newFileName: '',
    newFileNamePrefix: '',
    newFileNameSuffix: ''
  };
  for (let k in optionsDefault) {
    if (!options[k]) {
      options[k] = optionsDefault[k];
    }
  }
  var spreadsheetList = getSpreadsheetList_(spreadsheetListName);
  var targetSpreadsheet = spreadsheetList.filter(value => (value.YEAR == targetYear && value.PLATFORM == platform));
  try {
    if (targetSpreadsheet.length) {
      return { 'url': targetSpreadsheet[0].URL, 'created': false };
    } else if (options.createNewFile && options.templateFileId) {
      let targetFolder = (options.driveFolderId ? DriveApp.getFolderById(options.driveFolderId) : DriveApp.getRootFolder());
      let templateFile = DriveApp.getFileById(options.templateFileId);
      let newFileName = (options.newFileName ? options.newFileName : templateFile.getName());
      newFileName = options.newFileNamePrefix + newFileName + options.newFileNameSuffix;
      // Copy template into designated Drive folder
      // https://developers.google.com/apps-script/reference/drive/file#makeCopy(String,Folder)
      let newFileUrl = templateFile.makeCopy(newFileName, targetFolder).getUrl();
      // Add the new file to the spreadsheet list
      let newRow = [targetYear, platform, newFileName, newFileUrl];
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(spreadsheetListName).appendRow(newRow);
      // Add log to the new file
      enterLog_(SpreadsheetApp.openByUrl(newFileUrl).getId(), LOG_SHEET_NAME, 'Spreadsheet created.');
      return { 'url': newFileUrl, 'created': true };
    } else if (options.createNewFile && !options.templateFileId) {
      throw new Error('The key "templateFileId" is missing in "options".');
    } else {
      return { 'url': null, 'created': false };
    }
  } catch (error) {
    let message = errorMessage_(error);
    throw new Error(message);
  }
}

/**
 * Enter log on designated spreadsheet.
 * @param {string} spreadsheetId ID of spreadsheet to enter log.
 * @param {string} sheetName Name of log sheet.
 * @param {string} logMessage Log message
 * @param {Date} timestamp [Optional] Date object for timestamp. If omitted, enters the time of execution.
 */
function enterLog_(spreadsheetId, sheetName, logMessage, timestamp = new Date()) {
  var targetSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  var timestampString = formattedDate_(timestamp, targetSpreadsheet.getSpreadsheetTimeZone());
  targetSpreadsheet.getSheetByName(sheetName).appendRow([timestampString, logMessage]);
}