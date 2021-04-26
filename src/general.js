// Copyright 2021 Taro TSUKAGOSHI
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// 
// For latest information, see https://github.com/ttsukagoshi/sns-report

/* exported checkSettings, checkYear_, errorMessage_, getConfig_, initialSettings, onOpen, spreadsheetUrl_, weeklyAnalyticsUpdate */
/* global LocalizedMessage, updateYouTubeAnalyticsData */

const LOG_SHEET_NAME = '99_Log';
const GITHUB_URL = 'https://github.com/ttsukagoshi/sns-report';

/**
 * onOpen()
 * Add menu to spreadsheet
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale()); // See locale.js for the list of localized messages
  ui.createMenu('YouTube')
    .addItem(localizedMessages.messageList.general.index.authorize, 'showSidebarYouTubeApi')
    .addItem(localizedMessages.messageList.general.index.logoutReset, 'logoutYouTube')
    .addSeparator()
    .addSubMenu(
      ui.createMenu(localizedMessages.messageList.general.index.updateChannelVideoList)
        .addItem(localizedMessages.messageList.general.index.updateAll, 'updateAllYouTubeList')
        .addSeparator()
        .addItem(localizedMessages.messageList.general.index.updateChannelList, 'updateYouTubeSummaryChannelList')
        .addItem(localizedMessages.messageList.general.index.updateVideoList, 'updateYouTubeSummaryVideoList')
    )
    .addSubMenu(
      ui.createMenu(localizedMessages.messageList.general.index.analytics)
        .addItem(localizedMessages.messageList.general.index.getLatestData, 'updateYouTubeAnalyticsData')
    )
    .addSeparator()
    .addItem(localizedMessages.messageList.general.index.createAnalyticsSummary, 'createYouTubeAnalyticsSummary')
    .addToUi();
  ui.createMenu('Facebook')
    .addItem(localizedMessages.messageList.general.index.authorize, 'showSidebarFacebookApi')
    .addItem(localizedMessages.messageList.general.index.logoutReset, 'logoutFacebook')
    .addSeparator()
    .addSubMenu(
      ui.createMenu(localizedMessages.messageList.general.index.updatePageList)
        .addItem(localizedMessages.messageList.general.index.updateAll, 'updateAllFbList')
        .addSeparator()
        .addItem(localizedMessages.messageList.general.index.updatePageList, 'updateFbSummaryPageList')
    )
    .addToUi();
  ui.createMenu(localizedMessages.messageList.general.index.settings)
    .addItem(localizedMessages.messageList.general.index.setup, 'initialSettings')
    .addItem(localizedMessages.messageList.general.index.checkSettings, 'checkSettings')
    .addToUi();
  ui.createMenu('test')/////////////////////////////////////////////////////
    .addItem('test', 'test')
    .addItem('catchup', 'archiveCatchup')
    .addToUi();
}

//////////////
// Settings //
//////////////

/**
 * Prompt to set intial settings
 */
 function initialSettings() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  if (!scriptProperties.setupComplete || !toBoolean_(scriptProperties.setupComplete)) {
    setup_(ui);
  } else {
    let alreadySetup = 'Initial settings are already complete. Do you want to overwrite the settings?\n\n';
    for (let k in scriptProperties) {
      alreadySetup += `${k}: ${scriptProperties[k]}\n`;
    }
    let response = ui.alert(alreadySetup, ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      setup_(ui, scriptProperties);
    }
  }
}

/**
 * Shows the list of current script properties.
 */
function checkSettings() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var currentSettings = '';
  for (let k in scriptProperties) {
    currentSettings += `${k}: ${scriptProperties[k]}\n`;
  }
  ui.alert(currentSettings);
}

/**
 * Convert string booleans into boolean data
 * @param {string} stringBoolean 
 * @return {boolean}
 */
function toBoolean_(stringBoolean) {
  return stringBoolean.toLowerCase() === 'true';
}


/**
 * Sets the required script properties
 * @param {Object} ui Apps Script Ui class object, as retrieved by SpreadsheetApp.getUi()
 * @param {Object} currentSettings [Optional] Current script properties
 */
function setup_(ui, currentSettings = {}) {
  try {
    // YouTube Client ID
    let promptYtClientId = 'YouTube Client ID: OAuth 2.0 client ID for this script/spreadsheet. See https://developers.google.com/youtube/reporting/guides/registering_an_application';
    promptYtClientId += (currentSettings.ytClientId ? `\n\nCurrent Value: ${currentSettings.ytClientId}` : '');
    let responseYtClientId = ui.prompt(promptYtClientId, ui.ButtonSet.OK_CANCEL);
    if (responseYtClientId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let ytClientId = responseYtClientId.getResponseText();
    // YouTube Client Secret
    let promptYtClientSecret = 'YouTube Client Secret: OAuth 2.0 client secret for this script/spreadsheet. See https://developers.google.com/youtube/reporting/guides/registering_an_application';
    promptYtClientSecret += (currentSettings.ytClientSecret ? `\n\nCurrent Value: ${currentSettings.ytClientSecret}` : '');
    let responseYtClientSecret = ui.prompt(promptYtClientSecret, ui.ButtonSet.OK_CANCEL);
    if (responseYtClientSecret.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let ytClientSecret = responseYtClientSecret.getResponseText();
    // Facebook App ID
    let promptFbAppId = 'Facebook App ID: ID of Facebook App to process request. See https://developers.facebook.com/docs/facebook-login/access-tokens/';
    promptFbAppId += (currentSettings.fbAppId ? `\n\nCurrent Value: ${currentSettings.fbAppId}` : '');
    let responseFbAppId = ui.prompt(promptFbAppId, ui.ButtonSet.OK_CANCEL);
    if (responseFbAppId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fbAppId = responseFbAppId.getResponseText();
    // Facebook App ID
    let promptFbAppSecret = 'Facebook App Secret: Facebook App Secret to process request. See https://developers.facebook.com/docs/facebook-login/access-tokens/';
    promptFbAppSecret += (currentSettings.fbAppSecret ? `\n\nCurrent Value: ${currentSettings.fbAppSecret}` : '');
    let responseFbAppSecret = ui.prompt(promptFbAppSecret, ui.ButtonSet.OK_CANCEL);
    if (responseFbAppSecret.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fbAppSecret = responseFbAppId.getResponseText();
    // Drive Folder ID
    let promptDriveFolderId = 'Drive Folder ID: The workspace Google Drive folder ID that all analytics spreadsheets are to be stored in.';
    promptDriveFolderId += (currentSettings.driveFolderId ? `\n\nCurrent Value: ${currentSettings.driveFolderId}` : '');
    let responseDriveFolderId = ui.prompt(promptDriveFolderId, ui.ButtonSet.OK_CANCEL);
    if (responseDriveFolderId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let driveFolderId = responseDriveFolderId.getResponseText();
    // Current Year (YouTube)
    let promptYtCurrentYear = 'Current Year (YouTube): Current year for YouTube data in "yyyy" format.';
    promptYtCurrentYear += (currentSettings.ytCurrentYear ? `\n\nCurrent Value: ${currentSettings.ytCurrentYear}` : '');
    let responseYtCurrentYear = ui.prompt(promptYtCurrentYear, ui.ButtonSet.OK_CANCEL);
    if (responseYtCurrentYear.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let ytCurrentYear = responseYtCurrentYear.getResponseText();
    // Current Year (Facebook)
    let promptFbCurrentYear = 'Current Year (Facebook): Current year for Facebook data in "yyyy" format.';
    promptFbCurrentYear += (currentSettings.fbCurrentYear ? `\n\nCurrent Value: ${currentSettings.fbCurrentYear}` : '');
    let responseFbCurrentYear = ui.prompt(promptFbCurrentYear, ui.ButtonSet.OK_CANCEL);
    if (responseFbCurrentYear.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fbCurrentYear = responseFbCurrentYear.getResponseText();
    // Current Spreadsheet ID
    let promptCurrentSpreadsheetId = 'Current Spreadsheet ID: Spreadsheet ID to record the analytics data in. This should be newly created (automatically) at the beginning of each year to avoid hitting the spreadsheet capacity limit.';
    promptCurrentSpreadsheetId += (currentSettings.currentSpreadsheetId ? `\n\nCurrent Value: ${currentSettings.currentSpreadsheetId}` : '');
    let responseCurrentSpreadsheetId = ui.prompt(promptCurrentSpreadsheetId, ui.ButtonSet.OK_CANCEL);
    if (responseCurrentSpreadsheetId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let currentSpreadsheetId = responseCurrentSpreadsheetId.getResponseText();
    // Set script properties
    let properties = {
      'ytClientId': ytClientId,
      'ytClientSecret': ytClientSecret,
      'fbAppId': fbAppId,
      'fbAppSecret': fbAppSecret,
      'driveFolderId': driveFolderId,
      'ytCurrentYear': ytCurrentYear,
      'fbCurrentYear': fbCurrentYear,
      'currentSpreadsheetId': currentSpreadsheetId,
      'setupComplete': true
    };
    PropertiesService.getScriptProperties().setProperties(properties, false);
    ui.alert('Complete: setup of script properties');
  } catch (error) {
    let message = errorMessage_(error);
    ui.alert(message);
  }
}

/* exported  */
/* global  */

////////////////////////
// Background Updates //
////////////////////////

/**
 * Periodically update analytics data using Google Apps Script's trigger.
 */
 function weeklyAnalyticsUpdate() {
  console.info('[weeklyAnalyticsUpdate] Initiating: A periodical task to update analytics data using Google Apps Script\'s trigger...'); // log
  var muteUiAlert = true;
  var muteMailNotification = false;
  var yearLimit = false;
  try {
    updateYouTubeAnalyticsData(muteUiAlert, muteMailNotification, yearLimit);
    console.info('[weeklyAnalyticsUpdate] Complete: Periodical task to update analytics data using Google Apps Script\'s trigger is complete.'); // log
  } catch (error) {
    console.error('[weeklyAnalyticsUpdate] Terminated.'); // log
  }
}

//////////////////////
// Common Functions //
//////////////////////

/**
 * Standarized error message
 * @param {Object} e Error object
 * @return {string} Standarized error message
 */
function errorMessage_(e) {
  return e.stack;
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
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale()); // See locale.js for the list of localized messages
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
    throw new Error(localizedMessages.messageList.general.error.spreadsheetUrl_.templateFileIdIsMissingInOptions);
  } else {
    return { 'url': null, 'created': false };
  }
}

/**
 * Enter log on designated spreadsheet.
 * @param {string} spreadsheetUrl URL of spreadsheet to enter log.
 * @param {string} sheetName Name of log sheet.
 * @param {string} logMessage Log message
 * @param {Date} timestamp [Optional] Date object for timestamp. If omitted, enters the time of execution.
 */
function enterLog_(spreadsheetUrl, sheetName, logMessage, timestamp = new Date()) {
  var targetSpreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
  var timestampString = formattedDate_(timestamp, targetSpreadsheet.getSpreadsheetTimeZone());
  targetSpreadsheet.getSheetByName(sheetName).appendRow([timestampString, logMessage]);
}
