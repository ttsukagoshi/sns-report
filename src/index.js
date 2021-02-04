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

/* exported checkYear_, errorMessage_, getConfig_, MESSAGE, onOpen, spreadsheetUrl_ */
/* global LocalizedMessage */

const LOG_SHEET_NAME = '99_Log';
const GITHUB_URL = 'https://github.com/ttsukagoshi/sns-report';
const MESSAGE = {
  'en_US': {
    'general': {
      'index': {
        'authorize': 'Authorize',
        'logoutReset': 'Logout/Reset',
        'updateChannelVideoList': 'Update Channel/Video List',
        'updateAll': 'Update All',
        'updateChannelList': 'Update Channel List',
        'updateVideoList': 'Update Video List',
        'analytics': 'Analytics',
        'getLatestData': 'Get Latest Data',
        'createAnalyticsSummary': 'Create Analytics Summary',
        'updatePageList': 'Update Page List',
        'settings': 'Settings',
        'setup': 'Setup',
        'checkSettings': 'Check Settings'
      },
      'error': {
        'errorTitle': 'Error',
        'errorMailSubject': '[SNS Report] Error Detected',
        'spreadsheetUrl_': {
          'templateFileIdIsMissingInOptions': 'The key "templateFileId" is missing in "options".'
        }
      },
      'misc': {
        'completedTitle': 'Completed'
      }
    },
    'youtube': {
      'authorizeYouTubeAPI': 'Authorize YouTube Data/Analytics API',
      'alreadyAuthorized': '[YouTube Data/Analytics API] You are already authorized.',
      'authorizationSuccessful': '[YouTube Data/Analytics API] Success! You can close this tab.',
      'authorizationDenied': '[YouTube Data/Analytics API] Denied. You can close this tab',
      'updatedChannelListLog': 'Success: updated channel list.',
      'updatedChannelListAlert': 'Updated summary channel list.',
      'updatedVideoListLog': 'Success: updated video list.',
      'updatedVideoListAlert': 'Updated summary video list.',
      'updateYouTubeAnalyticsDataMailTemplate': `This is an automatic mail that can be stopped or modified at:\n{{spreadsheetUrl}}\n\nFor more information on the Google Apps Script behind the spreadsheet, see ${GITHUB_URL}`,
      'newYouTubeSpreadsheetCreatedAlert': 'New YouTube spreadsheet created for {{year}}:\n{{url}}',
      'newYouTubeSpreadsheetCreatedMailSubject': '[SNS Report] New YouTube Spreadsheet Created for {{year}}',
      'newYouTubeSpreadsheetCreatedMailBody': 'New spreadsheet created to record YouTube analytics data for year {{year}} at:\n{{url}}\n\n',
      'updatedYouTubeChannelAnalyticsLog': 'Success: updated YouTube channel analytics for {{startDate}} to {{endDate}}.',
      'noUpdatesForYouTubeChannelAnalyticsLog': 'Success: no updates for YouTube channel analytics.',
      'updatedYouTubeChannelDemographicsLog': 'Success: updated YouTube channel demographics for {{latestMonth}} to {{thisYearMonth}}.',
      'updatedYouTubeVideoAnalyticsLog': 'Success: updated YouTube video analytics for {{startDate}} to {{endDate}}.',
      'noUpdatesForYouTubeVideoAnalyticsLog': 'Success: no updates for YouTube video analytics.',
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > YouTube > Authorize',
      'errorNoTextEnteredForChannelId': 'No text entered for channel ID.',
      'errorInvalidChannelId': 'Invalid Channel ID: "{{targetChannelId}}".\nMake sure to enter the ID of the YouTube channel that you own.',
      'errorNoTextEnteredForReportMonth': 'No text entered for report month.',
      'errorInvalidReportMonth': 'Invalid Report Month: "{{reportMonth}}".\nMake sure the report month is expressed in "yyyy-MM", e.g., enter "2020-03" for getting report for March 2020.',
      'reportCreated': 'Report for {{reportMonth}} of YouTube channel "{{targetChannelName}}" created.\nScript Time: {{scriptExeTime}} secs.'
    },
    'facebook': {
      'authorizeFacebookAPI': 'Authorize Facebook Graph API',
      'alreadyAuthorized': '[Facebook API] You are already authorized.',
      'authorizationSuccessful': '[Facebook API] Success! You can close this tab.',
      'authorizationDenied': '[Facebook API] Denied. You can close this tab',
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > Facebook > Authorize',
      'updatedPageListLog': 'Success: updated page list.',
      'updatedPageListAlert': 'Updated summary page list.',
      'updatedPagePostListLog': 'Success: updated page post list.',
      'updatedPagePostListAlert': 'Updated page post list.'
    }
  },
  'ja_JP': {
    'general': {
      'index': {
        'authorize': '認証',
        'logoutReset': 'ログアウト/リセット',
        'updateChannelVideoList': 'チャンネル/ビデオ一覧を更新',
        'updateAll': 'すべてを更新',
        'updateChannelList': 'チャンネル一覧を更新',
        'updateVideoList': 'ビデオ一覧を更新',
        'analytics': 'アナリティクス',
        'getLatestData': '最新データを取得',
        'createAnalyticsSummary': 'アナリティクスのサマリー作成',
        'updatePageList': 'ページ一覧を更新',
        'settings': '設定',
        'setup': '初期設定',
        'checkSettings': '設定確認'
      },
      'error': {
        'errorTitle': 'エラー',
        'errorMailSubject': '[SNS Report] Error Detected',
        'spreadsheetUrl_': {
          'templateFileIdIsMissingInOptions': '変数"option"内のキー"templateFileId"が指定されていません。'
        }
      },
      'misc': {
        'completedTitle': '完了'
      }
    },
    'youtube': {
      'authorizeYouTubeAPI': 'YouTube Data/Analytics APIを認証',
      'alreadyAuthorized': '[YouTube Data/Analytics API] すでに認証済みです。',
      'authorizationSuccessful': '[YouTube Data/Analytics API] 認証成功。このタブを閉じても大丈夫です。',
      'authorizationDenied': '[YouTube Data/Analytics API] 認証に失敗しました。このタブは閉じてください。',
      'updatedChannelListLog': 'Success: updated channel list.', // Log message will not be translated.
      'updatedChannelListAlert': 'チャンネル一覧を更新完了。',
      'updatedVideoListLog': 'Success: updated video list.', // Log message will not be translated.
      'updatedVideoListAlert': 'ビデオ一覧を更新完了。',
      'updateYouTubeAnalyticsDataMailTemplate': `This is an automatic mail that can be stopped or modified at:\n{{spreadsheetUrl}}\n\nFor more information on the Google Apps Script behind the spreadsheet, see ${GITHUB_URL}`,
      'newYouTubeSpreadsheetCreatedAlert': 'New YouTube spreadsheet created for {{year}}:\n{{url}}',
      'newYouTubeSpreadsheetCreatedMailSubject': '[SNS Report] New YouTube Spreadsheet Created for {{year}}',
      'newYouTubeSpreadsheetCreatedMailBody': 'New spreadsheet created to record YouTube analytics data for year {{year}} at:\n{{url}}\n\n',
      'updatedYouTubeChannelAnalyticsLog': 'Success: updated YouTube channel analytics for {{startDate}} to {{endDate}}.', // Log message will not be translated.
      'noUpdatesForYouTubeChannelAnalyticsLog': 'Success: no updates for YouTube channel analytics.', // Log message will not be translated.
      'updatedYouTubeChannelDemographicsLog': 'Success: updated YouTube channel demographics for {{latestMonth}} to {{thisYearMonth}}.', // Log message will not be translated.
      'updatedYouTubeVideoAnalyticsLog': 'Success: updated YouTube video analytics for {{startDate}} to {{endDate}}.', // Log message will not be translated.
      'noUpdatesForYouTubeVideoAnalyticsLog': 'Success: no updates for YouTube video analytics.', // Log message will not be translated.
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > YouTube > Authorize',
      'errorNoTextEnteredForChannelId': 'No text entered for channel ID.',
      'errorInvalidChannelId': 'Invalid Channel ID: "{{targetChannelId}}".\nMake sure to enter the ID of the YouTube channel that you own.',
      'errorNoTextEnteredForReportMonth': 'No text entered for report month.',
      'errorInvalidReportMonth': 'Invalid Report Month: "{{reportMonth}}".\nMake sure the report month is expressed in "yyyy-MM", e.g., enter "2020-03" for getting report for March 2020.',
      'reportCreated': 'Report for {{reportMonth}} of YouTube channel "{{targetChannelName}}" created.\nScript Time: {{scriptExeTime}} secs.'
    },
    'facebook': {
      'authorizeFacebookAPI': 'Facebook Graph APIを認証',
      'alreadyAuthorized': '[Facebook API] すでに認証済みです。',
      'authorizationSuccessful': '[Facebook API] 認証成功。このタブを閉じても大丈夫です。',
      'authorizationDenied': '[Facebook API] 認証に失敗しました。このタブは閉じてください。',
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > Facebook > Authorize',
      'updatedPageListLog': 'Success: updated page list.', // Log message will not be translated.
      'updatedPageListAlert': 'Updated summary page list.',
      'updatedPagePostListLog': 'Success: updated page post list.', // Log message will not be translated.
      'updatedPagePostListAlert': 'Updated page post list.'
    }
  }
};

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
  /*
  let message = `Error: line - ${e.lineNumber}\n${e.stack}`;
  return message;
  */
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
