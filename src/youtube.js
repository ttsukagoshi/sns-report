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

/* exported authCallbackYouTubeAPI_, createYouTubeAnalyticsSummary, logoutYouTube, showSidebarYouTubeApi, updateAllYouTubeList, updateYouTubeAnalyticsData */
/* global enterLog_, errorMessage_, formattedDate_, getConfig_, getSpreadsheetList_, LocalizedMessage, LOG_SHEET_NAME, OAuth2, spreadsheetUrl_ */

// API versions
const YOUTUBE_DATA_API_VERSION = 'v3';
const YOUTUBE_ANALYTICS_API_VERSION = 'v2';
// Spreadsheet ID to use as template for creating a new spreadsheet to enter YouTube analytics data
const YOUTUBE_NEW_SPREADSHEET_ID = '16duRDHJ8d6k6xy0C2_m2IkHxIJc-OPLp7SuQD02gzig';
// Other global variables and common functions are defined on index.js

///////////////////
// Authorize 認証//
///////////////////

/**
 * Direct the user to the authorization URL
 * https://github.com/gsuitedevs/apps-script-oauth2#2-direct-the-user-to-the-authorization-url
 */
function showSidebarYouTubeApi() {
  var youtubeAPIService = getYouTubeAPIService_();
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  if (!youtubeAPIService.hasAccess()) {
    let authorizationUrl = youtubeAPIService.getAuthorizationUrl();
    let template = HtmlService.createTemplate(`<a href="<?= authorizationUrl ?>" target="_blank">${localizedMessages.messageList.youtube.authorizeYouTubeAPI}</a>.`);
    template.authorizationUrl = authorizationUrl;
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    let template = HtmlService.createTemplate(
      localizedMessages.messageList.youtube.alreadyAuthorized);
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  }
}
/**
 * Logout
 * https://github.com/gsuitedevs/apps-script-oauth2#logout
 */
function logoutYouTube() {
  var service = getYouTubeAPIService_();
  service.reset();
}

/**
 * Create the OAuth2 service
 * https://github.com/gsuitedevs/apps-script-oauth2#1-create-the-oauth2-service
 */
function getYouTubeAPIService_() {
  // Script Properties
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var [ytClientId, ytClientSecret] = [scriptProperties.ytClientId, scriptProperties.ytClientSecret];

  return OAuth2.createService('youtubeAPI')
    // Set the endpoint URLs, which are the same for all Google services.
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')

    // Set the client ID and secret, from the Google Developers Console.
    .setClientId(ytClientId)
    .setClientSecret(ytClientSecret)

    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction('authCallbackYouTubeAPI_')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())

    // Set the scopes to request (space-separated for Google services).
    .setScope('https://www.googleapis.com/auth/youtube https://www.googleapis.com/auth/youtubepartner https://www.googleapis.com/auth/yt-analytics-monetary.readonly https://www.googleapis.com/auth/yt-analytics.readonly')

    // Below are Google-specific OAuth2 parameters.
    /*
    // Sets the login hint, which will prevent the account chooser screen
    // from being shown to users logged in with multiple accounts.
    .setParam('login_hint', Session.getEffectiveUser().getEmail())
    */
    // Requests offline access.
    .setParam('access_type', 'offline')

    // Consent prompt is required to ensure a refresh token is always
    // returned when requesting offline access.
    .setParam('prompt', 'consent');
}

/**
 * Handle the callback
 * https://github.com/gsuitedevs/apps-script-oauth2#3-handle-the-callback
 * @param {*} request 
 */
function authCallbackYouTubeAPI_(request) {
  var youtubeAPIService = getYouTubeAPIService_();
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var isAuthorized = youtubeAPIService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput(localizedMessages.messageList.youtube.authorizationSuccessful);
  } else {
    return HtmlService.createHtmlOutput(localizedMessages.messageList.youtube.authorizationDenied);
  }
}

///////////////////////////////
// Update Channel/Video List //
///////////////////////////////

/**
 * Update both channel and video summary list.
 */
function updateAllYouTubeList() {
  console.info('[updateAllYouTubeList] Initiating: Updating channel and video summary list...'); // log
  updateYouTubeSummaryChannelList();
  updateYouTubeSummaryVideoList();
  console.info('[updateAllYouTubeList] Complete: Updating of channel and video summary list is complete.'); // log
}

/**
 * List the authorized user's channel(s) on the summary spreadsheet
 * @param {boolean} muteUiAlert [Optional] Mute ui.alert() when true; defaults to false.
 */
function updateYouTubeSummaryChannelList(muteUiAlert = false) {
  console.info('[updateYouTubeSummaryChannelList] Initiating: Listing the authorized user\'s channel(s) on the summary spreadsheet...'); // log
  console.info(`[updateYouTubeSummaryChannelList] UI alert is ${muteUiAlert ? 'muted' : 'not muted'}.`); // log
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  var config = getConfig_();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
  var targetSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == scriptProperties.ytCurrentYear && value.PLATFORM == 'YouTube'))[0].URL;
  var now = new Date();
  try {
    // Get list of channel(s)
    let channelListFull = youtubeMyChannelList_();
    // Extract data for copying into spreadsheet
    console.info('[updateYouTubeSummaryChannelList] Extracting channel list data for copying into spreadsheet...'); // log
    let channelList = channelListFull.map((element, index) => {
      let num = index + 1;
      let thumbnailUrl = element.snippet.thumbnails.default.url;
      let thumbnailUrlFunction = `=image("${thumbnailUrl}")`; // For using the image function on spreadsheet
      let id = element.id;
      let title = element.snippet.title;
      let description = element.snippet.description;
      let publishedAt = element.snippet.publishedAt;
      let publishedAtLocal = formattedDate_(new Date(publishedAt), timeZone);
      let viewCount = element.statistics.viewCount;
      let subscriberCount = element.statistics.subscriberCount;
      let videoCount = element.statistics.videoCount;
      let timestamp = formattedDate_(now, timeZone);
      return [timestamp, num, thumbnailUrlFunction, thumbnailUrl, id, title, description, publishedAt, publishedAtLocal, viewCount, subscriberCount, videoCount];
    }
    );
    // Set the text values into spreadsheets (summary and individual)
    //// Renew channel list of summary spreadsheet
    let myChannelsSheet = ss.getSheetByName(config.SHEET_NAME_MY_CHANNELS);
    if (myChannelsSheet.getLastRow() > 1) {
      myChannelsSheet.getRange(2, 1, myChannelsSheet.getLastRow() - 1, myChannelsSheet.getLastColumn())
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
    myChannelsSheet.getRange(2, 1, channelList.length, channelList[0].length) // Assuming that table body to which the list is copied starts from the 2nd row of column 1 ('A' column).
      .setValues(channelList);
    //// Add row(s) to current spreadsheet of this year
    let currentSheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl).getSheetByName(config.SHEET_NAME_MY_CHANNELS);
    currentSheet.getRange(currentSheet.getLastRow() + 1, 1, channelList.length, channelList[0].length) // Assuming that table body to which the list is copied starts from column 1 ('A' column).
      .setValues(channelList);
    // Log & Notify
    enterLog_(targetSpreadsheetUrl, LOG_SHEET_NAME, localizedMessages.messageList.youtube.updatedChannelListLog, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.misc.completedTitle, localizedMessages.messageList.youtube.updatedChannelListAlert, ui.ButtonSet.OK);
    }
    console.info('[updateYouTubeSummaryChannelList] Complete: updated channel list.'); // log
    return channelList;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheetUrl, LOG_SHEET_NAME, message, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.error.errorTitle, message, ui.ButtonSet.OK);
    }
    console.error('[updateYouTubeSummaryChannelList] ' + message); // log
    return null;
  }
}

/**
 * List the authorized user's video(s) on the summary spreadsheet
 * @param {boolean} muteUiAlert [Optional] Mute ui.alert() when true; defaults to false.
 * @returns {Array} The 2d JavaScript array of the updated video list.
 * This function returns null if no video was returned or if an error was thrown.
 */
function updateYouTubeSummaryVideoList(muteUiAlert = false) {
  console.info('[updateYouTubeSummaryVideoList] Initiating: Listing the authorized user\'s video(s) on the summary spreadsheet...'); // log
  console.info(`[updateYouTubeSummaryVideoList] UI alert is ${muteUiAlert ? 'muted' : 'not muted'}.`); // log
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  var config = getConfig_();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
  var targetSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == scriptProperties.ytCurrentYear && value.PLATFORM == 'YouTube'))[0].URL;
  var now = new Date();
  try {
    // Get list of video(s)
    let videoListFull = youtubeMyVideoList_();
    // Extract data for copying into spreadsheet
    console.info('[updateYouTubeSummaryVideoList] Extracting video list data for copying into spreadsheet...'); // log
    let videoList = videoListFull.map((element, index) => {
      let num = index + 1;
      let thumbnailUrl = element.snippet.thumbnails.high.url;
      let thumbnailUrlFunction = `=image("${thumbnailUrl}")`; // For using the image function on spreadsheet
      let channelId = element.snippet.channelId;
      let channelTitle = element.snippet.channelTitle;
      let videoId = element.id.videoId;
      let title = element.snippet.title;
      let description = element.snippet.description;
      let videoUrl = `https://www.youtube.com/watch?v=${videoId}`;
      let publishedAt = element.snippet.publishedAt;
      let publishedAtLocal = formattedDate_(new Date(publishedAt), timeZone);
      let timestamp = formattedDate_(now, timeZone);
      // Statistics and status of each video
      let videoStats = youtubeVideos_([videoId], false)[0];
      let duration = videoStats.contentDetails.duration;
      let caption = videoStats.contentDetails.caption;
      let privacyStatus = videoStats.status.privacyStatus;
      let viewCount = videoStats.statistics.viewCount;
      let likeCount = videoStats.statistics.likeCount;
      let dislikeCount = videoStats.statistics.dislikeCount;
      // List elements to copy to spreadsheet
      let videoListElement = [
        timestamp,
        num,
        thumbnailUrlFunction,
        thumbnailUrl,
        channelId,
        channelTitle,
        videoId,
        title,
        description,
        videoUrl,
        duration,
        caption,
        publishedAt,
        publishedAtLocal,
        privacyStatus,
        viewCount,
        likeCount,
        dislikeCount
      ];
      return videoListElement;
    });
    // Set the text values into spreadsheets (summary and individual)
    //// Renew channel list of summary spreadsheet
    let myVideosSheet = ss.getSheetByName(config.SHEET_NAME_MY_VIDEOS);
    if (myVideosSheet.getLastRow() > 1) {
      myVideosSheet.getRange(2, 1, myVideosSheet.getLastRow() - 1, myVideosSheet.getLastColumn())
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
    myVideosSheet.getRange(2, 1, videoList.length, videoList[0].length) // Assuming that table body to which the list is copied starts from the 4th row of column 1 ('A' column).
      .setValues(videoList);
    //// Add row(s) to current spreadsheet of this year
    let currentSheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl).getSheetByName(config.SHEET_NAME_MY_VIDEOS);
    currentSheet.getRange(currentSheet.getLastRow() + 1, 1, videoList.length, videoList[0].length) // Assuming that table body to which the list is copied starts from the 4th row of column 1 ('A' column).
      .setValues(videoList);
    // Log & Notify
    enterLog_(targetSpreadsheetUrl, LOG_SHEET_NAME, localizedMessages.messageList.youtube.updatedVideoListLog, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.misc.completedTitle, localizedMessages.messageList.youtube.updatedVideoListAlert, ui.ButtonSet.OK);
    }
    console.info('[updateYouTubeSummaryVideoList] Complete: updated video list.'); // log
    return videoList;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheetUrl, LOG_SHEET_NAME, message, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.error.errorTitle, message, ui.ButtonSet.OK);
    }
    console.error('[updateYouTubeSummaryVideoList] ' + message); // log
    return null;
  }
}

/**
 * Get the list of all channels that the authorized user owns.
 * Channel properties are described at https://developers.google.com/youtube/v3/docs/channels#resource
 * @returns {array} Array of Javascript objects of channel properties.
 */
function youtubeMyChannelList_() {
  console.info('[youtubeMyChannelList_] Initiating: Getting the list of all channels that the authorized user owns...'); // log
  // Get the list of channels
  var channelParameters = {
    part: 'snippet,statistics',
    mine: true
  };
  var channelList = JSON.parse(youtubeData_('channels', channelParameters));
  console.info(`[youtubeMyChannelList_] Initial channel list retrieved.`); // log
  if (channelList.nextPageToken) {
    console.info(`[youtubeMyChannelList_] Next Page Token exists. Retrieving next page...`); // log
    let nextChannelList = {};
    nextChannelList['nextPageToken'] = channelList.nextPageToken;
    while (nextChannelList.nextPageToken) {
      let nextChannelParameters = {};
      nextChannelParameters = {
        part: 'snippet, statistics',
        mine: true,
        pageToken: nextChannelList.nextPageToken
      };
      console.info(`[youtubeMyChannelList_] Next Page Token: "${nextChannelParameters.pageToken}"`); // log
      nextChannelList = {};
      nextChannelList = JSON.parse(youtubeData_('channels', nextChannelParameters));
      nextChannelList.items.forEach(value => channelList.items.push(value));
      console.info(`[youtubeMyChannelList_] Next channel list retrieved.`); // log
    }
  }
  console.info(`[youtubeMyChannelList_] Completed: returning data for ${channelList.items.length} channel(s).`); // log
  return channelList.items;
}

/**
 * Get the list of all video(s) that the authorized user owns.
 * Video properties are described at https://developers.google.com/youtube/v3/docs/videos#resource
 * @returns {array} Array of Javascript objects of video properties.
 */
function youtubeMyVideoList_() {
  console.info('[youtubeMyVideoList_] Initiating: Getting the list of all video(s) that the authorized user owns...'); // log
  var videoParameters = {
    part: 'snippet',
    forMine: true,
    type: 'video'
  };
  var videoListString = youtubeData_('search', videoParameters);
  var videoList = JSON.parse(videoListString);
  console.info(`[youtubeMyVideoList_] Initial video list retrieved.`); // log
  if (videoList.nextPageToken) {
    console.info(`[youtubeMyVideoList_] Next Page Token exists. Retrieving next page...`); // log
    let nextVideoListString = '';
    let nextVideoList = {};
    nextVideoList['nextPageToken'] = videoList.nextPageToken;
    while (nextVideoList.nextPageToken) {
      let nextParameters = {};
      nextParameters = {
        part: 'snippet',
        forMine: true,
        type: 'video',
        pageToken: nextVideoList.nextPageToken
      };
      console.info(`[youtubeMyVideoList_] Next Page Token: "${nextParameters.pageToken}"`); // log
      nextVideoListString = youtubeData_('search', nextParameters);
      nextVideoList = JSON.parse(nextVideoListString);
      nextVideoList.items.forEach(value => videoList.items.push(value));
      console.info(`[youtubeMyVideoList_] Next video list retrieved.`); // log
    }
  }
  console.info(`[youtubeMyVideoList_] Completed: returning data for ${videoList.items.length} video(s).`); // log
  return videoList.items;
}

/**
 * Get the details of specific YouTube video(s).
 * Video properties are described at https://developers.google.com/youtube/v3/docs/videos#resource
 * @param {array} videoIds Array of YouTube video IDs to retrieve the details. 
 * Subject to API quota; bulk request may result in HTTP error 400. See https://developers.google.com/youtube/v3/docs/videos/list#parameters
 * @param {boolean} getDetails [Optional] Gets the snippet of the YouTube video(s) as well as its status and statistics. Defaults to false.
 * @returns {array} Array of Javascript objects of video properties.
 */
function youtubeVideos_(videoIds, getDetails = false) {
  console.info(`[youtubeVideos_] Initiating: Getting the details of ${videoIds.length} video(s)...`); // log
  var videoParameters = {
    part: (getDetails == true ? 'snippet,statistics,status,contentDetails' : 'statistics,status,contentDetails'),
    id: videoIds.join()
  };
  var videoDetails = JSON.parse(youtubeData_('videos', videoParameters));
  console.info(`[youtubeVideos_] Initial video data retrieved.`); // log
  if (videoDetails.nextPageToken) {
    console.info(`[youtubeVideos_] Next Page Token exists. Retrieving next page...`); // log
    let nextVideoDetails = {};
    nextVideoDetails['nextPageToken'] = videoDetails.nextPageToken;
    while (nextVideoDetails.nextPageToken) {
      let nextParameters = {};
      nextParameters = {
        part: (getDetails == true ? 'snippet,statistics,status,contentDetails' : 'statistics,status,contentDetails'),
        id: videoIds.join(),
        pageToken: nextVideoDetails.nextPageToken
      };
      console.info(`[youtubeVideos_] Next Page Token: "${nextParameters.pageToken}"`); // log
      nextVideoDetails = {};
      nextVideoDetails = JSON.parse(youtubeData_('videos', nextParameters));
      nextVideoDetails.items.forEach(value => videoDetails.items.push(value));
      console.info(`[youtubeVideos_] Next video data retrieved.`); // log
    }
  }
  console.info(`[youtubeVideos_] Completed: returning data for ${videoDetails.items.length} video(s).`); // log
  return videoDetails.items;
}

/**
 * GET request using YouTube Data API.
 * https://developers.google.com/youtube/v3/docs
 * @param {string} resourceType Resource type to target the GET request. https://developers.google.com/youtube/v3/docs#resource-types
 * @param {Object} parameters Parameters for the request in form of a Javascript object.
 */
function youtubeData_(resourceType, parameters) {
  var youtubeAPIService = getYouTubeAPIService_();
  if (!youtubeAPIService.hasAccess()) {
    throw new Error('Unauthorized. Get authorized by Menu > YouTube > Authorize');
  }
  var baseUrl = `https://www.googleapis.com/youtube/${YOUTUBE_DATA_API_VERSION}/${resourceType}`;
  var paramString = '?';
  for (let k in parameters) {
    let param = `${k}=${encodeURIComponent(parameters[k])}`;
    if (paramString.slice(-1) !== '?') {
      param = '&' + param;
    }
    paramString += param;
  }
  var url = baseUrl + paramString;
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + encodeURIComponent(youtubeAPIService.getAccessToken())
    }
  });
  return response;
}

///////////////
// Analytics //
///////////////

/**
 * Get latest analytics data for YouTube channel and videos that the authorized user owns.
 * @param {boolean} muteUiAlert [Optional] Mute SpreadsheetApp.getUi().alert() when true; defaults to false.
 * @param {boolean} muteMailNotification [Optional] Mute email notification when true; defaults to true.
 * @param {boolean} yearLimit [Optional] When true, limit the latest data to obtain to the last month of the targetYear, i.e., December. Defaults to true.
 */
function updateYouTubeAnalyticsData(muteUiAlert = false, muteMailNotification = true, yearLimit = true) {
  console.info('[updateYouTubeAnalyticsData] Initiating: Getting the latest analytics data for YouTube channel and videos that the authorized user owns...'); // log
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  var myEmail = Session.getActiveUser().getEmail();
  var mailTemplate = localizedMessages.replaceUpdateYouTubeAnalyticsDataMailTemplate(ss.getUrl());
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  // var yearLimit = true;
  console.info(`[updateYouTubeAnalyticsData] UI alert is ${muteUiAlert ? 'muted' : 'not muted'}.\nMail notification is ${muteMailNotification ? 'disabled' : 'enabled'}.`); // log
  try {
    let ytCurrentYear = parseInt(scriptProperties.ytCurrentYear);
    // Get latest data
    console.info('[updateYouTubeAnalyticsData] Retrieving latest YouTube analytics data...'); // log
    let updatedChannelAnalyticsDate = youtubeAnalyticsChannel(ytCurrentYear, yearLimit);
    let updateChannelDemographics = youtubeAnalyticsDemographics(ytCurrentYear, yearLimit);
    let updatedVideoAnalyticsDate = youtubeAnalyticsVideo(ytCurrentYear, yearLimit);
    // Determine change-of-the-year
    let changeOfYear = (
      updatedChannelAnalyticsDate.latestDateReturned.getFullYear() > ytCurrentYear
      && updateChannelDemographics >= `${ytCurrentYear}-12`
      && updatedVideoAnalyticsDate.latestDateReturned.getFullYear() > ytCurrentYear
    );
    if (changeOfYear) {
      console.warn('[updateYouTubeAnalyticsData] Change of year detected. Checking to see if new files need to be created...'); // log
      let config = getConfig_();
      let options = {
        createNewFile: true,
        driveFolderId: scriptProperties.driveFolderId,
        templateFileId: YOUTUBE_NEW_SPREADSHEET_ID,
        newFileName: 'YouTube Analytics',
        newFileNamePrefix: config.SPREADSHEET_NAME_PREFIX
      };
      // Create an array of new year(s)
      let year = ytCurrentYear + 1;
      let now = new Date();
      while (year <= now.getFullYear()) {
        options['newFileNameSuffix'] = ` ${year}`;
        // Check for existing spreadsheet(s) and create a new one if there are no matches.
        let newFileUrl = spreadsheetUrl_(config.SHEET_NAME_SPREADSHEET_LIST, year, 'YouTube', options); // spreadsheetUrl_() is defined on index.js
        youtubeAnalyticsChannel(year, yearLimit);
        youtubeAnalyticsDemographics(year, yearLimit);
        youtubeAnalyticsVideo(year, yearLimit);
        // Update script property
        PropertiesService.getScriptProperties().setProperty('ytCurrentYear', year);
        if (newFileUrl.created) {
          if (!muteUiAlert) {
            SpreadsheetApp.getUi().alert(localizedMessages.replaceNewYouTubeSpreadsheetCreatedAlert(year, newFileUrl.url));
          }
          if (!muteMailNotification) {
            let subject = localizedMessages.replaceNewYouTubeSpreadsheetCreatedMailSubject(year);
            let body = localizedMessages.replaceNewYouTubeSpreadsheetCreatedMailBody(year, newFileUrl.url) + mailTemplate;
            MailApp.sendEmail(myEmail, subject, body);
          }
        }
        // Update year for loop process
        year += 1;
      }
    } else {
      console.info('[updateYouTubeAnalyticsData] No change of year detected.'); // log
    }
  } catch (error) {
    let message = errorMessage_(error);
    console.error('[updateYouTubeAnalyticsData] ' + message); // log
    if (!muteUiAlert) {
      SpreadsheetApp.getUi().alert(message);
    }
    if (!muteMailNotification) {
      let subject = localizedMessages.messageList.general.error.errorMailSubject;
      let body = `${message}\n\n` + mailTemplate;
      MailApp.sendEmail(myEmail, subject, body);
    }
    throw error;
  }
}

/**
 * Update day-by-day YouTube Channel Analytics for the target year
 * If no previous data is available, this function will retrieve channel analytics starting from Jan 1 of the target year.
 * @param {number} targetYear Target year in yyyy
 * @param {boolean} yearLimit When true, limit the latest data to obtain to the end of the targetYear, i.e., Dec 31. Defaults to true.
 * @returns {Object} JavaScript object with two keys: latestDateOnSpreadsheet and latestDateReturned.
 * The value of the former is the latest date object on the updated analytics spreadsheet,
 * while the latter is paired with a date object value expressing the latest date returned in the original youtubeAnalyticsReportsQuery_().
 */
function youtubeAnalyticsChannel(targetYear, yearLimit = true) {
  // Get target spreadsheet
  var config = getConfig_();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
  var targetSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == targetYear && value.PLATFORM == 'YouTube'))[0].URL;
  var targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  var targetSheet = targetSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_ANALYTICS);
  var now = new Date();
  var localizedMessages = new LocalizedMessage(targetSpreadsheet.getSpreadsheetLocale());
  try {
    // Check the date of the latest analytics data and define startDate for youtubeAnalyticsReportsQuery_()
    // If the value returned for getLatestDate_() is null, i.e., there are no previous dates recorded in targetSheet,
    // latestDate will be Dec 31 of the previous year of targetYear
    let currentLatestOnSpreadsheet = getLatestDate_(targetSheet, 1); // Assuming that the date is recorded on column A of the targetSheet.
    let latestDate = (currentLatestOnSpreadsheet ? new Date(currentLatestOnSpreadsheet.getTime()) : new Date(targetYear - 1, 11, 31));
    let startDateObj = new Date(latestDate.setDate(latestDate.getDate() + 1));
    // Setting parameters for youtubeAnalyticsReportsQuery_()
    let startDate = formattedDateAnalytics_(startDateObj);
    let endDateObj = new Date(now.getTime());
    let endDate = formattedDateAnalytics_(endDateObj);
    let metrics = 'views,likes,dislikes,subscribersGained,subscribersLost,estimatedMinutesWatched,averageViewDuration,cardImpressions,cardClicks';
    let ids = 'channel==MINE';
    let options = {
      dimensions: 'day,channel'
    };
    // Get analytics data
    let reports = JSON.parse(youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids, options));
    // Get day-by-day count of published videos per channel
    let myVideosCountByPubDate = youtubeMyVideoCountByPubDate_();
    // Process data for later analysis
    let data = reports.rows.reduce((acc, row) => {
      // Assuming that the DAY and CHANNEL ID fields come in the 1st and 2nd column of 'row', respectively
      if ((yearLimit && row[0] <= `${targetYear}-12-31`) || !yearLimit) {
        let publishedVideoCount = (myVideosCountByPubDate[row[1]][row[0]] ? myVideosCountByPubDate[row[1]][row[0]] : 0);
        let yearMonth = row[0].slice(0, 7);
        let concatRow = row.concat([publishedVideoCount, yearMonth]);
        acc.channelData.push(concatRow);
      }
      if (row[0] > acc.latest) {
        acc.latest = row[0];
      }
      return acc;
    }, { channelData: [], latest: `${targetYear}-01-01` });
    let updatedLatestDateObj = {
      latestDateOnSpreadsheet: currentLatestOnSpreadsheet,
      latestDateReturned: currentLatestOnSpreadsheet
    };
    if (data.channelData && data.channelData.length > 0) {
      // Copy on spreadsheet
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, data.channelData.length, data.channelData[0].length).setValues(data.channelData);
      // Get latest updated date
      updatedLatestDateObj.latestDateOnSpreadsheet = getLatestDate_(targetSheet, 1);
      updatedLatestDateObj.latestDateReturned = yMd2Date_(data.latest);
      // Log
      enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, localizedMessages.replaceUpdatedYouTubeChannelAnalyticsLog(startDate, formattedDateAnalytics_(updatedLatestDateObj.latestDateOnSpreadsheet)), now);
    } else {
      enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, localizedMessages.messageList.youtube.noUpdatesForYouTubeChannelAnalyticsLog, now);
    }
    return updatedLatestDateObj;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, message, now);
    throw error;
  }
}

/**
 * Get a day-by-day count of published videos per channel in form of a JavaScript object.
 * See detailed definition of publishedAt at https://developers.google.com/youtube/v3/docs/videos#properties
 * @returns {Object} {<Channel ID>:{<yyyy-MM-dd>: <count>}}
 */
function youtubeMyVideoCountByPubDate_() {
  console.info('[youtubeMyVideoCountByPubDate_] Initiating: Getting a day-by-day count of published videos per channel in form of a JavaScript object.'); // log
  var myVideos = youtubeMyVideoList_();
  var countList = myVideos.reduce(function (acc, cur) {
    let channelId = cur.snippet.channelId;
    let publishedAtPT = formattedDateAnalytics_(new Date(cur.snippet.publishedAt), 'PST'); // yyyy-MM-dd in Pacific Time, i.e., PST or PDT depending on the date
    if (!acc[channelId]) {
      acc[channelId] = {};
    }
    if (!acc[channelId][publishedAtPT]) {
      acc[channelId][publishedAtPT] = 0;
    }
    acc[channelId][publishedAtPT] += 1;
    return acc;
  }, {});
  console.info(`[youtubeMyVideoCountByPubDate_] Retrieved video count: ${JSON.stringify(countList)}`); // log
  return countList;
}

/**
 * Update monthly YouTube Channel demographics for the target year
 * If no previous data is available, this function will retrieve channel analytics starting from Jan of the target year.
 * @param {number} targetYear Target year in yyyy
 * @param {boolean} yearLimit When true, limit the latest data to obtain to the last month of the targetYear, i.e., December. Defaults to true.
 * @returns {string} The last year-month (yyyy-MM) entered in the spreadsheet.
 */
function youtubeAnalyticsDemographics(targetYear, yearLimit = true) {
  // Get target spreadsheet
  var config = getConfig_();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
  var targetSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == targetYear && value.PLATFORM == 'YouTube'))[0].URL;
  var targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  var targetSheet = targetSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_DEMOGRAPHICS);
  var now = new Date();
  var localizedMessages = new LocalizedMessage(targetSpreadsheet.getSpreadsheetLocale());
  try {
    // Check the month of the latest analytics data and define startDate for youtubeAnalyticsReportsQuery_()
    // If the value returned for getLatestMonth_() is null, i.e., there are no previous month recorded in targetSheet,
    // latestMonth will be January of the targetYear ('yyyy-01').
    let latestMonthPre = getLatestMonth_(targetSheet, 1);
    let latestMonth = (latestMonthPre ? latestMonthPre : targetYear + '-01'); // Assuming that the year-month (yyyy-MM) is recorded on column A of the targetSheet.
    let latestMonthDate = yearMonth2Date_(latestMonth); // latestMonth in Date object (the first day of that yyyy-MM)
    // Get existing data in form of a 2d-array to overwrite the latest data
    let existingData = targetSheet.getDataRange().getValues();
    existingData.shift(); // Assuming that the first row is a header row.
    let existingDataUpdate = existingData.filter(element => element[0] !== latestMonth);
    // Determine the final endDate as currentLatestMonthDate. If yearLimit is true, this will be no later than December 31st of the targetYear.
    let currentLatestMonthDate = (
      yearLimit == true && parseInt(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy')) > parseInt(targetYear)
        ? new Date(targetYear, 11, 31)
        : now
    );
    // Set common parameters for youtubeAnalyticsReportsQuery_()
    let startDate = latestMonth + '-01'; // This must be yyyy-MM-01, i.e., the first day of the target month(s) for this query
    let endDate = '';
    let metrics = 'viewerPercentage';
    let ids = 'channel==MINE';
    let options = {
      dimensions: 'channel,ageGroup,gender'
    };
    let thisYearMonth = formattedDateAnalytics_(currentLatestMonthDate).slice(0, 7);
    let reports = {};
    let rowsData = [];
    while (latestMonthDate.getTime() <= currentLatestMonthDate.getTime()) {
      // Define endDate
      let endDatePre = new Date(latestMonthDate.setMonth(latestMonthDate.getMonth() + 1));
      endDate = formattedDateAnalytics_(new Date(endDatePre.setDate(endDatePre.getDate() - 1)));
      // Define thisYearMonth
      thisYearMonth = startDate.slice(0, 7);
      // API query
      reports = JSON.parse(youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids, options));
      // Add yyyy-MM to the result
      rowsData = reports.rows.slice();
      for (let i = 0; i < rowsData.length; i++) {
        let element = rowsData[i];
        element.unshift(thisYearMonth);
        existingDataUpdate.push(element);
      }
      // Set variable(s) for next loop
      startDate = formattedDateAnalytics_(latestMonthDate);
    }
    // Copy on spreadsheet
    targetSheet.getRange(2, 1, existingDataUpdate.length, existingDataUpdate[0].length).setValues(existingDataUpdate); // Assuming that the 1st row of the targetSheet is header row and that the actual data starts from the 2nd row
    // Log
    enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, localizedMessages.replaceUpdatedYouTubeChannelDemographicsLog(latestMonth, thisYearMonth), now);
    return thisYearMonth;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, message, now);
    throw error;
  }
}

/**
 * Update day-by-day YouTube video analytics for the target year
 * If no previous data is available, this function will retrieve vidoe analytics starting from Jan 1 of the target year.
 * @param {number} targetYear Target year in yyyy
 * @param {boolean} yearLimit When true, limit the latest data to obtain to the end of the targetYear, i.e., Dec 31. Defaults to true.
 * @returns {Date} Latest date object of the updated analytics data.
 */
function youtubeAnalyticsVideo(targetYear, yearLimit = true) {
  // Get target spreadsheet
  var config = getConfig_();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
  var targetSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == targetYear && value.PLATFORM == 'YouTube'))[0].URL;
  var targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  var targetSheet = targetSpreadsheet.getSheetByName(config.SHEET_NAME_VIDEO_ANALYTICS);
  var now = new Date();
  var localizedMessages = new LocalizedMessage(targetSpreadsheet.getSpreadsheetLocale());
  try {
    // Check the date of the latest analytics data and define startDate for youtubeAnalyticsReportsQuery_()
    // If the value returned for getLatestDate_() is null, i.e., there are no previous dates recorded in targetSheet,
    // latestDate will be Dec 31 of the previous year of targetYear
    let currentLatestOnSpreadsheet = getLatestDate_(targetSheet, 1); // Assuming that the date is recorded on column A of the targetSheet.
    let latestDate = (currentLatestOnSpreadsheet ? new Date(currentLatestOnSpreadsheet.getTime()) : new Date(targetYear - 1, 11, 31));
    let startDateObj = new Date(latestDate.setDate(latestDate.getDate() + 1));
    // Setting parameters for youtubeAnalyticsReportsQuery_()
    let startDate = formattedDateAnalytics_(startDateObj);
    let endDateObj = new Date(now.getTime());
    let endDate = formattedDateAnalytics_(endDateObj);
    let metrics = 'views,likes,dislikes,subscribersGained,subscribersLost,estimatedMinutesWatched,averageViewDuration,cardImpressions,cardClicks';
    let ids = 'channel==MINE';
    let options = {
      dimensions: 'day',
      filters: ''
    };
    // Get full list of videos owned by the authorized user
    let videoList = youtubeMyVideoList_();
    // Get analytics data
    let data = videoList.reduce((accData, video) => {
      let videoId = video.id.videoId;
      let channelId = video.snippet.channelId;
      options.filters = `video==${videoId}`;
      let rawAnalytics = JSON.parse(youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids, options)).rows;
      let analytics = rawAnalytics.reduce((accAnalytics, dayVideo) => {
        let yearMonth = dayVideo[0].slice(0, 7);
        if (yearLimit && dayVideo[0] <= `${targetYear}-12-31` || !yearLimit) {
          // Assuming that the first column is the 'day' column, insert channel ID and video ID to the day-by-day data array
          dayVideo.splice(1, 0, channelId, videoId);
          let concatDayVideo = dayVideo.concat([yearMonth]);
          accAnalytics.dailyVideoAnalytics.push(concatDayVideo);
        }
        if (dayVideo[0] > accAnalytics.latest) {
          accAnalytics.latest = dayVideo[0];
        }
        return accAnalytics;
      }, { dailyVideoAnalytics: [], latest: `${targetYear}-01-01` });
      let acc = accData.videoData.slice();
      accData.videoData = acc.concat(analytics.dailyVideoAnalytics);
      if (analytics.latest > accData.latest) {
        accData.latest = analytics.latest;
      }
      return accData;
    }, { videoData: [], latest: `${targetYear}-01-01` });
    let updatedLatestDateObj = {
      latestDateOnSpreadsheet: currentLatestOnSpreadsheet,
      latestDateReturned: currentLatestOnSpreadsheet
    };
    if (data.videoData && data.videoData.length > 0) {
      // Copy on spreadsheet
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, data.videoData.length, data.videoData[0].length).setValues(data.videoData);
      // Get latest updated date
      updatedLatestDateObj.latestDateOnSpreadsheet = getLatestDate_(targetSheet, 1);
      updatedLatestDateObj.latestDateReturned = yMd2Date_(data.latest);
      // Log
      enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, localizedMessages.replaceUpdatedYouTubeVideoAnalyticsLog(startDate, formattedDateAnalytics_(updatedLatestDateObj.latestDateOnSpreadsheet)), now);
    } else {
      enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, localizedMessages.messageList.youtube.noUpdatesForYouTubeVideoAnalyticsLog, now);
    }
    return updatedLatestDateObj;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheet.getUrl(), LOG_SHEET_NAME, message, now);
    throw error;
  }
}

/**
 * GET request using YouTube Analytics API reports.query.
 * https://developers.google.com/youtube/analytics/reference/reports/query
 * @param {string} startDate The start date for fetching YouTube Analytics data. The value should be in YYYY-MM-DD format.
 * @param {string} endDate The end date for fetching YouTube Analytics data. The value should be in YYYY-MM-DD format.
 * @param {string} metrics A comma-separated list of YouTube Analytics metrics, such as views, likes, and dislikes.
 * See https://developers.google.com/youtube/analytics/metrics for a full list.
 * @param {string} ids [Optional] Identifies the YouTube channel or content owner for which you are retrieving YouTube Analytics data.
 * Defaults to 'channel==MINE' to represent the authorized user's channel(s).
 * @param {Object} options [Optional] A JavaScript object for additional parameters outlined in https://developers.google.com/youtube/analytics/reference/reports/query#Parameters
 */
function youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids = 'channel==MINE', options = null) {
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var parameters = {
    'startDate': startDate,
    'endDate': endDate,
    'metrics': metrics,
    'ids': ids,
  };
  if (options) {
    for (let k in options) {
      parameters[k] = options[k];
    }
  }
  var youtubeAPIService = getYouTubeAPIService_();
  if (!youtubeAPIService.hasAccess()) {
    throw new Error(localizedMessages.messageList.youtube.errorUnauthorized);
  }
  var baseUrl = `https://youtubeanalytics.googleapis.com/${YOUTUBE_ANALYTICS_API_VERSION}/reports`;
  var paramString = '?';
  for (let k in parameters) {
    let param = `${k}=${encodeURIComponent(parameters[k])}`;
    if (paramString.slice(-1) !== '?') {
      param = '&' + param;
    }
    paramString += param;
  }
  var url = baseUrl + paramString;
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
    contentType: 'application/json; charset=UTF-8',
    headers: {
      Authorization: 'Bearer ' + encodeURIComponent(youtubeAPIService.getAccessToken())
    }
  });
  return response;
}

//////////////////////////////
// Create Analytics Summary //
//////////////////////////////

/**
 * Update YouTube analytics summary sheet of the spreadsheet.
 * The range designated in this script assumes the cells to be positioned in the layout of the sample spreadsheet.
 * [URL of sample spreadsheet] /////////////////////
 */
function createYouTubeAnalyticsSummary() {
  console.info('[createYouTubeAnalyticsSummary] Initiating createYouTubeAnalyticsSummary: Creating a monthly analytics summary report for the designated YouTube channel and year-month...'); // log
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssTimeZone = ss.getSpreadsheetTimeZone();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  var config = getConfig_();
  var summarySheet = ss.getSheetByName(config.SHEET_NAME_YOUTUBE_SUMMARY);
  var videoSheet = ss.getSheetByName(config.SHEET_NAME_YOUTUBE_VIDEOS);
  var videoSheetRowOffset = 4; // Starting row number for data in videoSheet
  var videoSheetColOffset = 2; // Starting column number for data in videoSheet
  var now = new Date();
  try {
    // Channel ID
    let targetChannelId = summarySheet.getRange(3, 5).getValue(); // Assuming that the channel ID is entered in this cell
    // Get list of YouTube channel(s) that the authorized user owns; for testing.
    let myChannelListFull = youtubeMyChannelList_();
    let myChannelIds = myChannelListFull.map(element => element.id);
    if (!targetChannelId) {
      throw new Error(localizedMessages.messageList.youtube.errorNoTextEnteredForChannelId);
    } else if (!myChannelIds.includes(targetChannelId)) {
      throw new Error(localizedMessages.replaceErrorInvalidChannelId(targetChannelId));
    }
    let targetChannelName = myChannelListFull.filter(element => element.id == targetChannelId)[0].snippet.title;
    // Report month
    let reportMonth = summarySheet.getRange(4, 5).getValue(); // Assuming that the target report month is entered in this cell
    if (!reportMonth) {
      throw new Error(localizedMessages.messageList.youtube.errorNoTextEnteredForReportMonth);
    } else if (!reportMonth.match(/^\d{4}-\d{2}$/)) {
      throw new Error(localizedMessages.replaceErrorInvalidReportMonth(reportMonth));
    }
    console.info(`[createYouTubeAnalyticsSummary] Verified channel ID and report month; creating report for YouTube channel "${targetChannelName}" for year-month of ${reportMonth}...`); // log
    let targetPeriodStartPre = new Date(parseInt(reportMonth.slice(0, 4)) - 2, parseInt(reportMonth.slice(-2)) - 1, 1);
    let targetPeriodStart = new Date(targetPeriodStartPre.setMonth(targetPeriodStartPre.getMonth() + 1));
    let reportPeriodStartPre = new Date(parseInt(reportMonth.slice(0, 4)) - 1, parseInt(reportMonth.slice(-2)) - 1, 1);
    let reportPeriodStart = new Date(reportPeriodStartPre.setMonth(reportPeriodStartPre.getMonth() + 1));
    let targetPeriodEndPre = new Date(reportMonth.slice(0, 4), parseInt(reportMonth.slice(-2)) - 1, 1);
    targetPeriodEndPre.setMonth(targetPeriodEndPre.getMonth() + 1);
    let targetPeriodEnd = new Date(targetPeriodEndPre.setDate(targetPeriodEndPre.getDate() - 1));
    // Update Channel/Video list
    let muteUiAlert = true;
    console.info(`[createYouTubeAnalyticsSummary] UI alert is ${muteUiAlert ? 'muted' : 'not muted'}.`); // log
    let channelList = updateYouTubeSummaryChannelList(muteUiAlert);
    let videoList = updateYouTubeSummaryVideoList(muteUiAlert);
    let channelListTarget = channelList.filter(element => element[4] == targetChannelId);
    let channelListSummary = channelListTarget.map(element => [
      element[2], // thumbnail image function
      element[5], // channel name
      element[6], // channel description
      element[8], // channel publish date (local time)
      element[9], // view count
      element[10], // subscriber count
      element[11], // video count
      now // timestamp
    ]);
    summarySheet.getRange(7, 3, 1, channelListSummary[0].length).setValues(channelListSummary);
    // Get background data from the spreadsheet(s) of the target period
    let bgDataChannelAnalytics = [];
    let bgDataChannelDemographics = [];
    let bgDataVideoAnalytics = [];
    let channelAnalyticsDataHeader = [];
    let channelDemographicsDataHeader = [];
    let videoAnalyticsDataHeader = [];
    let spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
    console.info('[createYouTubeAnalyticsSummary] Updated channel summary. Scanning the list of spreadsheets for retrieving background data...'); // log
    for (let i = 0; i < spreadsheetList.length; i++) {
      let row = spreadsheetList[i];
      if (row.PLATFORM != 'YouTube' || row.YEAR > targetPeriodEnd.getFullYear()) {
        continue;
      }
      let dataSpreadsheetUrl = row.URL;
      let dataSpreadsheet = SpreadsheetApp.openByUrl(dataSpreadsheetUrl);
      // Get channel analytics
      let channelAnalyticsSheet = dataSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_ANALYTICS);
      let channelAnalyticsDataFull = channelAnalyticsSheet.getDataRange().getValues();
      let channelAnalyticsData = [];
      channelAnalyticsDataHeader = channelAnalyticsDataFull.shift();
      // Get channel demographics
      let channelDemographicsSheet = dataSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_DEMOGRAPHICS);
      let channelDemographicsDataFull = channelDemographicsSheet.getDataRange().getValues();
      let channelDemographicsData = [];
      channelDemographicsDataHeader = channelDemographicsDataFull.shift();
      // Get video analytics
      let videoAnalyticsSheet = dataSpreadsheet.getSheetByName(config.SHEET_NAME_VIDEO_ANALYTICS);
      let videoAnalyticsDataFull = videoAnalyticsSheet.getDataRange().getValues();
      let videoAnalyticsData = [];
      videoAnalyticsDataHeader = videoAnalyticsDataFull.shift();
      if (row.YEAR < targetPeriodStart.getFullYear()) {
        videoAnalyticsData = videoAnalyticsDataFull.filter(element => element[1] == targetChannelId); // Assuming that the second column of the video analytics data table is the channel ID
      } else if (row.YEAR == targetPeriodStart.getFullYear()) {
        channelAnalyticsData = channelAnalyticsDataFull.filter(element => {
          // Assuming that the first column of the channel analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          // Assuming that the second column of the channel analytics data table is the channel ID
          let channelId = element[1];
          return thisDate.getTime() >= targetPeriodStart.getTime() && channelId == targetChannelId;
        });
        channelDemographicsData = channelDemographicsDataFull.filter(element => {
          // Assuming that the first column of the channel demographics data table is the year-month in yyyy-MM
          let thisDate = yearMonth2Date_(element[0]);
          // Assuming that the second column of the channel demographics data table is the channel ID
          let channelId = element[1];
          // For demographics, return data for only the report period, not the full target period.
          return thisDate.getTime() >= reportPeriodStart.getTime() && channelId == targetChannelId;
        });
        videoAnalyticsData = videoAnalyticsDataFull.filter(element => element[1] == targetChannelId); // Assuming that the second column of the video analytics data table is the channel ID
      } else if (row.YEAR > targetPeriodStart.getFullYear() && row.YEAR < targetPeriodEnd.getFullYear()) {
        channelAnalyticsData = channelAnalyticsDataFull.slice();
        channelDemographicsData = channelDemographicsDataFull.filter(element => {
          // Assuming that the first column of the channel demographics data table is the year-month in yyyy-MM
          let thisDate = yearMonth2Date_(element[0]);
          // Assuming that the second column of the channel demographics data table is the channel ID
          let channelId = element[1];
          // For demographics, return data for only the report period, not the full target period.
          return thisDate.getTime() >= reportPeriodStart.getTime() && channelId == targetChannelId;
        });
        videoAnalyticsData = videoAnalyticsDataFull.filter(element => element[1] == targetChannelId); // Assuming that the second column of the video analytics data table is the channel ID
      } else if (row.YEAR == targetPeriodEnd.getFullYear()) {
        channelAnalyticsData = channelAnalyticsDataFull.filter(element => {
          // Assuming that the first column of the channel analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          // Assuming that the second column of the channel analytics data table is the channel ID
          let channelId = element[1];
          return thisDate.getTime() <= targetPeriodEnd.getTime() && channelId == targetChannelId;
        });
        channelDemographicsData = channelDemographicsDataFull.filter(element => {
          // Assuming that the first column of the channel demographics data table is the year-month in yyyy-MM
          let thisDate = yearMonth2Date_(element[0]);
          // Assuming that the second column of the channel demographics data table is the channel ID
          let channelId = element[1];
          return thisDate.getTime() <= targetPeriodEnd.getTime() && channelId == targetChannelId;
        });
        videoAnalyticsData = videoAnalyticsDataFull.filter(element => {
          // Assuming that the first column of the video analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          // Assuming that the second column of the video analytics data table is the channel ID
          let channelId = element[1];
          return thisDate.getTime() <= targetPeriodEnd.getTime() && channelId == targetChannelId;
        });
      } else {
        continue;
      }
      // Concatenate
      let newBgDataChannelAnalytics = bgDataChannelAnalytics.concat(channelAnalyticsData);
      bgDataChannelAnalytics = newBgDataChannelAnalytics.slice();
      let newBgDataChannelDemographics = bgDataChannelDemographics.concat(channelDemographicsData);
      bgDataChannelDemographics = newBgDataChannelDemographics.slice();
      let newBgDataVideoAnalytics = bgDataVideoAnalytics.concat(videoAnalyticsData);
      bgDataVideoAnalytics = newBgDataVideoAnalytics.slice();
    }
    // Add headers to the obtained data
    bgDataChannelAnalytics.unshift(channelAnalyticsDataHeader);
    bgDataChannelDemographics.unshift(channelDemographicsDataHeader);
    bgDataVideoAnalytics.unshift(videoAnalyticsDataHeader);
    // Process data for report
    console.info('[createYouTubeAnalyticsSummary] Background data retrieved. Processing data for report...'); // log
    let bgDataChannelAnalyticsMod = bgDataChannelAnalytics.map((element, index) => {
      if (index == 0) {
        let concatElement = element.concat(['CURRENT-YEAR', 'YEAR-MONTH_AGG', 'DISLIKES_INV', 'SUBSCRIBERS TOTAL']);
        return concatElement;
      } else {
        // Assuming that the first column of the channel analytics data table is the date
        let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
        let [ytStatus, yearMonthAgg] = dateStatus_(thisDate, reportPeriodStart, ssTimeZone, targetPeriodStart);
        let dislikesInv = -parseInt(element[4]); // Invert postive counts of dislikes to negative for better visualization
        let subscribersTotal = parseInt(element[5]) - parseInt(element[6]); // [SUBSCRIBERS GAINED] - [SUBSCRIBERS LOST]
        // Add column(s) to the original data
        let concatElement = element.concat([ytStatus, yearMonthAgg, dislikesInv, subscribersTotal]);
        return concatElement;
      }
    });
    let bgDataVideoAnalyticsMod = bgDataVideoAnalytics.map((element, index) => {
      if (index == 0) {
        let concatElement = element.concat(['DISLIKES_INV', 'SUBSCRIBERS TOTAL', 'ESTIMATED SECONDS WATCHED', 'CURRENT-YEAR', 'YEAR-MONTH_AGG']);
        return concatElement;
      } else {
        let dislikesInv = -parseInt(element[5]); // Invert postive counts of dislikes to negative for better visualization
        let subscribersTotal = parseInt(element[6]) - parseInt(element[7]); // [SUBSCRIBERS GAINED] - [SUBSCRIBERS LOST]
        let estimatedSecWatchedLatest = element[8] * 60; // Convert [ESTIMATED MINUTES WATCHED] into seconds
        // Assuming that the first column of the channel analytics data table is the date
        let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
        let [ytStatus, yearMonthAgg] = dateStatus_(thisDate, reportPeriodStart, ssTimeZone, targetPeriodStart);
        // Add column(s) to the original data
        let concatElement = element.concat([dislikesInv, subscribersTotal, estimatedSecWatchedLatest, ytStatus, yearMonthAgg]);
        return concatElement;
      }
    });
    // Delete existing temporary data and copy the new data on spreadsheet
    console.info('[createYouTubeAnalyticsSummary] Data procession complete. Copying data on spreadsheet...'); // log
    let tempSheetChannelAnalytics = ss.getSheetByName(config.SHEET_NAME_TEMP_CHANNEL_ANALYTICS);
    let tempSheetChannelDemographics = ss.getSheetByName(config.SHEET_NAME_TEMP_CHANNEL_DEMOGRAPHICS);
    let tempSheetVideoAnalytics = ss.getSheetByName(config.SHEET_NAME_TEMP_VIDEO_ANALYTICS);
    //// Delete existing data
    tempSheetChannelAnalytics.getDataRange().deleteCells(SpreadsheetApp.Dimension.ROWS);
    tempSheetChannelDemographics.getDataRange().deleteCells(SpreadsheetApp.Dimension.ROWS);
    tempSheetVideoAnalytics.getDataRange().deleteCells(SpreadsheetApp.Dimension.ROWS);
    //// Copy new data on spreadsheet
    tempSheetChannelAnalytics.getRange(1, 1, bgDataChannelAnalyticsMod.length, bgDataChannelAnalyticsMod[0].length)
      .setValues(bgDataChannelAnalyticsMod);
    tempSheetChannelDemographics.getRange(1, 1, bgDataChannelDemographics.length, bgDataChannelDemographics[0].length)
      .setValues(bgDataChannelDemographics);
    tempSheetVideoAnalytics.getRange(1, 1, bgDataVideoAnalyticsMod.length, bgDataVideoAnalyticsMod[0].length)
      .setValues(bgDataVideoAnalyticsMod);
    console.info('[createYouTubeAnalyticsSummary] Basic analytics data is copied on spreadsheet. Creating advanced statistics for each video...'); // log
    // Convert bgDataVideoAnalyticsMod into an array of objects grouped by video ID
    let bgDataVideoAnalyticsObj = groupArray_(bgDataVideoAnalyticsMod, 'VIDEO ID');
    // Key dates for advanced statistics based on reporting period
    let baseDate = new Date(reportMonth.slice(0, 4), parseInt(reportMonth.slice(-2)) - 1, 1);
    let targetPeriodStartString = formattedDateAnalytics_(targetPeriodStart, 'PST');
    let reportPeriodStartString = formattedDateAnalytics_(reportPeriodStart, 'PST');
    let targetPeriodEndString = formattedDateAnalytics_(targetPeriodEnd, 'PST');
    let reportMonthStartString = formattedDateAnalytics_(baseDate, 'PST');
    let prevMonthStartString = formattedDateAnalytics_(new Date(baseDate.setMonth(baseDate.getMonth() - 1)), 'PST');
    let prevYearSameMonthStartString = formattedDateAnalytics_(new Date(parseInt(reportMonth.slice(0, 4)) - 1, parseInt(reportMonth.slice(-2)) - 1, 1), 'PST');
    // Create modified list of authenticated user's videos with advanced statistics
    let videoListMod = videoList.map(element => {
      // Name the values in element
      let [timestamp, num, thumbnailUrlFunction, thumbnailUrl, /*channelId*/, /*channelName*/, videoId, videoTitle, videoDesc, videoUrl, duration, caption, publishedAtUtc, publishedAtLocal, privacyStatus, latestViewCount, latestLikeCount, latestDislikeCount] = element;
      let durationSec = iso8601duration2sec_(duration);
      // Key dates for advanced statistics based on publication period
      let publishedAt = new Date(publishedAtUtc);
      let week1 = new Date(publishedAt.setDate(publishedAt.getDate() + 7));
      let week1string = formattedDateAnalytics_(week1, 'PST');
      let week4 = new Date(publishedAt.setDate(publishedAt.getDate() + 21)); // 7 + 21 = 28 (= 7 * 4)
      let week4string = formattedDateAnalytics_(week4, 'PST');
      let week12 = new Date(publishedAt.setDate(publishedAt.getDate() * 56)); // 7 + 21 + 56 = 84 (= 7 * 12)
      let week12string = formattedDateAnalytics_(week12, 'PST');
      // Objects to store the advanced statistics
      //// Aggregated statistics for the first 7 days after the respective videos were released,
      //// i.e., published or created at 'unlisted' status.
      let week1Stats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the first 28 days after the respective videos were released.
      let week4Stats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the first 84 days after the respective videos were released.
      let week12Stats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the respective videos during the period of
      //// on and after targetPeriodStartString date
      //// and before reportPeriodStartString Date
      let prevYearAggStats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the respective videos during the inclusive period 
      //// between reportPeriodStartString and targetPeriodEndString dates.
      let curYearAggStats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the respective videos during the inclusive period 
      //// between reportMonthStartString and targetPeriodEndString dates.
      let curMonthAggStats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the respective videos during the period 
      //// starting from prevMonthStartString to before reportMonthStartString dates.
      let prevMonthAggStats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      //// Aggregated statistics for the respective videos during the period 
      //// starting from prevYearSameMonthStart to before reportPeriodStartString dates.
      let prevYearSameMonthAggStats = {
        'days': [],
        'views': 0,
        'likes': 0,
        'dislikes': 0,
        'subscribersGained': 0,
        'subscribersLost': 0,
        'estimatedSecWatched': 0,
        'cardImpression': 0,
        'cardClicks': 0
      };
      // Advanced statistics using bgDataVideoAnalyticsMod
      let thisVideoAnalytics = bgDataVideoAnalyticsObj[videoId];
      thisVideoAnalytics.forEach(dailyData => {
        // Aggregate advanced statistics based on publication period
        if (dailyData['DAY'] <= week1string) {
          week1Stats.days.push(dailyData['DAY']);
          week4Stats.days.push(dailyData['DAY']);
          week12Stats.days.push(dailyData['DAY']);
          // Views
          week1Stats.views += dailyData['VIEWS'];
          week4Stats.views += dailyData['VIEWS'];
          week12Stats.views += dailyData['VIEWS'];
          // Likes
          week1Stats.likes += dailyData['LIKES'];
          week4Stats.likes += dailyData['LIKES'];
          week12Stats.likes += dailyData['LIKES'];
          // Dislikes
          week1Stats.dislikes += dailyData['DISLIKES'];
          week4Stats.dislikes += dailyData['DISLIKES'];
          week12Stats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          week1Stats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          week4Stats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          week12Stats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          week1Stats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          week4Stats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          week12Stats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          week1Stats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          week4Stats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          week12Stats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          week1Stats.cardImpression += dailyData['CARD IMPRESSIONS'];
          week4Stats.cardImpression += dailyData['CARD IMPRESSIONS'];
          week12Stats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          week1Stats.cardClicks += dailyData['CARD CLICKS'];
          week4Stats.cardClicks += dailyData['CARD CLICKS'];
          week12Stats.cardClicks += dailyData['CARD CLICKS'];
        } else if (dailyData['DAY'] > week1string && dailyData['DAY'] <= week4string) {
          week4Stats.days.push(dailyData['DAY']);
          week12Stats.days.push(dailyData['DAY']);
          // Views
          week4Stats.views += dailyData['VIEWS'];
          week12Stats.views += dailyData['VIEWS'];
          // Likes
          week4Stats.likes += dailyData['LIKES'];
          week12Stats.likes += dailyData['LIKES'];
          // Dislikes
          week4Stats.dislikes += dailyData['DISLIKES'];
          week12Stats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          week4Stats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          week12Stats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          week4Stats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          week12Stats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          week4Stats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          week12Stats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          week4Stats.cardImpression += dailyData['CARD IMPRESSIONS'];
          week12Stats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          week4Stats.cardClicks += dailyData['CARD CLICKS'];
          week12Stats.cardClicks += dailyData['CARD CLICKS'];
        } else if (dailyData['DAY'] > week4string && dailyData['DAY'] <= week12string) {
          week12Stats.days.push(dailyData['DAY']);
          // Views
          week12Stats.views += dailyData['VIEWS'];
          // Likes
          week12Stats.likes += dailyData['LIKES'];
          // Dislikes
          week12Stats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          week12Stats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          week12Stats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          week12Stats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          week12Stats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          week12Stats.cardClicks += dailyData['CARD CLICKS'];
        }
        // Aggregate advanced statistics based on reporting period
        if (dailyData['DAY'] >= targetPeriodStartString && dailyData['DAY'] < prevYearSameMonthStartString) {
          prevYearAggStats.days.push(dailyData['DAY']);
          // Views
          prevYearAggStats.views += dailyData['VIEWS'];
          // Likes
          prevYearAggStats.likes += dailyData['LIKES'];
          // Dislikes
          prevYearAggStats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          prevYearAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          prevYearAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          prevYearAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          prevYearAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          prevYearAggStats.cardClicks += dailyData['CARD CLICKS'];
        } else if (dailyData['DAY'] >= prevYearSameMonthStartString && dailyData['DAY'] < reportPeriodStartString) {
          prevYearAggStats.days.push(dailyData['DAY']);
          prevYearSameMonthAggStats.days.push(dailyData['DAY']);
          // Views
          prevYearAggStats.views += dailyData['VIEWS'];
          prevYearSameMonthAggStats.views += dailyData['VIEWS'];
          // Likes
          prevYearAggStats.likes += dailyData['LIKES'];
          prevYearSameMonthAggStats.likes += dailyData['LIKES'];
          // Dislikes
          prevYearAggStats.dislikes += dailyData['DISLIKES'];
          prevYearSameMonthAggStats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          prevYearAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          prevYearSameMonthAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          prevYearAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          prevYearSameMonthAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          prevYearAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          prevYearSameMonthAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          prevYearAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          prevYearSameMonthAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          prevYearAggStats.cardClicks += dailyData['CARD CLICKS'];
          prevYearSameMonthAggStats.cardClicks += dailyData['CARD CLICKS'];
        } else if (dailyData['DAY'] >= reportPeriodStartString && dailyData['DAY'] < prevMonthStartString) {
          curYearAggStats.days.push(dailyData['DAY']);
          // Views
          curYearAggStats.views += dailyData['VIEWS'];
          // Likes
          curYearAggStats.likes += dailyData['LIKES'];
          // Dislikes
          curYearAggStats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          curYearAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          curYearAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          curYearAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          curYearAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          curYearAggStats.cardClicks += dailyData['CARD CLICKS'];
        } else if (dailyData['DAY'] >= prevMonthStartString && dailyData['DAY'] < reportMonthStartString) {
          curYearAggStats.days.push(dailyData['DAY']);
          prevMonthAggStats.days.push(dailyData['DAY']);
          // Views
          curYearAggStats.views += dailyData['VIEWS'];
          prevMonthAggStats.views += dailyData['VIEWS'];
          // Likes
          curYearAggStats.likes += dailyData['LIKES'];
          prevMonthAggStats.likes += dailyData['LIKES'];
          // Dislikes
          curYearAggStats.dislikes += dailyData['DISLIKES'];
          prevMonthAggStats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          curYearAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          prevMonthAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          curYearAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          prevMonthAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          curYearAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          prevMonthAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          curYearAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          prevMonthAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          curYearAggStats.cardClicks += dailyData['CARD CLICKS'];
          prevMonthAggStats.cardClicks += dailyData['CARD CLICKS'];
        } else if (dailyData['DAY'] >= reportMonthStartString && dailyData['DAY'] <= targetPeriodEndString) {
          curYearAggStats.days.push(dailyData['DAY']);
          curMonthAggStats.days.push(dailyData['DAY']);
          // Views
          curYearAggStats.views += dailyData['VIEWS'];
          curMonthAggStats.views += dailyData['VIEWS'];
          // Likes
          curYearAggStats.likes += dailyData['LIKES'];
          curMonthAggStats.likes += dailyData['LIKES'];
          // Dislikes
          curYearAggStats.dislikes += dailyData['DISLIKES'];
          curMonthAggStats.dislikes += dailyData['DISLIKES'];
          // Subscribers Gained
          curYearAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          curMonthAggStats.subscribersGained += dailyData['SUBSCRIBERS GAINED'];
          // Subscribers Lost
          curYearAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          curMonthAggStats.subscribersLost += dailyData['SUBSCRIBERS LOST'];
          // Estimated Seconds Watched
          curYearAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          curMonthAggStats.estimatedSecWatched += (dailyData['ESTIMATED MINUTES WATCHED'] * 60);
          // Card Impression
          curYearAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          curMonthAggStats.cardImpression += dailyData['CARD IMPRESSIONS'];
          // Card Clicks
          curYearAggStats.cardClicks += dailyData['CARD CLICKS'];
          curMonthAggStats.cardClicks += dailyData['CARD CLICKS'];
        } else {
          return;
        }
      });
      let week1AverageViewPercentage = 100 * ((week1Stats.estimatedSecWatched / week1Stats.views) / durationSec);
      let week4AverageViewPercentage = 100 * ((week4Stats.estimatedSecWatched / week4Stats.views) / durationSec);
      let week12AverageViewPercentage = 100 * ((week12Stats.estimatedSecWatched / week12Stats.views) / durationSec);
      let prevYearAggStatsAverageViewPercentage = 100 * ((prevYearAggStats.estimatedSecWatched / prevYearAggStats.views) / durationSec);
      let curYearAggStatsAverageViewPercentage = 100 * ((curYearAggStats.estimatedSecWatched / curYearAggStats.views) / durationSec);
      let curMonthAggStatsAverageViewPercentage = 100 * ((curMonthAggStats.estimatedSecWatched / curMonthAggStats.views) / durationSec);
      let prevMonthAggStatsAverageViewPercentage = 100 * ((prevMonthAggStats.estimatedSecWatched / prevMonthAggStats.views) / durationSec);
      let prevYearSameMonthAggStatsAverageViewPercentage = 100 * ((prevYearSameMonthAggStats.estimatedSecWatched / prevYearSameMonthAggStats.views) / durationSec);
      let elementMod = [
        num,
        thumbnailUrlFunction,
        thumbnailUrl,
        videoId,
        videoTitle,
        videoDesc,
        videoUrl,
        durationSec,
        caption,
        publishedAtUtc,
        publishedAtLocal,
        privacyStatus,
        latestViewCount,
        latestLikeCount,
        latestDislikeCount,
        // Week 1 statistics
        week1Stats.views,
        week1Stats.likes,
        week1Stats.dislikes,
        week1Stats.subscribersGained,
        week1Stats.subscribersLost,
        week1Stats.estimatedSecWatched,
        (week1AverageViewPercentage >= 100 ? Number.parseFloat(week1AverageViewPercentage).toPrecision(4) : Number.parseFloat(week1AverageViewPercentage).toPrecision(3)),
        week1Stats.cardImpression,
        week1Stats.cardClicks,
        // Week 4 statistics
        week4Stats.views,
        week4Stats.likes,
        week4Stats.dislikes,
        week4Stats.subscribersGained,
        week4Stats.subscribersLost,
        week4Stats.estimatedSecWatched,
        (week4AverageViewPercentage >= 100 ? Number.parseFloat(week4AverageViewPercentage).toPrecision(4) : Number.parseFloat(week4AverageViewPercentage).toPrecision(3)),
        week4Stats.cardImpression,
        week4Stats.cardClicks,
        // Week 12 statistics
        week12Stats.views,
        week12Stats.likes,
        week12Stats.dislikes,
        week12Stats.subscribersGained,
        week12Stats.subscribersLost,
        week12Stats.estimatedSecWatched,
        (week12AverageViewPercentage >= 100 ? Number.parseFloat(week12AverageViewPercentage).toPrecision(4) : Number.parseFloat(week12AverageViewPercentage).toPrecision(3)),
        week12Stats.cardImpression,
        week12Stats.cardClicks,
        // Current month statistics
        curMonthAggStats.views,
        curMonthAggStats.likes,
        curMonthAggStats.dislikes,
        curMonthAggStats.subscribersGained,
        curMonthAggStats.subscribersLost,
        curMonthAggStats.estimatedSecWatched,
        (curMonthAggStatsAverageViewPercentage >= 100 ? Number.parseFloat(curMonthAggStatsAverageViewPercentage).toPrecision(4) : Number.parseFloat(curMonthAggStatsAverageViewPercentage).toPrecision(3)),
        curMonthAggStats.cardImpression,
        curMonthAggStats.cardClicks,
        // Previous month statistics
        prevMonthAggStats.views,
        prevMonthAggStats.likes,
        prevMonthAggStats.dislikes,
        prevMonthAggStats.subscribersGained,
        prevMonthAggStats.subscribersLost,
        prevMonthAggStats.estimatedSecWatched,
        (prevMonthAggStatsAverageViewPercentage >= 100 ? Number.parseFloat(prevMonthAggStatsAverageViewPercentage).toPrecision(4) : Number.parseFloat(prevMonthAggStatsAverageViewPercentage).toPrecision(3)),
        prevMonthAggStats.cardImpression,
        prevMonthAggStats.cardClicks,
        // Previous year, same month statistics
        prevYearSameMonthAggStats.views,
        prevYearSameMonthAggStats.likes,
        prevYearSameMonthAggStats.dislikes,
        prevYearSameMonthAggStats.subscribersGained,
        prevYearSameMonthAggStats.subscribersLost,
        prevYearSameMonthAggStats.estimatedSecWatched,
        (prevYearSameMonthAggStatsAverageViewPercentage >= 100 ? Number.parseFloat(prevYearSameMonthAggStatsAverageViewPercentage).toPrecision(4) : Number.parseFloat(prevYearSameMonthAggStatsAverageViewPercentage).toPrecision(3)),
        prevYearSameMonthAggStats.cardImpression,
        prevYearSameMonthAggStats.cardClicks,
        // Current year aggregated statistics
        curYearAggStats.views,
        curYearAggStats.likes,
        curYearAggStats.dislikes,
        curYearAggStats.subscribersGained,
        curYearAggStats.subscribersLost,
        curYearAggStats.estimatedSecWatched,
        (curYearAggStatsAverageViewPercentage >= 100 ? Number.parseFloat(curYearAggStatsAverageViewPercentage).toPrecision(4) : Number.parseFloat(curYearAggStatsAverageViewPercentage).toPrecision(3)),
        curYearAggStats.cardImpression,
        curYearAggStats.cardClicks,
        // Previous year aggregated statistics
        prevYearAggStats.views,
        prevYearAggStats.likes,
        prevYearAggStats.dislikes,
        prevYearAggStats.subscribersGained,
        prevYearAggStats.subscribersLost,
        prevYearAggStats.estimatedSecWatched,
        (prevYearAggStatsAverageViewPercentage >= 100 ? Number.parseFloat(prevYearAggStatsAverageViewPercentage).toPrecision(4) : Number.parseFloat(prevYearAggStatsAverageViewPercentage).toPrecision(3)),
        prevYearAggStats.cardImpression,
        prevYearAggStats.cardClicks,
        // Month-on-month diffs
        curMonthAggStats.views - prevMonthAggStats.views,
        curMonthAggStats.likes - prevMonthAggStats.likes,
        curMonthAggStats.dislikes - prevMonthAggStats.dislikes,
        curMonthAggStats.subscribersGained - prevMonthAggStats.subscribersGained,
        curMonthAggStats.subscribersLost - prevMonthAggStats.subscribersLost,
        curMonthAggStats.estimatedSecWatched - prevMonthAggStats.estimatedSecWatched,
        Number.parseFloat(curMonthAggStatsAverageViewPercentage - prevMonthAggStatsAverageViewPercentage).toPrecision(4),
        curMonthAggStats.cardImpression - prevMonthAggStats.cardImpression,
        curMonthAggStats.cardClicks - prevMonthAggStats.cardClicks,
        // Same month last year diffs
        curMonthAggStats.views - prevYearSameMonthAggStats.views,
        curMonthAggStats.likes - prevYearSameMonthAggStats.likes,
        curMonthAggStats.dislikes - prevYearSameMonthAggStats.dislikes,
        curMonthAggStats.subscribersGained - prevYearSameMonthAggStats.subscribersGained,
        curMonthAggStats.subscribersLost - prevYearSameMonthAggStats.subscribersLost,
        curMonthAggStats.estimatedSecWatched - prevYearSameMonthAggStats.estimatedSecWatched,
        Number.parseFloat(curMonthAggStatsAverageViewPercentage - prevYearSameMonthAggStatsAverageViewPercentage).toPrecision(4),
        curMonthAggStats.cardImpression - prevYearSameMonthAggStats.cardImpression,
        curMonthAggStats.cardClicks - prevYearSameMonthAggStats.cardClicks,
        // Year-on-year diffs
        curYearAggStats.views - prevYearAggStats.views,
        curYearAggStats.likes - prevYearAggStats.likes,
        curYearAggStats.dislikes - prevYearAggStats.dislikes,
        curYearAggStats.subscribersGained - prevYearAggStats.subscribersGained,
        curYearAggStats.subscribersLost - prevYearAggStats.subscribersLost,
        curYearAggStats.estimatedSecWatched - prevYearAggStats.estimatedSecWatched,
        Number.parseFloat(curYearAggStatsAverageViewPercentage - prevYearAggStatsAverageViewPercentage).toPrecision(4),
        curYearAggStats.cardImpression - prevYearAggStats.cardImpression,
        curYearAggStats.cardClicks - prevYearAggStats.cardClicks,
        timestamp
      ];
      return elementMod;
    });
    // Delete existing data
    console.info('[createYouTubeAnalyticsSummary] Created advanced statistics. Deleting existing data and copying the new statistics to spreadsheet...'); // log
    videoSheet.getRange(videoSheetRowOffset, videoSheetColOffset, videoSheet.getLastRow() - videoSheetRowOffset + 1, videoSheet.getLastColumn() - videoSheetColOffset + 1)
      .deleteCells(SpreadsheetApp.Dimension.ROWS);
    // Enter new data
    videoSheet.getRange(videoSheetRowOffset, videoSheetColOffset, videoListMod.length, videoListMod[0].length)
      .setValues(videoListMod);
    let scriptEnd = new Date();
    let scriptExeTime = (scriptEnd.getTime() - now.getTime()) / 1000;
    console.info(`[createYouTubeAnalyticsSummary] Complete: Report for ${reportMonth} of YouTube channel "${targetChannelName}" created.\nScript Time: ${scriptExeTime} secs.`); // log
    ui.alert(localizedMessages.replaceReportCreated(reportMonth, targetChannelName, scriptExeTime));
  } catch (error) {
    let message = errorMessage_(error);
    console.info('[createYouTubeAnalyticsSummary] ' + message); // log
    ui.alert(message);
  }
}

/////////////////////////////
// Configurations and Misc //
/////////////////////////////

/**
 * Gets the latest listed date string (yyyy-MM-dd format) in a specified column of a Spreadsheet sheet, in form of Date object.
 * Since yyyy-MM-dd format text will be automatically converted into a Date object in Google Spreadsheet,
 * this function will revert the data into string by .setNumberFormat('@')
 * @param {Object} sheet Sheet object of Google Spreadsheet. https://developers.google.com/apps-script/reference/spreadsheet/sheet
 * @param {number} columnNum Column number that the date is listed on.
 * @param {number} rowOffset [Optional] Number of rows to offset to get the body of the dates listed.
 * rowOffset defaults to 1; i.e., it is assumed, by default, that the dates start from the second row, where the first row is used as the header row. 
 * @return {Date} Returns null if no date was available
 */
function getLatestDate_(sheet, columnNum, rowOffset = 1) {
  if (sheet.getLastRow() <= rowOffset) {
    return null;
  } else {
    let dates = sheet.getRange(1 + rowOffset, columnNum, sheet.getLastRow() - rowOffset, 1).setNumberFormat('@').getValues();
    let latestDateObj = dates.reduce((curLatest, date) => {
      return (yMd2Date_(date[0]).getTime() >= yMd2Date_(curLatest[0]).getTime() || !curLatest ? date : curLatest);
    });
    return yMd2Date_(latestDateObj[0]);
  }
}

/**
 * Converts yyyy-MM-dd date string into date object based on the time zone of the script.
 * @param {string} yMd Date string in form of 'yyyy-MM-dd'
 * @returns {Date}
 */
function yMd2Date_(yMd) {
  let newDate = new Date(yMd.slice(0, 4), parseInt(yMd.slice(5, 7)) - 1, yMd.slice(-2));
  return newDate;
}

/**
 * Gets the latest listed year-month string (yyyy-MM format) in a specified column of a Spreadsheet sheet.
 * @param {Object} sheet Sheet object of Google Spreadsheet. https://developers.google.com/apps-script/reference/spreadsheet/sheet
 * @param {number} columnNum Column number that the date is listed on.
 * @param {number} rowOffset [Optional] Number of rows to offset to get the body of the dates listed.
 * rowOffset defaults to 1; i.e., it is assumed, by default, that the dates start from the second row, where the first row is used as the header row. 
 * @return {string} String in yyyy-MM format. Null if no date was available.
 */
function getLatestMonth_(sheet, columnNum, rowOffset = 1) {
  if (sheet.getLastRow() <= rowOffset) {
    return null;
  } else {
    let yearMonths = sheet.getRange(1 + rowOffset, columnNum, sheet.getLastRow() - rowOffset, 1).setNumberFormat('@').getValues();
    let latestYearMonth = yearMonths.reduce((curLatest, yearMonth) => {
      let curLatestDate = yearMonth2Date_(curLatest[0]);
      let yearMonthDate = yearMonth2Date_(yearMonth[0]);
      return (yearMonthDate.getTime() >= curLatestDate.getTime() || !curLatest ? yearMonth : curLatest);
    });
    return latestYearMonth[0];
  }
}

/**
 * Convert year-month string in form of 'yyyy-MM' into a date object specifying the first date of the designated year-month.
 * @param {string} yearMonth Year and month in 'yyyy-MM' format.
 * @returns {Date} Returns null value if yearMonth is null.
 */
function yearMonth2Date_(yearMonth) {
  var dateObj = (yearMonth ? new Date(yearMonth.slice(0, 4), parseInt(yearMonth.slice(-2)) - 1, 1) : null);
  return dateObj;
}

/**
 * Returns a formatted date for making reports query on YouTube Analytics API.
 * @param {Date} date Date object to format
 * @param {string} timeZone [Optional] Time zone; defaults to the script's time zone. https://developers.google.com/apps-script/reference/base/session#getScriptTimeZone().
 * https://developers.google.com/youtube/analytics/dimensions#Temporal_Dimensions
 * @returns {string} Formatted date string in yyyy-MM-dd
 */
function formattedDateAnalytics_(date, timeZone = Session.getScriptTimeZone()) {
  var dateString = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
  return dateString;
}

/**
 * Converts durations in ISO8601 format into seconds.
 * This function regards only the days, hours, minutes, and seconds to fit this script's needs; years, months, and weeks are discarded.
 * @param {string} iso8601duration Duration in ISO8601 format. https://en.wikipedia.org/wiki/ISO_8601#Durations
 * @returns {number} Duration in seconds.
 */
function iso8601duration2sec_(iso8601duration) {
  var SEC = {
    minute: 60, // minutes in seconds
    hour: 60 * 60, // hours in seconds
    day: 24 * 60 * 60, // days in seconds
    second: 1
  };
  var durationRegex = /^(?<sign>-)?P(((?<year>\d+)Y)?((?<month>\d+)M)?((?<day>\d+)D)?(?:T((?<hour>\d+)H)?((?<minute>\d+)M)?((?<second>\d+)S)?)?|(?<week>\d+)W)$/;
  var durationGroups = iso8601duration.match(durationRegex).groups;
  var duration = 0;
  for (let k in durationGroups) {
    if (!SEC[k]) {
      continue;
    }
    let durationInt = (durationGroups[k] ? parseInt(durationGroups[k]) : 0);
    duration += (durationInt * SEC[k]);
  }
  return duration;
}

/**
 * Create a Javascript object from a 2d array, grouped by a given property.
 * @param {array} data 2-dimensional array with a header as its first row.
 * @param {string} property [Optional] Name of field name in header to group by.
 * When property is not specified, this function will return an object with a key 'data', whose value is a simple array of objects converted from the given 2d array.
 * If the designated property is not included in the header, this function will return an empty object.
 * @return {Object}
 */
function groupArray_(data, property = null) {
  let header = data.shift();
  let index = header.indexOf(property);
  if (!property) {
    let groupedObj = {};
    groupedObj['data'] = data.map(values => header.reduce((obj, key, ind) => {
      obj[key] = values[ind];
      return obj;
    }, {}));
    return groupedObj;
  } else if (index < 0) {
    let invalidProperty = {};
    return invalidProperty;
  } else {
    let groupedObj = data.reduce((accObj, curArr) => {
      let key = curArr[index];
      if (!accObj[key]) {
        accObj[key] = [];
      }
      let rowObj = header.reduce((obj, headerKey, ind) => {
        obj[headerKey] = curArr[ind];
        return obj;
      }, {});
      accObj[key].push(rowObj);
      return accObj;
    }, {});
    return groupedObj;
  }
}

/**
 * Creates elements for pivot tables on the spreadsheet by determining whether the given date is before or after the baseline date
 * and returns an array whose first element is PREVIOUS or CURRENT for dates before or after the baseline date, respectively.
 * If this first element is 'CURRENT', the second element returned is the string of year-month in yyyy-MM of the given date;
 * if 'PREVIOUS', this will be the year-month of the year after the given date.
 * @param {Date} targetDate The target date object to determine.
 * @param {Date} baselineDate The baseline date object.
 * @param {string} timezone [Optional] Timezone to express the yyyy-MM in. Defaults to the executing script file's timezone.
 * @param {Date} baselineAnnualDate [Optional] The starting date to define 'PREVIOUS'.
 * If this value is, as by default, null, all targetDate dates before the baselineDate will be labeled 'PREVIOUS'.
 * If a valid date is defined, all targetDate dates before this value will be labeled 'NA'.
 * @returns {array} [status, yearMonth], where the status is 'PREVIOUS', 'CURRENT', or 'NA' and yearMonth is a string of 'yyyy-MM'.
 */
function dateStatus_(targetDate, baselineDate, timezone = Session.getScriptTimeZone(), baselineAnnualDate = null) {
  let status = (targetDate.getTime() >= baselineDate.getTime() ? 'CURRENT' : 'PREVIOUS');
  if (baselineAnnualDate && status == 'PREVIOUS') {
    status = (targetDate.getTime() < baselineAnnualDate.getTime() ? 'NA' : status);
  }
  // convert previous year's year-month into current year's corresponding months for aggregation
  let yearMonthAggPre = Utilities.formatDate(targetDate, timezone, 'yyyy-MM');
  let yearMonthAgg = (status == 'PREVIOUS' || status == 'NA' ? `${parseInt(yearMonthAggPre.slice(0, 4)) + 1}-${yearMonthAggPre.slice(-2)}` : yearMonthAggPre);
  return [status, yearMonthAgg];
}