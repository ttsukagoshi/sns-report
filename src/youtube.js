// Global variables are defined on index.js

///////////////////
// Authorize 認証//
///////////////////

/**
 * Direct the user to the authorization URL
 * https://github.com/gsuitedevs/apps-script-oauth2#2-direct-the-user-to-the-authorization-url
 */
function showSidebarYouTubeApi() {
  var youtubeAPIService = getYouTubeAPIService_();
  if (!youtubeAPIService.hasAccess()) {
    let authorizationUrl = youtubeAPIService.getAuthorizationUrl();
    let template = HtmlService.createTemplate('<a href="<?= authorizationUrl ?>" target="_blank">Authorize YouTube Data/Analytics API</a>.');
    template.authorizationUrl = authorizationUrl;
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    let template = HtmlService.createTemplate(
      '[YouTube Data/Analytics API] You are already authorized.');
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  }
}
/**
 * Logout
 * https://github.com/gsuitedevs/apps-script-oauth2#logout
 */
function logout() {
  var service = getYouTubeAPIService_()
  service.reset();
}

/**
 * Create the OAuth2 service
 * https://github.com/gsuitedevs/apps-script-oauth2#1-create-the-oauth2-service
 */
function getYouTubeAPIService_() {
  // Script Properties
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var [clientId, clientSecret] = [scriptProperties.clientId, scriptProperties.clientSecret];

  return OAuth2.createService('youtubeAPI')
    // Set the endpoint URLs, which are the same for all Google services.
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')

    // Set the client ID and secret, from the Google Developers Console.
    .setClientId(clientId)
    .setClientSecret(clientSecret)

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
  var isAuthorized = youtubeAPIService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('[YouTube Data/Analytics API] Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('[YouTube Data/Analytics API] Denied. You can close this tab');
  }
}

///////////////////////////////
// Update Channel/Video List //
///////////////////////////////

/**
 * Update both channel and video summary list.
 */
function updateAllYouTubeList() {
  updateYouTubeSummaryChannelList();
  updateYouTubeSummaryVideoList();
}

/**
 * List the authorized user's channel(s) on the summary spreadsheet
 */
function updateYouTubeSummaryChannelList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var config = getConfig_();
  var now = new Date();
  try {
    // Get list of channel(s)
    var channelListFull = youtubeMyChannelList_();
    // Extract data for copying into spreadsheet
    var channelList = channelListFull.map(
      function (element, index) {
        let num = index + 1;
        let thumbnailUrl = `=image("${element.snippet.thumbnails.default.url}")`; // For using the image function on spreadsheet
        let id = element.id;
        let title = element.snippet.title;
        let description = element.snippet.description;
        let publishedAt = formattedDate_(new Date(element.snippet.publishedAt), timeZone);
        let viewCount = element.statistics.viewCount;
        let subscriberCount = element.statistics.subscriberCount;
        let videoCount = element.statistics.videoCount;
        let timestamp = formattedDate_(now, timeZone);
        return [timestamp, num, thumbnailUrl, id, title, description, publishedAt, viewCount, subscriberCount, videoCount];
      }
    );
    // Set the text values into spreadsheets (summary and individual)
    //// Summary Spreadsheet
    ss.getSheetByName(config.SHEET_NAME_MY_CHANNELS)
      .getRange(2, 1, channelList.length, channelList[0].length) // Assuming that table body to which the list is copied starts from the 4th row of column 1 ('A' column).
      .setValues(channelList);
    //// Current Spreadsheet of this year
    var currentSheet = SpreadsheetApp.openById(scriptProperties.currentSpreadsheetId).getSheetByName(config.SHEET_NAME_MY_CHANNELS);
    currentSheet.getRange(currentSheet.getLastRow() + 1, 1, channelList.length, channelList[0].length) // Assuming that table body to which the list is copied starts from the 4th row of column 1 ('A' column).
      .setValues(channelList);
    // Log & Notify
    enterLog_(scriptProperties.currentSpreadsheetId, logSheetName, 'Success: updated channel list.', now)
    ui.alert('Completed', 'Updated summary channel list.', ui.ButtonSet.OK);
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(scriptProperties.currentSpreadsheetId, logSheetName, message, now)
    ui.alert('Error', message, ui.ButtonSet.OK);
  }
}

/**
 * List the authorized user's video(s) on the summary spreadsheet
 */
function updateYouTubeSummaryVideoList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var config = getConfig_();
  var now = new Date();
  try {
    // Get list of channel(s)
    let videoListFull = youtubeMyVideoList_();
    // Extract data for copying into spreadsheet
    let videoList = videoListFull.map(
      function (element, index) {
        let num = index + 1;
        let thumbnailUrl = element.snippet.thumbnails.high.url;
        let thumbnailUrlFunction = `=image("${thumbnailUrl}")`; // For using the image function on spreadsheet
        let channelId = element.snippet.channelId;
        let channelTitle = element.snippet.channelTitle;
        let videoId = element.id.videoId;
        let title = element.snippet.title;
        let description = element.snippet.description;
        let videoUrl = `https://www.youtube.com/watch?v=${videoId}`;
        let publishedAt = formattedDate_(new Date(element.snippet.publishedAt), timeZone);
        let timestamp = formattedDate_(now, timeZone);
        // Statistics and status of each video
        let videoStats = youtubeVideos_([videoId], false)[0];
        let duration = videoStats.contentDetails.duration;
        let caption = videoStats.contentDetails.caption
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
          privacyStatus,
          viewCount,
          likeCount,
          dislikeCount
        ];
        return videoListElement;
      }
    );
    // Set the text values into spreadsheets (summary and individual)
    //// Summary Spreadsheet
    ss.getSheetByName(config.SHEET_NAME_MY_VIDEOS)
      .getRange(2, 1, videoList.length, videoList[0].length) // Assuming that table body to which the list is copied starts from the 4th row of column 1 ('A' column).
      .setValues(videoList);
    //// Summary Spreadsheet
    var currentSheet = SpreadsheetApp.openById(scriptProperties.currentSpreadsheetId).getSheetByName(config.SHEET_NAME_MY_VIDEOS);
    currentSheet.getRange(currentSheet.getLastRow() + 1, 1, videoList.length, videoList[0].length) // Assuming that table body to which the list is copied starts from the 4th row of column 1 ('A' column).
      .setValues(videoList);
    // Log & Notify
    enterLog_(scriptProperties.currentSpreadsheetId, logSheetName, 'Success: updated video list.', now);
    ui.alert('Completed', 'Updated summary video list.', ui.ButtonSet.OK);
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(scriptProperties.currentSpreadsheetId, logSheetName, message, now);
    ui.alert('Error', message, ui.ButtonSet.OK);
  }
}

/**
 * Get the list of all channels that the authorized user owns.
 * Channel properties are described at https://developers.google.com/youtube/v3/docs/channels#resource
 * @returns {array} Array of Javascript objects of channel properties.
 */
function youtubeMyChannelList_() {
  // Get the list of channels
  var channelParameters = {
    part: 'snippet,statistics',
    mine: true
  };
  try {
    var channelList = JSON.parse(youtubeData_('channels', channelParameters));
    if (channelList.nextPageToken) {
      let nextChannelList = {};
      nextChannelList['nextPageToken'] = channelList.nextPageToken;
      while (nextChannelList.nextPageToken) {
        let nextChannelParameters = {};
        nextChannelParameters = {
          part: 'snippet, statistics',
          mine: true,
          pageToken: nextChannelList.nextPageToken
        };
        nextChannelList = {};
        nextChannelList = JSON.parse(youtubeData_('channels', nextChannelParameters));
        nextChannelList.items.forEach(value => channelList.items.push(value));
      }
    }
    return channelList.items;
  } catch (error) {
    throw new Error(errorMessage_(error));
  }
}

/**
 * Get the list of all video(s) that the authorized user owns.
 * Video properties are described at https://developers.google.com/youtube/v3/docs/videos#resource
 * @returns {array} Array of Javascript objects of video properties.
 */
function youtubeMyVideoList_() {
  var videoParameters = {
    part: 'snippet',
    forMine: true,
    type: 'video'
  };
  try {
    var videoList = JSON.parse(youtubeData_('search', videoParameters));
    if (videoList.nextPageToken) {
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
        nextVideoList = {};
        nextVideoList = JSON.parse(youtubeData_('search', nextParameters));
        nextVideoList.items.forEach(value => videoList.items.push(value));
      }
    }
    return videoList.items;
  } catch (error) {
    throw new Error(errorMessage_(error));
  }
}

/**
 * Get the details of specific YouTube channel(s).
 * Channel properties are described at https://developers.google.com/youtube/v3/docs/channels#resource
 * @param {array} channelIds Array of YouTube channel IDs to retrieve the details.
 * Subject to API quota; bulk request may result in HTTP error 400. See https://developers.google.com/youtube/v3/docs/channels/list#parameters
 * @param {boolean} getDetails [Optional] Gets the snippet of the YouTube channel(s) as well as its status and statistics. Defaults to false.
 * @returns {array} Array of Javascript objects of channel properties. 
 */
function youtubeChannels_(channelIds, getDetails = false) {
  var channelParameters = {
    part: (getDetails == true ? 'snippet,statistics,status,contentDetails' : 'statistics,status,contentDetails'),
    id: channelIds.join()
  };
  try {
    var channelDetails = JSON.parse(youtubeData_('channels', channelParameters));
    if (channelDetails.nextPageToken) {
      let nextChannelDetails = {};
      nextChannelDetails['nextPageToken'] = nextChannelDetails.nextPageToken;
      while (nextChannelDetails.nextPageToken) {
        let nextParameters = {};
        nextParameters = {
          part: (getDetails == true ? 'snippet,statistics,status,contentDetails' : 'statistics,status,contentDetails'),
          id: channelIds.join(),
          pageToken: nextChannelDetails.nextPageToken
        };
        nextChannelDetails = {};
        nextChannelDetails = JSON.parse(youtubeData_('channels', nextParameters));
        nextChannelDetails.items.forEach(value => channelDetails.items.push(value));
      }
    }
    return channelDetails.items;
  } catch (error) {
    throw new Error(errorMessage_(error));
  }
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
  var videoParameters = {
    part: (getDetails == true ? 'snippet,statistics,status,contentDetails' : 'statistics,status,contentDetails'),
    id: videoIds.join()
  };
  try {
    var videoDetails = JSON.parse(youtubeData_('videos', videoParameters));
    if (videoDetails.nextPageToken) {
      let nextVideoDetails = {};
      nextVideoDetails['nextPageToken'] = videoDetails.nextPageToken;
      while (nextVideoDetails.nextPageToken) {
        let nextParameters = {};
        nextParameters = {
          part: (getDetails == true ? 'snippet,statistics,status,contentDetails' : 'statistics,status,contentDetails'),
          id: videoIds.join(),
          pageToken: nextVideoList.nextPageToken
        };
        nextVideoDetails = {};
        nextVideoDetails = JSON.parse(youtubeData_('videos', nextParameters));
        nextVideoDetails.items.forEach(value => videoDetails.items.push(value));
      }
    }
    return videoDetails.items;
  } catch (error) {
    throw new Error(errorMessage_(error));
  }
}

/**
 * GET request using YouTube Data API.
 * https://developers.google.com/youtube/v3/docs
 * @param {string} resourceType Resource type to target the GET request. https://developers.google.com/youtube/v3/docs#resource-types
 * @param {Object} parameters Parameters for the request in form of a Javascript object.
 */
function youtubeData_(resourceType, parameters) {
  var youtubeAPIService = getYouTubeAPIService_();
  try {
    if (!youtubeAPIService.hasAccess()) {
      throw new Error('Unauthorized. Get authorized by Menu > YouTube > Authorize');
    }
    let baseUrl = `https://www.googleapis.com/youtube/v3/${resourceType}`;
    let paramString = '?';
    for (let k in parameters) {
      let param = `${k}=${encodeURIComponent(parameters[k])}`;
      if (paramString.slice(-1) !== '?') {
        param = '&' + param;
      }
      paramString += param;
    }
    let url = baseUrl + paramString;
    let response = UrlFetchApp.fetch(url, {
      method: 'GET',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + encodeURIComponent(youtubeAPIService.getAccessToken())
      }
    });
    return response;
  } catch (error) {
    throw error;
  }
}

///////////////
// Analytics //
///////////////

/**
 * Get latest analytics data for YouTube channel and videos that the authorized user owns.
 */
function updateYouTubeAnalyticsData() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var currentYear = scriptProperties.currentYear;
  var yearLimit = true;
  try {
    // Get latest data
    let updatedChannelAnalyticsDate = youtubeAnalyticsChannel(currentYear, yearLimit);
    let updateChannelDemographics = youtubeAnalyticsDemographics(currentYear, yearLimit);
    let updatedVideoAnalyticsDate = youtubeAnalyticsVideo(currentYear, yearLimit);
    // Determine change-of-the-year
    let newYear = (
      formattedDateAnalytics_(updatedChannelAnalyticsDate).slice(-5) == '12-31'
      && updateChannelDemographics.slice(-2) == '12'
      && formattedDateAnalytics_(updatedVideoAnalyticsDate).slice(-5) == '12-31'
    );
    console.log(newYear);///////////////////
    // newYear = trueの場合
    // 新しいspreadsheet作成、URLをspreadsheet listに記録
    // scriptPropertiesのcurrentYear、currentSpreadsheetId更新
  } catch (error) {
    ui.alert(errorMessage_(error));
  }

}

/**
 * Update day-by-day YouTube Channel Analytics for the target year
 * If no previous data is available, this function will retrieve channel analytics starting from Jan 1 of the target year.
 * @param {number} targetYear Target year in yyyy
 * @param {boolean} yearLimit When true, limit the latest data to obtain to the end of the targetYear, i.e., Dec 31. Defaults to true.
 * @returns {Date} Latest date object of the updated analytics data.
 */
function youtubeAnalyticsChannel(targetYear, yearLimit = true) {
  // Get target spreadsheet
  var config = getConfig_();
  var spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
  var targetSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == targetYear && value.PLATFORM == 'YouTube'))[0].URL;
  var targetSpreadsheet = SpreadsheetApp.openByUrl(targetSpreadsheetUrl);
  var targetSheet = targetSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_ANALYTICS);
  var now = new Date();
  try {
    // Check the date of the latest analytics data and define startDate for youtubeAnalyticsReportsQuery_()
    // If the value returned for getLatestDate_() is null, i.e., there are no previous dates recorded in targetSheet,
    // latestDate will be Dec 31 of the previous year of targetYear
    let latestDate = (getLatestDate_(targetSheet, 1) ? getLatestDate_(targetSheet, 1) : new Date(targetYear - 1, 11, 31)); // Assuming that the date is recorded on column A of the targetSheet.
    console.log(latestDate);//////////////////////////////////////////
    startDateObj = new Date(latestDate.setDate(latestDate.getDate() + 1));
    // Setting parameters for youtubeAnalyticsReportsQuery_()
    let startDate = formattedDateAnalytics_(startDateObj);
    let endDateObj = (yearLimit ? new Date(targetYear, 11, 31) : now);
    let endDate = formattedDateAnalytics_(endDateObj);
    let metrics = 'views,likes,dislikes,subscribersGained,subscribersLost,estimatedMinutesWatched,averageViewDuration,cardImpressions,cardClicks';
    let ids = 'channel==MINE';
    let options = {
      dimensions: 'day,channel'
    };
    // Get analytics data
    let reports = JSON.parse(youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids, options));
    // Copy on spreadsheet
    let data = reports.rows.slice();
    let updatedLatestDateObj = getLatestDate_(targetSheet, 1);
    if (data && data.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
      // Get latest updated date
      updatedLatestDateObj = getLatestDate_(targetSheet, 1);
      // Log
      enterLog_(targetSpreadsheet.getId(), logSheetName, `Success: updated YouTube channel analytics for ${startDate} to ${formattedDateAnalytics_(updatedLatestDateObj)}.`, now);
    } else {
      enterLog_(targetSpreadsheet.getId(), logSheetName, `Success: no updates for YouTube channel analytics.`, now);
    }
    return updatedLatestDateObj;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheet.getId(), logSheetName, message, now);
    throw new Error(message);
  }
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
  try {
    // Check the month of the latest analytics data and define startDate for youtubeAnalyticsReportsQuery_()
    // If the value returned for getLatestMonth_() is null, i.e., there are no previous month recorded in targetSheet,
    // latestMonth will be January of the targetYear ('yyyy-01').
    let latestMonth = (getLatestMonth_(targetSheet, 1) ? getLatestMonth_(targetSheet, 1) : targetYear + '-01'); // Assuming that the year-month (yyyy-MM) is recorded on column A of the targetSheet.
    let latestMonthDate = yearMonth2Date_(latestMonth); // latestMonth in Date object (the first day of that yyyy-MM)
    // Get existing data in form of a 2d-array to overwrite the latest data
    let existingData = targetSheet.getDataRange().getValues();
    existingData.shift(); // Assuming that the first row is a header row.
    let existingDataUpdate = existingData.filter(element => element[0] !== latestMonth);
    // Determine the final endDate as currentLatestMonthDate. If yearLimit is true, this will be no later than December of the targetYear.
    let currentLatestMonthDate = (
      yearLimit == true && parseInt(Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy')) > parseInt(targetYear)
        ? new Date(targetYear, 11, 1)
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
      endDate = formattedDateAnalytics_(new Date(latestMonthDate.setMonth(latestMonthDate.getMonth() + 1)));
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
      startDate = endDate.slice();
    }
    // Copy on spreadsheet
    targetSheet.getRange(2, 1, existingDataUpdate.length, existingDataUpdate[0].length).setValues(existingDataUpdate); // Assuming that the 1st row of the targetSheet is header row and that the actual data starts from the 2nd row
    // Log
    enterLog_(targetSpreadsheet.getId(), logSheetName, `Success: updated YouTube channel demographics for ${latestMonth} to ${thisYearMonth}.`, now);
    return thisYearMonth;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheet.getId(), logSheetName, message, now);
    throw new Error(error);
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
  try {
    // Check the date of the latest analytics data and define startDate for youtubeAnalyticsReportsQuery_()
    // If the value returned for getLatestDate_() is null, i.e., there are no previous dates recorded in targetSheet,
    // latestDate will be Dec 31 of the previous year of targetYear
    let latestDate = (getLatestDate_(targetSheet, 1) ? getLatestDate_(targetSheet, 1) : new Date(targetYear - 1, 11, 31)); // Assuming that the date is recorded on column A of the targetSheet.
    let startDateObj = new Date(latestDate.setDate(latestDate.getDate() + 1));

    // Setting parameters for youtubeAnalyticsReportsQuery_()
    let startDate = formattedDateAnalytics_(startDateObj);
    let endDateObj = (yearLimit ? new Date(targetYear, 11, 31) : now);
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
    let data = [];
    for (let i = 0; i < videoList.length; i++) {
      let video = videoList[i];
      let videoId = video.id.videoId;
      let channelId = video.snippet.channelId;
      options.filters = `video==${videoId}`;
      let rawAnalytics = JSON.parse(youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids, options)).rows;
      for (let j = 0; j < rawAnalytics.length; j++) {
        let dayVideo = rawAnalytics[j];
        // Assuming that the first column is the 'day' column, insert channel ID and video ID to the day-by-day data array
        dayVideo.splice(1, 0, channelId, videoId);
        data.push(dayVideo);
      }
    }
    // Copy on spreadsheet
    let updatedLatestDateObj = getLatestDate_(targetSheet, 1);
    if (data && data.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
      // Get latest updated date
      updatedLatestDateObj = getLatestDate_(targetSheet, 1);
      // Log
      enterLog_(targetSpreadsheet.getId(), logSheetName, `Success: updated YouTube video analytics for ${startDate} to ${formattedDateAnalytics_(updatedLatestDateObj)}.`, now);
    } else {
      enterLog_(targetSpreadsheet.getId(), logSheetName, `Success: no updates for YouTube video analytics.`, now);
    }
    return updatedLatestDateObj;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(targetSpreadsheet.getId(), logSheetName, message, now);
    throw new Error(message);
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
  try {
    let youtubeAPIService = getYouTubeAPIService_();
    if (!youtubeAPIService.hasAccess()) {
      throw new Error('Unauthorized. Get authorized by Menu > YouTube > Authorize');
    }
    let baseUrl = `https://youtubeanalytics.googleapis.com/v2/reports`;
    let paramString = '?';
    for (let k in parameters) {
      let param = `${k}=${encodeURIComponent(parameters[k])}`;
      if (paramString.slice(-1) !== '?') {
        param = '&' + param;
      }
      paramString += param;
    }
    let url = baseUrl + paramString;
    let response = UrlFetchApp.fetch(url, {
      method: 'GET',
      contentType: 'application/json; charset=UTF-8',
      headers: {
        Authorization: 'Bearer ' + encodeURIComponent(youtubeAPIService.getAccessToken())
      }
    });
    return response;
  } catch (error) {
    throw new Error(errorMessage_(error));
  }
}

//////////////////////////////
// Create Analytics Summary //
//////////////////////////////

/**
 * Update YouTube analytics summary sheet of the spreadsheet.
 * The range designated in this script assumes the cells to be positioned in the layout of the sample spreadsheet.
 */
function createYouTubeAnalyticsSummary() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfig_();
  var summarySheet = ss.getSheetByName(config.SHEET_NAME_YOUTUBE_SUMMARY);
  var now = new Date();
  try {
    // Prompt to enter channel ID and target period
    //// Channel ID
    let promptMessageChannelId = 'Enter ID of the YouTube channel to create report. Channel IDs can most simply be obtained by looking at its URL: https://www.youtube.com/channel/**********';
    let responseChannelId = ui.prompt(promptMessageChannelId, ui.ButtonSet.OK_CANCEL);
    if (responseChannelId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let targetChannelId = responseChannelId.getResponseText();
    ////// Get list of YouTube channel(s) that the authorized user owns; for testing.
    let myChannelListFull = youtubeMyChannelList_();
    let myChannelIds = myChannelListFull.map(element => element.id);
    if (!targetChannelId) {
      throw new Error('No text entered for channel ID.');
    } else if (!myChannelIds.includes(targetChannelId)) {
      throw new Error(`Invalid Channel ID: "${targetChannelId}".\nMake sure to enter the ID of the YouTube channel that you own.`)
    }
    let targetChannelName = myChannelListFull.filter(element => element.id == targetChannelId)[0].snippet.title;
    summarySheet.getRange(4, 2).setValue(`${targetChannelName} (${targetChannelId})`);
    //// Report month
    let promptMessageReportMonth = 'Enter the month to create report for report in form of "yyyy-MM", e.g., enter "2020-03" for getting report for March 2020.';
    let responseReportMonth = ui.prompt(promptMessageReportMonth, ui.ButtonSet.OK_CANCEL);
    if (responseReportMonth.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let reportMonth = responseReportMonth.getResponseText();
    if (!reportMonth) {
      throw new Error('No text entered for report month.');
    } else if (reportMonth.length !== 7 || !reportMonth.match(/\d{4}-\d{2}/g)) {
      throw new Error(`Invalid period: "${reportMonth}".\nMake sure the report month is expressed in "yyyy-MM", e.g., enter "2020-03" for getting report for March 2020.`);
    }
    let targetPeriodStartPre = new Date(parseInt(reportMonth.slice(0, 4) - 2), parseInt(reportMonth.slice(-2)) - 1, 1);
    console.log(['targetPeriodStartPre', targetPeriodStartPre]);//////////////////////////
    let targetPeriodStart = new Date(targetPeriodStartPre.setMonth(targetPeriodStartPre.getMonth() + 1));
    console.log(['targetPeriodStart', targetPeriodStart]);//////////////////////////
    let reportPeriodStart = new Date(targetPeriodStart.setFullYear(targetPeriodStart.getFullYear() + 1));
    console.log(['reportPeriodStart', reportPeriodStart]);//////////////////////////
    let targetPeriodEndPre = new Date(reportMonth.slice(0, 4), parseInt(reportMonth.slice(-2) - 1), 1);
    targetPeriodEndPre.setMonth(targetPeriodEndPre.getMonth() + 1);
    let targetPeriodEnd = new Date(targetPeriodEndPre.setDate(targetPeriodEndPre.getDate() - 1));
    console.log(['targetPeriodEnd', targetPeriodEnd]);//////////////////////////
    summarySheet.getRange(6, 2).setValue(`${formattedDateAnalytics_(reportPeriodStart)} - ${formattedDateAnalytics_(targetPeriodEnd)}`)
    //// Report as of
    summarySheet.getRange(8, 2).setValue(`${formattedDate_(now)}`);

    // Update Channel/Video list
    // updateAllYouTubeList(); /////////////////////////////////// un-comment out when all is complete

    // Get background data from the spreadsheet(s) of the target period
    let bgDataChannelAnalytics = [];
    let bgDataChannelDemographics = [];
    let bgDataVideoAnalytics = [];
    let spreadsheetList = getSpreadsheetList_(config.SHEET_NAME_SPREADSHEET_LIST);
    let dataYear = targetPeriodStart.getFullYear();
    while (dataYear <= targetPeriodEnd.getFullYear()) {
      let dataSpreadsheetUrl = spreadsheetList.filter(value => (value.YEAR == dataYear && value.PLATFORM == 'YouTube'))[0].URL;
      let dataSpreadsheet = SpreadsheetApp.openByUrl(dataSpreadsheetUrl);
      // Get channel analytics
      let channelAnalyticsSheet = dataSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_ANALYTICS);
      let channelAnalyticsDataFull = channelAnalyticsSheet.getDataRange().getValues();
      let channelAnalyticsData = [];
      channelAnalyticsDataFull.shift();
      // Get channel demographics
      let channelDemographicsSheet = dataSpreadsheet.getSheetByName(config.SHEET_NAME_CHANNEL_DEMOGRAPHICS);
      let channelDemographicsDataFull = channelDemographicsSheet.getDataRange().getValues();
      let channelDemographicsData = [];
      channelDemographicsDataFull.shift();
      // Get video analytics
      let videoAnalyticsSheet = dataSpreadsheet.getSheetByName(config.SHEET_NAME_VIDEO_ANALYTICS);
      let videoAnalyticsDataFull = videoAnalyticsSheet.getDataRange().getValues();
      let videoAnalyticsData = [];
      videoAnalyticsDataFull.shift();
      if (dataYear == targetPeriodStart.getFullYear()) {
        channelAnalyticsData = channelAnalyticsDataFull.filter(element => {
          console.log(element);////////////////////////
          // Assuming that the first column of the channel analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          return thisDate.getTime() >= targetPeriodStart.getTime();
        });
        channelDemographicsData = channelDemographicsDataFull.filter(element => {
          // Assuming that the first column of the channel demographics data table is the year-month in yyyy-MM
          let thisDate = yearMonth2Date_(element[0]);
          return thisDate.getMonth() >= targetPeriodStart.getMonth();
        });
        videoAnalyticsData = videoAnalyticsDataFull.filter(element => {
          // Assuming that the first column of the video analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          return thisDate.getTime() >= targetPeriodStart.getTime();
        });
      } else if (dataYear == targetPeriodEnd.getFullYear()) {
        channelAnalyticsData = channelAnalyticsDataFull.filter(element => {
          // Assuming that the first column of the channel analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          return thisDate.getTime() <= targetPeriodEnd.getTime();
        });
        channelDemographicsData = channelDemographicsDataFull.filter(element => {
          // Assuming that the first column of the channel demographics data table is the year-month in yyyy-MM
          let thisDate = yearMonth2Date_(element[0]);
          return thisDate.getMonth() <= targetPeriodEnd.getMonth();
        });
        videoAnalyticsData = videoAnalyticsDataFull.filter(element => {
          // Assuming that the first column of the video analytics data table is the date
          let thisDate = new Date(element[0].slice(0, 4), parseInt(element[0].slice(5, 7)) - 1, element[0].slice(-2));
          return thisDate.getTime() <= targetPeriodEnd.getTime();
        });
      } else {
        channelAnalyticsData = channelAnalyticsDataFull.slice();
        channelDemographicsData = channelDemographicsDataFull.slice();
        videoAnalyticsData = videoAnalyticsDataFull.slice();
      }
      let newBgDataChannelAnalytics = bgDataChannelAnalytics.concat(channelAnalyticsData);
      bgDataChannelAnalytics = newBgDataChannelAnalytics.slice();
      let newBgDataChannelDemographics = bgDataChannelDemographics.concat(channelDemographicsData);
      bgDataChannelDemographics = newBgDataChannelDemographics.slice();
      let newBgDataVideoAnalytics = bgDataVideoAnalytics.concat(videoAnalyticsData);
      bgDataVideoAnalytics = newBgDataVideoAnalytics.slice();
      // Update variables for next loop
      dataYear += 1;
    }
    // Copy on spreadsheet
    ss.getSheetByName(config.SHEET_NAME_CHANNEL_ANALYTICS_TEMP)
      .getRange(1, 1, bgDataChannelAnalytics.length, bgDataChannelAnalytics[0].length)
      .setValues(bgDataChannelAnalytics);
    ss.getSheetByName(config.SHEET_NAME_CHANNEL_DEMOGRAPHICS_TEMP)
      .getRange(1, 1, bgDataChannelDemographics.length, bgDataChannelDemographics[0].length)
      .setValues(bgDataChannelDemographics);
    ss.getSheetByName(config.SHEET_NAME_VIDEO_ANALYTICS_TEMP)
      .getRange(1, 1, bgDataVideoAnalytics.length, bgDataVideoAnalytics[0].length)
      .setValues(bgDataVideoAnalytics);

    //////////////////////////////
  } catch (error) {
    ui.alert(errorMessage_(error));
  }
}


/////////////////////////////
// Configurations and Misc //
/////////////////////////////


/**
 * Returns an object of configurations from spreadsheet.
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
 * Gets the latest listed date string (yyyy-MM-dd format) in a specified column of a Spreadsheet sheet, in form of Date object.
 * Since yyyy-MM-dd format text will be automatically converted into a Date object in Google Spreadsheet,
 * this function assumes that the cell values obtained by getValues() function is an array of date objects.
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
    let latestDateObj = dates.reduce(function (curLatest, date) {
      return (yMd2Date_(date[0]).getTime() >= yMd2Date_(curLatest[0]).getTime() || !curLatest ? date : curLatest);
    });
    return yMd2Date_(latestDateObj[0]);
  }
}

/**
 * Gets the latest listed year-month string (yyyy-MM format) in a specified column of a Spreadsheet sheet.
 * @param {Object} sheet Sheet object of Google Spreadsheet. https://developers.google.com/apps-script/reference/spreadsheet/sheet
 * @param {number} columnNum Column number that the date is listed on.
 * @param {number} rowOffset [Optional] Number of rows to offset to get the body of the dates listed.
 * rowOffset defaults to 1; i.e., it is assumed, by default, that the dates start from the second row, where the first row is used as the header row. 
 * @return {string} Returns null if no date was available
 */
function getLatestMonth_(sheet, columnNum, rowOffset = 1) {
  if (sheet.getLastRow() <= rowOffset) {
    return null;
  } else {
    let yearMonths = sheet.getRange(1 + rowOffset, columnNum, sheet.getLastRow() - rowOffset, 1).setNumberFormat('@').getValues();
    let latestYearMonth = yearMonths.reduce(function (curLatest, yearMonth) {
      let curLatestDate = yearMonth2Date_(curLatest[0]);
      let yearMonthDate = yearMonth2Date_(yearMonth[0]);
      return (yearMonthDate.getTime() >= curLatestDate.getTime() || !curLatest ? yearMonth : curLatest);
    });
    return latestYearMonth[0];
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
 * Convert year-month string in form of 'yyyy-MM' into a date object specifying the first date of the designated year-month.
 * @param {string} yearMonth Year and month in 'yyyy-MM' format.
 * @returns {Date} Returns null value if yearMonth is null.
 */
function yearMonth2Date_(yearMonth) {
  var dateObj = (yearMonth ? new Date(yearMonth.slice(0, 4), parseInt(yearMonth.slice(-2)) - 1, 1) : null);
  return dateObj;
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
 * Returns a formatted date for making reports query on YouTube Analytics API.
 * @param {Date} date Date object to format
 * @param {string} timeZone [Optional] Time zone; defaults to the script's time zone. https://developers.google.com/apps-script/reference/base/session#getScriptTimeZone().
 * https://developers.google.com/youtube/analytics/dimensions#Temporal_Dimensions
 * @returns {string} Formatted date string
 */
function formattedDateAnalytics_(date, timeZone = Session.getScriptTimeZone()) {
  var dateString = Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
  return dateString;
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

/**
 * Set a column of numbers or dates in the designated Google Spreadsheet sheet into text to fix its format.
 * @param {Object} sheet Sheet object of Google Spreadsheet.
 * @param {number} colNum Target column number to set as text. Defaults to 1, i.e., column "A".
 * @returns {Object} The entered sheet object; for chaining.
 */
function setColumnAsText_(sheet, colNum = 1) {
  sheet.getRange(1, colNum, sheet.getLastRow(), 1).setNumberFormat('@');
  return sheet;
}