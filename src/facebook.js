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

/* exported authCallbackFacebookAPI_, logoutFacebook, showSidebarFacebookApi */
/* global checkYear_, enterLog_, errorMessage_, formattedDate_, getConfig_, LocalizedMessage, LOG_SHEET_NAME, OAuth2, spreadsheetUrl_, updateAllFbList, updateFbPagePostList */

// Facebook Graph API version to be used
const FB_API_VERSION = 'v8.0';
// Spreadsheet ID to use as template for creating a new spreadsheet to enter YouTube analytics data
const FB_NEW_SPREADSHEET_ID = '';///////////////////////////////////
// Other global variables are defined on index.js

///////////////////
// Authorize 認証//
///////////////////

/**
 * Direct the user to the authorization URL
 * https://github.com/gsuitedevs/apps-script-oauth2#2-direct-the-user-to-the-authorization-url
 */
function showSidebarFacebookApi() {
  var facebookAPIService = getFacebookAPIService_();
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  if (!facebookAPIService.hasAccess()) {
    let authorizationUrl = facebookAPIService.getAuthorizationUrl();
    let template = HtmlService.createTemplate(`<a href="<?= authorizationUrl ?>" target="_blank">${localizedMessages.messageList.facebook.authorizeFacebookAPI}</a>.`);
    template.authorizationUrl = authorizationUrl;
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    let template = HtmlService.createTemplate(
      localizedMessages.messageList.facebook.alreadyAuthorized
    );
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  }
}
/**
 * Logout
 * https://github.com/gsuitedevs/apps-script-oauth2#logout
 */
function logoutFacebook() {
  var service = getFacebookAPIService_();
  service.reset();
}

/**
 * Create the OAuth2 service
 * https://github.com/gsuitedevs/apps-script-oauth2#1-create-the-oauth2-service
 */
function getFacebookAPIService_() {
  // Script Properties
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var [fbAppId, fbAppSecret] = [scriptProperties.fbAppId, scriptProperties.fbAppSecret];

  return OAuth2.createService('facebookAPI')
    // Set the endpoint URLs
    .setAuthorizationBaseUrl(`https://www.facebook.com/${FB_API_VERSION}/dialog/oauth`)
    .setTokenUrl(`https://graph.facebook.com/${FB_API_VERSION}/oauth/access_token`)

    // Set the Facebook App ID and App secret, from the Facebook Developers App Dashboard
    .setClientId(fbAppId)
    .setClientSecret(fbAppSecret)

    // Set the name of the callback function in the script referenced
    // above that should be invoked to complete the OAuth flow.
    .setCallbackFunction('authCallbackFacebookAPI_')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())

    // Set the scopes to request (space-separated for Google services).
    .setScope('pages_read_engagement pages_show_list pages_read_user_content read_insights business_management');

  // Below are Google-specific OAuth2 parameters.
  /*
  // Sets the login hint, which will prevent the account chooser screen
  // from being shown to users logged in with multiple accounts.
  .setParam('login_hint', Session.getEffectiveUser().getEmail())
  
  // Requests offline access.
  .setParam('access_type', 'offline')
 
  // Consent prompt is required to ensure a refresh token is always
  // returned when requesting offline access.
  .setParam('prompt', 'consent');
  */
}

/**
 * Handle the callback
 * https://github.com/gsuitedevs/apps-script-oauth2#3-handle-the-callback
 * @param {Object} request 
 */
function authCallbackFacebookAPI_(request) {
  var facebookAPIService = getFacebookAPIService_();
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var isAuthorized = facebookAPIService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput(localizedMessages.messageList.facebook.authorizationSuccessful);
  } else {
    return HtmlService.createHtmlOutput(localizedMessages.messageList.facebook.authorizationDenied);
  }
}

////////////////////
// Page Analytics //
////////////////////

/**
 * GET request using Facebook Graph API.
 * https://developers.facebook.com/docs/graph-api/using-graph-api
 * @param {string} node Individual objects in the Facebook Graph API, e.g., User, Photo, Page, or Comment
 * @param {string} edge [Optional] Connections between a collection of objects and a single object,
 * such as Photos (edge) on a Page (node) or Comments (edge) on a Photo (node).
 * @param {array} fields [Optional] Data about an object, such as the name and id of a Page
 */
function getFbGraphData(node, edge = '', fields = []) {
  var facebookAPIService = getFacebookAPIService_();
  var localizedMessages = new LocalizedMessage(SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  if (!facebookAPIService.hasAccess()) {
    throw new Error(localizedMessages.messageList.facebook.errorUnauthorized);
  }
  let baseUrl = `https://graph.facebook.com/${FB_API_VERSION}/${node}`;
  if (edge) {
    baseUrl += `/${edge}`;
  }
  let fieldsUrl = (fields.length ? `&fields=${encodeURIComponent(fields.join())}` : '');
  let url = `${baseUrl}?access_token=${facebookAPIService.getAccessToken()}${fieldsUrl}`;
  let response = UrlFetchApp.fetch(url, {
    method: 'GET',
    contentType: 'application/json; charset=UTF-8'
  });
  return response;
}

/**
 * Retrieve basic information and page access tokens for the page that the authorized user has access to.
 * @returns {array} Array of objects containing page data
 */
function getFbPages_() {
  var node = 'me';
  var edge = 'accounts';
  var fields = [
    'name',
    'link',
    'about',
    'description',
    'category',
    'category_list',
    'emails',
    'website',
    'cover',
    'checkins',
    'country_page_likes',
    'fan_count',
    'access_token',
    'tasks'
  ];
  let pages = JSON.parse(getFbGraphData(node, edge, fields));
  let pagesData = pages.data.slice();
  let nextUrl = pages.paging.next || null;
  while (nextUrl) {
    let nextResponse = JSON.parse(UrlFetchApp.fetch(nextUrl, {
      method: 'GET',
      contentType: 'application/json; charset=UTF-8'
    }));
    pagesData = pagesData.concat(nextResponse.data);
    nextUrl = nextResponse.paging.next || null;
  }
  return pagesData;
}

/**
 * List the authorized user's page(s) on the summary spreadsheet
 * @param {boolean} muteUiAlert [Optional] Mute ui.alert() when true; defaults to false.
 * @returns {array} 2d array containing the page data
 */
function updateFbSummaryPageList(muteUiAlert = false) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var config = getConfig_();
  var fbCurrentYear = parseInt(scriptProperties.fbCurrentYear);
  var now = new Date();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  // Getting the target spreadsheet
  var spreadsheetListName = config.SHEET_NAME_SPREADSHEET_LIST;
  var options = {
    createNewFile: true,
    driveFolderId: scriptProperties.driveFolderId,
    templateFileId: FB_NEW_SPREADSHEET_ID,
    newFileName: 'Facebook Insights',
    newFileNamePrefix: config.SPREADSHEET_NAME_PREFIX
  };
  var spreadsheetUrl = spreadsheetUrl_(spreadsheetListName, fbCurrentYear, 'Facebook', options).url;
  try {
    // Get list of page(s)
    let pageListFull = getFbPages_();
    // Extract data for copying into spreadsheet
    let pageList = pageListFull.map((element, index) => {
      let num = index + 1;
      let coverUrl = element.cover.source;
      let coverUrlFunction = `=image("${coverUrl}")`; // For using the image function on spreadsheet
      let id = element.id;
      let pageUrl = element.link;
      let name = element.name;
      let about = element.about;
      let description = element.description;
      let category = element.category;
      let category_all = element.category_list.reduce((acc, val) => {
        acc.push(val.name);
        return acc;
      }, []).join();
      let emails = element.emails.join();
      let website = element.website;
      let checkins = element.checkins;
      let country_page_likes = element.country_page_likes;
      let fan_count = element.fan_count;
      let timestamp = formattedDate_(now, timeZone);
      return [
        timestamp,
        num,
        coverUrlFunction,
        coverUrl,
        id,
        pageUrl,
        name,
        about,
        description,
        category,
        category_all,
        emails,
        website,
        checkins,
        country_page_likes,
        fan_count
      ];
    }
    );
    // Set the text values into spreadsheets (summary and individual)
    //// Renew channel list of summary spreadsheet
    let myPagesSheet = ss.getSheetByName(config.SHEET_NAME_MY_PAGES);
    if (myPagesSheet.getLastRow() > 1) {
      myPagesSheet.getRange(2, 1, myPagesSheet.getLastRow() - 1, myPagesSheet.getLastColumn())
        .deleteCells(SpreadsheetApp.Dimension.ROWS);
    }
    myPagesSheet.getRange(2, 1, pageList.length, pageList[0].length) // Assuming that table body to which the list is copied starts from the 2nd row of column 1 ('A' column).
      .setValues(pageList);
    //// Add row(s) to current spreadsheet of this year
    let currentSheet = SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName(config.SHEET_NAME_MY_PAGES);
    currentSheet.getRange(currentSheet.getLastRow() + 1, 1, pageList.length, pageList[0].length) // Assuming that table body to which the list is copied starts from column 1 ('A' column).
      .setValues(pageList);
    // Log & Notify
    enterLog_(SpreadsheetApp.openByUrl(spreadsheetUrl).getId(), LOG_SHEET_NAME, localizedMessages.messageList.facebook.updatedPageListLog, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.misc.completedTitle, localizedMessages.messageList.facebook.updatedPageListAlert, ui.ButtonSet.OK);
    }
    return pageList;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(SpreadsheetApp.openByUrl(spreadsheetUrl).getId(), LOG_SHEET_NAME, message, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.error.errorTitle, message, ui.ButtonSet.OK);
    }
    return null;
  }
}

/**
 * List the authorized user's page(s) posts on the summary & individual year's spreadsheet
 * @param {boolean} muteUiAlert [Optional] Mute ui.alert() when true; defaults to false.
 * @returns {array} 2d array containing the page post data
 */

// posts
// docs/graph-api/reference/v8.0/page/feed
// me/feed?fields=id,permalink_url,created_time,backdated_time,place,picture,message,attachments,properties
function updateFbPagePostList(muteUiAlert = false) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // var timeZone = ss.getSpreadsheetTimeZone();
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var config = getConfig_();
  var fbCurrentYear = parseInt(scriptProperties.fbCurrentYear);
  var now = new Date();
  var localizedMessages = new LocalizedMessage(ss.getSpreadsheetLocale());
  // Getting the target spreadsheet
  var spreadsheetListName = config.SHEET_NAME_SPREADSHEET_LIST;
  var options = {
    createNewFile: true,
    driveFolderId: scriptProperties.driveFolderId,
    templateFileId: FB_NEW_SPREADSHEET_ID,
    newFileName: 'Facebook Insights',
    newFileNamePrefix: config.SPREADSHEET_NAME_PREFIX
  };
  var spreadsheetUrl = spreadsheetUrl_(spreadsheetListName, fbCurrentYear, 'Facebook', options).url;
  // Setting the parameters for getting page posts
  var postEdge = 'feed';
  var postFields = [
    'id',
    'permalink_url',
    'created_time',
    'backdated_time',
    'place',
    'picture',
    'message',
    'attachments',
    'properties'
  ];
  try {
    // Get list of page(s)
    let pageListFull = getFbPages_();
    // Extract data for copying into spreadsheet
    let postList = [];
    for (let i = 0; i < pageListFull.length; i++) {
      let pageId = pageListFull[i].id;
      let posts = JSON.parse(getFbGraphData(pageId, postEdge, postFields));
      let postsData = posts.data.slice();
      let nextUrl = posts.paging.next || null;
      while (nextUrl) {
        let nextResponse = JSON.parse(UrlFetchApp.fetch(nextUrl, {
          method: 'GET',
          contentType: 'application/json; charset=UTF-8'
        }));
        postsData = postsData.concat(nextResponse.data);
        nextUrl = nextResponse.paging.next || null;
      }
      postList = postList.concat(postsData);
    }
    let postListSS = postList.reduce((accList, post) => {
      if (checkYear_(post.created_time, fbCurrentYear, 'PST')) { // See index.js for definition of checkYear_()
        // Create post of the current year based on Pacific Time (daylight saving time taken into account)
        let postId = post.id;
        let postPermalink_url = post.permalink_url;
        let postCreatedTime = post.created_time;
        let postBackdatedTime = post.backdated_time || 'NA';
        let postPlace = post.place || { 'name': 'NA', 'id': 'NA' };
        let postPlaceName = postPlace.name;
        let postPlaceId = postPlace.id;
        let postPictureUrl = post.picture || 'NA';
        let postPictureImage = (post.picture ? `=image("${post.picture}")` : 'NA');
        let postMessage = post.message;
        let postData = [
          postId,
          postPermalink_url,
          postCreatedTime,
          postBackdatedTime,
          postPlaceName,
          postPlaceId,
          postPictureUrl,
          postPictureImage,
          postMessage,
        ];
        //////////////////////////////////////////////////////
        accList.push(postData);
      }
      return accList;
    }, []);
    /////////////////////////////////////////////////////////////
    // Set the text values into spreadsheet
    let currentSheet = SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName(config.SHEET_NAME_PAGE_POSTS);
    currentSheet.getRange(currentSheet.getLastRow() + 1, 1, postListSS.length, postListSS[0].length) // Assuming that table body to which the list is copied starts from column 1 ('A' column).
      .setValues(postListSS);
    // Log & Notify
    enterLog_(SpreadsheetApp.openByUrl(spreadsheetUrl).getId(), LOG_SHEET_NAME, localizedMessages.messageList.facebook.updatedPagePostListLog, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.misc.completedTitle, localizedMessages.messageList.facebook.updatedPagePostListAlert, ui.ButtonSet.OK);
    }
    return postListSS;
  } catch (error) {
    let message = errorMessage_(error);
    enterLog_(SpreadsheetApp.openByUrl(spreadsheetUrl).getId(), LOG_SHEET_NAME, message, now);
    if (!muteUiAlert) {
      ui.alert(localizedMessages.messageList.general.error.errorTitle, message, ui.ButtonSet.OK);
    }
    return null;
  }
}

/**
 * Update all summary list for Facebook.
 */
function updateAllFbList() {
  updateFbSummaryPageList();
  ///////////////////////////////////////////
}
