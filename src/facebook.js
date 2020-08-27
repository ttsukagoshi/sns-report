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

var FB_API_VERSION = 'v8.0'; // Facebook Graph API version to be used
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
  if (!facebookAPIService.hasAccess()) {
    let authorizationUrl = facebookAPIService.getAuthorizationUrl();
    let template = HtmlService.createTemplate('<a href="<?= authorizationUrl ?>" target="_blank">Authorize Facebook API</a>.');
    template.authorizationUrl = authorizationUrl;
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  } else {
    let template = HtmlService.createTemplate(
      '[Facebook API] You are already authorized.');
    let page = template.evaluate();
    SpreadsheetApp.getUi().showSidebar(page);
  }
}
/**
 * Logout
 * https://github.com/gsuitedevs/apps-script-oauth2#logout
 */
function logoutFacebook() {
  var service = getFacebookAPIService_()
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
    .setScope('pages_read_engagement pages_show_list pages_read_user_content read_insights business_management')

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
  var isAuthorized = facebookAPIService.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('[Facebook API] Success! You can close this tab.');
  } else {
    return HtmlService.createHtmlOutput('[Facebook API] Denied. You can close this tab');
  }
}

////////////////////
// Page Analytics //
////////////////////

function getPageAccessTokens_() {
  try {
    let userId = JSON.parse(getFbGraphData('me')).id;
    let node = `${userId}/accounts`;
    let pages = JSON.parse(getFbGraphData(node));
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
  } catch (error) {
    throw error;
  }
}

/**
 * GET request using Facebook Graph API.
 * https://developers.facebook.com/docs/graph-api/using-graph-api
 * @param {string} node Individual objects in the Facebook Graph API, e.g., User, Photo, Page, or Comment
 * @param {Object} edges [Optional] Parameters for the request in form of a Javascript object.//////////////////////////
 */
function getFbGraphData(node, edges = {}) {
  var facebookAPIService = getFacebookAPIService_();
  try {
    if (!facebookAPIService.hasAccess()) {
      throw new Error('Unauthorized. Get authorized by Menu > Facebook > Authorize');
    }
    let baseUrl = `https://graph.facebook.com/${FB_API_VERSION}/${node}`;
    let paramString = '';
    for (let k in edges) {
      let param = `${k}=${encodeURIComponent(parameters[k])}`;
      if (paramString.slice(-1) !== '?') {
        param = '&' + param;
      }
      paramString += param;
    }
    let url = `${baseUrl}?access_token=${facebookAPIService.getAccessToken()}${paramString}`;
    let response = UrlFetchApp.fetch(url, {
      method: 'GET',
      contentType: 'application/json; charset=UTF-8'
    });
    return response;
  } catch (error) {
    throw error;
  }
}