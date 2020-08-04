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

// Global variables are defined on index.js

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
    var [fbClientId, fbClientSecret] = [scriptProperties.fbClientId, scriptProperties.fbClientSecret];
  
    return OAuth2.createService('facebookAPI')
      // Set the endpoint URLs, which are the same for all Google services.
      .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
      .setTokenUrl('https://accounts.google.com/o/oauth2/token')
  
      // Set the client ID and secret, from the Google Developers Console.
      .setClientId(fbClientId)
      .setClientSecret(fbClientSecret)
  
      // Set the name of the callback function in the script referenced
      // above that should be invoked to complete the OAuth flow.
      .setCallbackFunction('authCallbackFacebookAPI_')
  
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
  