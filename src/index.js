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

var logSheetName = '99_Log';

/**
 * onOpen()
 * Add menu to spreadsheet
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('Facebook')
    .addItem('Authorize', 'showSidebarFacebookApi')
    .addItem('Logout/Reset', 'logoutFacebook')
    .addSeparator()
    .addToUi();
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
  ui.createMenu('Settings')
    .addItem('Setup', 'initialSettings')
    .addItem('Check Settings', 'checkSettings')
    .addToUi();
  ui.createMenu('test')/////////////////////////////////////////////////////
    .addItem('test', 'test')
    .addItem('catchup', 'archiveCatchup')
    .addToUi()
}

/**
 * Standarized error message
 * @param {Object} e Error object
 * @return {string} Standarized error message
 */
function errorMessage_(e) {
  let message = `Error: line - ${e.lineNumber}\n${e.stack}`
  return message;
}