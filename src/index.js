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
  ui.createMenu('YouTube')
    .addItem('Authorize', 'showSidebarYouTubeApi')
    .addItem('Logout', 'logout')
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