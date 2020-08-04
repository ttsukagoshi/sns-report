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

/**
 * Check current date; 
 * If current year is not equal with the script property 'currentYear', a new spreadsheet will be created under a designated Google Drive folder.
 * The spreadsheet ID of this new spreadsheet will replace script property 'currentSpreadsheetId'
 */
function dailyCheck() {
  var myEmail = Session.getActiveUser().getEmail();
  var userName = myEmail.substring(0, myEmail.indexOf('@')); // the *** in ***@myDomain.com
  var config = getConfig_();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var year = Utilities.formatDate(new Date(), timeZone, 'yyyy');
  try {
    if (year !== scriptProperties.currentYear || !scriptProperties.currentSpreadsheetId) {
      let driveFolder = DriveApp.getFolderById(scriptProperties.driveFolderId);
      let spreadsheetName = `YouTube Analytics ${year} (${userName})`;

      // Create spreadsheet in designated Google Drive folder; see below for details of createSpreadsheet() function.
      let createdSpreadsheetId = createSpreadsheet_(driveFolder, spreadsheetName);
      let createdSpreadsheet = SpreadsheetApp.openById(createdSpreadsheetId);
      let createdSpreadsheetUrl = createdSpreadsheet.getUrl();

      // Format the created spreadsheet
      // Set sheet name & create a new sheet
      let recordSheet = createdSpreadsheet.getSheets()[0].setName('Toggl_Record'); ////////////////////////取得データの詳細がわかってから、以下設定///////////
      var logSheet = createdSpreadsheet.insertSheet(1).setName('Log');
      var sheets = [recordSheet, logSheet];
      // Create a header row in the sheet
      var header = [];
      var headerItems = [];
      // for recordSheet
      headerItems[0] = ['TIME_ENTRY_ID',
        'WORKSPACE_ID',
        'WORKSPACE',
        'PROJECT_ID',
        'PROJECT',
        'DESCRIPTION',
        'TAGS',
        'START',
        'STOP',
        'DURATION_SEC',
        'USER_ID',
        'GUID',
        'BILLABLE',
        'DURONLY',
        'LAST_MODIFIED',
        'iCalID',
        'TIMESTAMP',
        'CALENDAR_ID',
        'updateFlag'
      ];
      // for logSheet
      headerItems[1] = ['TIMESTAMP', 'USERNAME', 'LOG'];
      // Define header style
      var headerStyle = SpreadsheetApp.newTextStyle().setBold(true).build();

      for (var i = 0; i < sheets.length; i++) {
        var sheet = sheets[i];
        header[0] = headerItems[i]; // header must be two-dimensional array
        // Set header items and set text style
        var headerRange = sheet.getRange(1, 1, 1, headerItems[i].length)
          .setValues(header)
          .setHorizontalAlignment('center')
          .setTextStyle(headerStyle);
        // Freeze the first row
        sheet.setFrozenRows(1);
        // Delete empty columns
        sheet.deleteColumns(headerItems[i].length + 1, sheet.getMaxColumns() - headerItems[i].length);
        // Set vertical alignment to 'top'
        sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setVerticalAlignment('top');
      }

      // Update script properties
      var updatedProperties = {
        'currentYear': year,
        'currentSpreadsheetId': createdSpreadsheetId
      };
      if (currentSpreadsheetId !== null) {
        updatedProperties['prevSpreadsheetId'] = currentSpreadsheetId;
      }
      sp.setProperties(updatedProperties, false);

      // Log result
      logText = 'Created spreadsheet';
      log = [logTimestamp, userName, logText]; // one-dimensional array for appendRow()
      logSheet.appendRow(log);

      // Notification by email
      var stop = new Date();
      var executionTime = (stop - now) / 1000; // convert milliseconds to seconds
      var notification = 'New spreadsheet for Toggl record created at ' + createdSpreadsheetUrl + '\nScript execution time: ' + executionTime + ' sec';
      MailApp.sendEmail(myEmail, '[Toggl] New Spreadsheet for Toggl Record Created', notification);
    } else {
      // Log result
      logText = 'Checked: use current spreadsheet';
      log = [logTimestamp, userName, logText];
      currentSpreadsheet.getSheetByName('Log').appendRow(log);
    }
  } catch (e) {
    var thisScriptId = ScriptApp.getScriptId();
    var url = 'https://script.google.com/d/' + thisScriptId + '/edit';
    var body = TogglScript.errorMessage(e) + '\n\nCheck script at ' + url;
    MailApp.sendEmail(myEmail, '[Toggl] Error in Daily Date Check', body)
  }
}

/**
 * Function to create a Google Spreadsheet in a particular Google Drive folder
 * @param {Object} targetFolder - Google Drive folder object in which you want to place the spreadsheet
 * @param {string} ssName - name of spreadsheet 
 * @return {string} ssId - spreadsheet ID of created spreadsheet
 */
function createSpreadsheet_(targetFolder, ssName) {
  var ssId = SpreadsheetApp.create(ssName).getId();
  var temp = DriveApp.getFileById(ssId);
  targetFolder.addFile(temp);
  DriveApp.getRootFolder().removeFile(temp);
  return ssId;
}