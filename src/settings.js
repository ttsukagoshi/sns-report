/**
 * Prompt to set intial settings
 */
function initialSettings() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  if (!scriptProperties.setupComplete || toBoolean_(scriptProperties.setupComplete) == false) {
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
    // Client ID
    let promptClientId = 'Client ID: OAuth 2.0 client ID for this script/spreadsheet. See https://developers.google.com/youtube/reporting/guides/registering_an_application';
    promptClientId += (currentSettings.clientId ? `\n\nCurrent Value: ${currentSettings.clientId}` : '');
    let responseClientId = ui.prompt(promptClientId, ui.ButtonSet.OK_CANCEL);
    if (responseClientId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let clientId = responseClientId.getResponseText();
    // Client Secret
    let promptClientSecret = 'Client Secret: OAuth 2.0 client secret for this script/spreadsheet. See https://developers.google.com/youtube/reporting/guides/registering_an_application';
    promptClientSecret += (currentSettings.clientSecret ? `\n\nCurrent Value: ${currentSettings.clientSecret}` : '');
    let responseClientSecret = ui.prompt(promptClientSecret, ui.ButtonSet.OK_CANCEL);
    if (responseClientSecret.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    var clientSecret = responseClientSecret.getResponseText();
    // Drive Folder ID
    let promptDriveFolderId = 'Drive Folder ID: The workspace Google Drive folder ID that all analytics spreadsheets are to be stored in.';
    promptDriveFolderId += (currentSettings.driveFolderId ? `\n\nCurrent Value: ${currentSettings.driveFolderId}` : '');
    let responseDriveFolderId = ui.prompt(promptDriveFolderId, ui.ButtonSet.OK_CANCEL);
    if (responseDriveFolderId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let driveFolderId = responseDriveFolderId.getResponseText();
    // Current Year
    let promptCurrentYear = 'Current Year: Current year in "yyyy" format.';
    promptCurrentYear += (currentSettings.currentYear ? `\n\nCurrent Value: ${currentSettings.currentYear}` : '');
    let responseCurrentYear = ui.prompt(promptCurrentYear, ui.ButtonSet.OK_CANCEL);
    if (responseCurrentYear.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let currentYear = responseCurrentYear.getResponseText();
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
      'clientId': clientId,
      'clientSecret': clientSecret,
      'driveFolderId': driveFolderId,
      'currentYear': currentYear,
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
