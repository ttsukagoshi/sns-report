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

/* exported initialSettings, checkSettings */
/* global errorMessage_ */

/**
 * Prompt to set intial settings
 */
function initialSettings() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  if (!scriptProperties.setupComplete || !toBoolean_(scriptProperties.setupComplete)) {
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
    // YouTube Client ID
    let promptYtClientId = 'YouTube Client ID: OAuth 2.0 client ID for this script/spreadsheet. See https://developers.google.com/youtube/reporting/guides/registering_an_application';
    promptYtClientId += (currentSettings.ytClientId ? `\n\nCurrent Value: ${currentSettings.ytClientId}` : '');
    let responseYtClientId = ui.prompt(promptYtClientId, ui.ButtonSet.OK_CANCEL);
    if (responseYtClientId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let ytClientId = responseYtClientId.getResponseText();
    // YouTube Client Secret
    let promptYtClientSecret = 'YouTube Client Secret: OAuth 2.0 client secret for this script/spreadsheet. See https://developers.google.com/youtube/reporting/guides/registering_an_application';
    promptYtClientSecret += (currentSettings.ytClientSecret ? `\n\nCurrent Value: ${currentSettings.ytClientSecret}` : '');
    let responseYtClientSecret = ui.prompt(promptYtClientSecret, ui.ButtonSet.OK_CANCEL);
    if (responseYtClientSecret.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let ytClientSecret = responseYtClientSecret.getResponseText();
    // Facebook App ID
    let promptFbAppId = 'Facebook App ID: ID of Facebook App to process request. See https://developers.facebook.com/docs/facebook-login/access-tokens/';
    promptFbAppId += (currentSettings.fbAppId ? `\n\nCurrent Value: ${currentSettings.fbAppId}` : '');
    let responseFbAppId = ui.prompt(promptFbAppId, ui.ButtonSet.OK_CANCEL);
    if (responseFbAppId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fbAppId = responseFbAppId.getResponseText();
    // Facebook App ID
    let promptFbAppSecret = 'Facebook App Secret: Facebook App Secret to process request. See https://developers.facebook.com/docs/facebook-login/access-tokens/';
    promptFbAppSecret += (currentSettings.fbAppSecret ? `\n\nCurrent Value: ${currentSettings.fbAppSecret}` : '');
    let responseFbAppSecret = ui.prompt(promptFbAppSecret, ui.ButtonSet.OK_CANCEL);
    if (responseFbAppSecret.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fbAppSecret = responseFbAppId.getResponseText();
    // Drive Folder ID
    let promptDriveFolderId = 'Drive Folder ID: The workspace Google Drive folder ID that all analytics spreadsheets are to be stored in.';
    promptDriveFolderId += (currentSettings.driveFolderId ? `\n\nCurrent Value: ${currentSettings.driveFolderId}` : '');
    let responseDriveFolderId = ui.prompt(promptDriveFolderId, ui.ButtonSet.OK_CANCEL);
    if (responseDriveFolderId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let driveFolderId = responseDriveFolderId.getResponseText();
    // Current Year (YouTube)
    let promptYtCurrentYear = 'Current Year (YouTube): Current year for YouTube data in "yyyy" format.';
    promptYtCurrentYear += (currentSettings.ytCurrentYear ? `\n\nCurrent Value: ${currentSettings.ytCurrentYear}` : '');
    let responseYtCurrentYear = ui.prompt(promptYtCurrentYear, ui.ButtonSet.OK_CANCEL);
    if (responseYtCurrentYear.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let ytCurrentYear = responseYtCurrentYear.getResponseText();
    // Current Year (Facebook)
    let promptFbCurrentYear = 'Current Year (Facebook): Current year for Facebook data in "yyyy" format.';
    promptFbCurrentYear += (currentSettings.fbCurrentYear ? `\n\nCurrent Value: ${currentSettings.fbCurrentYear}` : '');
    let responseFbCurrentYear = ui.prompt(promptFbCurrentYear, ui.ButtonSet.OK_CANCEL);
    if (responseFbCurrentYear.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fbCurrentYear = responseFbCurrentYear.getResponseText();
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
      'ytClientId': ytClientId,
      'ytClientSecret': ytClientSecret,
      'fbAppId': fbAppId,
      'fbAppSecret': fbAppSecret,
      'driveFolderId': driveFolderId,
      'ytCurrentYear': ytCurrentYear,
      'fbCurrentYear': fbCurrentYear,
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
