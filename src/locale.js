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

/* exported LocalizedMessage */
/* global MESSAGE */

class LocalizedMessage {
  constructor(userLocale = 'en_US') {
    this.locale = userLocale;
    this.messageList = (MESSAGE[this.locale] ? MESSAGE[this.locale] : MESSAGE.en_US); // const MESSAGE is defined on index.js
  }
  /**
   * Replace placeholder values in the designated text. String.prototype.replace() is executed using regular expressions with the 'global' flag on.
   * @param {string} text 
   * @param {array} placeholderValues Array of objects containing a placeholder string expressed in regular expression and its corresponding value.
   * @returns {string} The replaced text.
   */
  replacePlaceholders_(text, placeholderValues = []) {
    let replacedText = placeholderValues.reduce((acc, cur) => acc.replace(new RegExp(cur.regexp, 'g'), cur.value), text);
    return replacedText;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.updateYouTubeAnalyticsDataMailTemplate
   * @param {string} spreadsheetUrl Text to replace the placeholder.
   * @returns {string} The replaced text.
   */
  replaceUpdateYouTubeAnalyticsDataMailTemplate(spreadsheetUrl) {
    let text = this.messageList.youtube.updateYouTubeAnalyticsDataMailTemplate;
    let placeholderValues = [
      {
        'regexp': '\{\{url\}\}',
        'value': spreadsheetUrl
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.newYouTubeSpreadsheetCreatedAlert
   * @param {number|string} year 
   * @param {string} url
   * @returns {string} The replaced text.
   */
  replaceNewYouTubeSpreadsheetCreatedAlert(year, url) {
    let text = this.messageList.youtube.newYouTubeSpreadsheetCreatedAlert;
    let placeholderValues = [
      {
        'regexp': '\{\{year\}\}',
        'value': year
      },
      {
        'regexp': '\{\{url\}\}',
        'value': url
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.newYouTubeSpreadsheetCreatedMailSubject
   * @param {number|string} year
   * @returns {string} The replaced text.
   */
  replaceNewYouTubeSpreadsheetCreatedMailSubject(year) {
    let text = this.messageList.youtube.newYouTubeSpreadsheetCreatedMailSubject;
    let placeholderValues = [
      {
        'regexp': '\{\{year\}\}',
        'value': year
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.newYouTubeSpreadsheetCreatedMailBody
   * @param {number|string} year 
   * @param {string} url
   * @returns {string} The replaced text.
   */
  replaceNewYouTubeSpreadsheetCreatedMailBody(year, url) {
    let text = this.messageList.youtube.newYouTubeSpreadsheetCreatedMailBody;
    let placeholderValues = [
      {
        'regexp': '\{\{year\}\}',
        'value': year
      },
      {
        'regexp': '\{\{url\}\}',
        'value': url
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.errorInvalidChannelId
   * @param {string} startDate 
   * @param {string} endDate
   * @returns {string} The replaced text.
   */
  replaceUpdatedYouTubeChannelAnalyticsLog(startDate, endDate) {
    let text = this.messageList.youtube.updatedYouTubeChannelAnalyticsLog;
    let placeholderValues = [
      {
        'regexp': '\{\{startDate\}\}',
        'value': startDate
      },
      {
        'regexp': '\{\{endDate\}\}',
        'value': endDate
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.updatedYouTubeChannelDemographicsLog
   * @param {string} latestMonth 
   * @param {string} thisYearMonth
   * @returns {string} The replaced text.
   */
  replaceUpdatedYouTubeChannelDemographicsLog(latestMonth, thisYearMonth) {
    let text = this.messageList.youtube.updatedYouTubeChannelDemographicsLog;
    let placeholderValues = [
      {
        'regexp': '\{\{latestMonth\}\}',
        'value': latestMonth
      },
      {
        'regexp': '\{\{thisYearMonth\}\}',
        'value': thisYearMonth
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.updatedYouTubeVideoAnalyticsLog
   * @param {string} startDate 
   * @param {string} endDate
   * @returns {string} The replaced text.
   */
  replaceUpdatedYouTubeVideoAnalyticsLog(startDate, endDate) {
    let text = this.messageList.youtube.updatedYouTubeVideoAnalyticsLog;
    let placeholderValues = [
      {
        'regexp': '\{\{startDate\}\}',
        'value': startDate
      },
      {
        'regexp': '\{\{endDate\}\}',
        'value': endDate
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.errorInvalidChannelId
   * @param {string} targetChannelId
   * @returns {string} The replaced text.
   */
  replaceErrorInvalidChannelId(targetChannelId) {
    let text = this.messageList.youtube.errorInvalidChannelId;
    let placeholderValues = [
      {
        'regexp': '\{\{targetChannelId\}\}',
        'value': targetChannelId
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
  /**
   * Replace placeholder string in this.messageList.youtube.errorInvalidReportMonth
   * @param {string} reportMonth
   * @returns {string} The replaced text.
   */
  replaceErrorInvalidReportMonth(reportMonth) {
    let text = this.messageList.youtube.errorInvalidReportMonth;
    let placeholderValues = [
      {
        'regexp': '\{\{reportMonth\}\}',
        'value': reportMonth
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
    /**
   * Replace placeholder string in this.messageList.youtube.reportCreated
   * @param {string} reportMonth
   * @param {string} targetChannelName
   * @param {string|number} scriptExeTime
   * @returns {string} The replaced text.
   */
  replaceReportCreated(reportMonth, targetChannelName, scriptExeTime) {
    let text = this.messageList.youtube.reportCreated;
    let placeholderValues = [
      {
        'regexp': '\{\{reportMonth\}\}',
        'value': reportMonth
      },
      {
        'regexp': '\{\{targetChannelName\}\}',
        'value': targetChannelName
      },
      {
        'regexp': '\{\{scriptExeTime\}\}',
        'value': scriptExeTime
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}