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