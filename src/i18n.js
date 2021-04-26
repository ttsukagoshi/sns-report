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
const MESSAGE = {
  'en_US': {
    'general': {
      'index': {
        'authorize': 'Authorize',
        'logoutReset': 'Logout/Reset',
        'updateChannelVideoList': 'Update Channel/Video List',
        'updateAll': 'Update All',
        'updateChannelList': 'Update Channel List',
        'updateVideoList': 'Update Video List',
        'analytics': 'Analytics',
        'getLatestData': 'Get Latest Data',
        'createAnalyticsSummary': 'Create Analytics Summary',
        'updatePageList': 'Update Page List',
        'settings': 'Settings',
        'setup': 'Setup',
        'checkSettings': 'Check Settings'
      },
      'error': {
        'errorTitle': 'Error',
        'errorMailSubject': '[SNS Report] Error Detected',
        'spreadsheetUrl_': {
          'templateFileIdIsMissingInOptions': 'The key "templateFileId" is missing in "options".'
        }
      },
      'misc': {
        'completedTitle': 'Completed'
      }
    },
    'youtube': {
      'authorizeYouTubeAPI': 'Authorize YouTube Data/Analytics API',
      'alreadyAuthorized': '[YouTube Data/Analytics API] You are already authorized.',
      'authorizationSuccessful': '[YouTube Data/Analytics API] Success! You can close this tab.',
      'authorizationDenied': '[YouTube Data/Analytics API] Denied. You can close this tab',
      'updatedChannelListLog': 'Success: updated channel list.',
      'updatedChannelListAlert': 'Updated summary channel list.',
      'updatedVideoListLog': 'Success: updated video list.',
      'updatedVideoListAlert': 'Updated summary video list.',
      'updateYouTubeAnalyticsDataMailTemplate': `This is an automatic mail that can be stopped or modified at:\n{{spreadsheetUrl}}\n\nFor more information on the Google Apps Script behind the spreadsheet, see ${GITHUB_URL}`,
      'newYouTubeSpreadsheetCreatedAlert': 'New YouTube spreadsheet created for {{year}}:\n{{url}}',
      'newYouTubeSpreadsheetCreatedMailSubject': '[SNS Report] New YouTube Spreadsheet Created for {{year}}',
      'newYouTubeSpreadsheetCreatedMailBody': 'New spreadsheet created to record YouTube analytics data for year {{year}} at:\n{{url}}\n\n',
      'updatedYouTubeChannelAnalyticsLog': 'Success: updated YouTube channel analytics for {{startDate}} to {{endDate}}.',
      'noUpdatesForYouTubeChannelAnalyticsLog': 'Success: no updates for YouTube channel analytics.',
      'updatedYouTubeChannelDemographicsLog': 'Success: updated YouTube channel demographics for {{latestMonth}} to {{thisYearMonth}}.',
      'updatedYouTubeVideoAnalyticsLog': 'Success: updated YouTube video analytics for {{startDate}} to {{endDate}}.',
      'noUpdatesForYouTubeVideoAnalyticsLog': 'Success: no updates for YouTube video analytics.',
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > YouTube > Authorize',
      'errorNoTextEnteredForChannelId': 'No text entered for channel ID.',
      'errorInvalidChannelId': 'Invalid Channel ID: "{{targetChannelId}}".\nMake sure to enter the ID of the YouTube channel that you own.',
      'errorNoTextEnteredForReportMonth': 'No text entered for report month.',
      'errorInvalidReportMonth': 'Invalid Report Month: "{{reportMonth}}".\nMake sure the report month is expressed in "yyyy-MM", e.g., enter "2020-03" for getting report for March 2020.',
      'reportCreated': 'Report for {{reportMonth}} of YouTube channel "{{targetChannelName}}" created.\nScript Time: {{scriptExeTime}} secs.'
    },
    'facebook': {
      'authorizeFacebookAPI': 'Authorize Facebook Graph API',
      'alreadyAuthorized': '[Facebook API] You are already authorized.',
      'authorizationSuccessful': '[Facebook API] Success! You can close this tab.',
      'authorizationDenied': '[Facebook API] Denied. You can close this tab',
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > Facebook > Authorize',
      'updatedPageListLog': 'Success: updated page list.',
      'updatedPageListAlert': 'Updated summary page list.',
      'updatedPagePostListLog': 'Success: updated page post list.',
      'updatedPagePostListAlert': 'Updated page post list.'
    }
  },
  'ja_JP': {
    'general': {
      'index': {
        'authorize': '認証',
        'logoutReset': 'ログアウト/リセット',
        'updateChannelVideoList': 'チャンネル/ビデオ一覧を更新',
        'updateAll': 'すべてを更新',
        'updateChannelList': 'チャンネル一覧を更新',
        'updateVideoList': 'ビデオ一覧を更新',
        'analytics': 'アナリティクス',
        'getLatestData': '最新データを取得',
        'createAnalyticsSummary': 'アナリティクスのサマリー作成',
        'updatePageList': 'ページ一覧を更新',
        'settings': '設定',
        'setup': '初期設定',
        'checkSettings': '設定確認'
      },
      'error': {
        'errorTitle': 'エラー',
        'errorMailSubject': '[SNS Report] Error Detected',
        'spreadsheetUrl_': {
          'templateFileIdIsMissingInOptions': '変数"option"内のキー"templateFileId"が指定されていません。'
        }
      },
      'misc': {
        'completedTitle': '完了'
      }
    },
    'youtube': {
      'authorizeYouTubeAPI': 'YouTube Data/Analytics APIを認証',
      'alreadyAuthorized': '[YouTube Data/Analytics API] すでに認証済みです。',
      'authorizationSuccessful': '[YouTube Data/Analytics API] 認証成功。このタブを閉じても大丈夫です。',
      'authorizationDenied': '[YouTube Data/Analytics API] 認証に失敗しました。このタブは閉じてください。',
      'updatedChannelListLog': 'Success: updated channel list.', // Log message will not be translated.
      'updatedChannelListAlert': 'チャンネル一覧を更新完了。',
      'updatedVideoListLog': 'Success: updated video list.', // Log message will not be translated.
      'updatedVideoListAlert': 'ビデオ一覧を更新完了。',
      'updateYouTubeAnalyticsDataMailTemplate': `This is an automatic mail that can be stopped or modified at:\n{{spreadsheetUrl}}\n\nFor more information on the Google Apps Script behind the spreadsheet, see ${GITHUB_URL}`,
      'newYouTubeSpreadsheetCreatedAlert': 'New YouTube spreadsheet created for {{year}}:\n{{url}}',
      'newYouTubeSpreadsheetCreatedMailSubject': '[SNS Report] New YouTube Spreadsheet Created for {{year}}',
      'newYouTubeSpreadsheetCreatedMailBody': 'New spreadsheet created to record YouTube analytics data for year {{year}} at:\n{{url}}\n\n',
      'updatedYouTubeChannelAnalyticsLog': 'Success: updated YouTube channel analytics for {{startDate}} to {{endDate}}.', // Log message will not be translated.
      'noUpdatesForYouTubeChannelAnalyticsLog': 'Success: no updates for YouTube channel analytics.', // Log message will not be translated.
      'updatedYouTubeChannelDemographicsLog': 'Success: updated YouTube channel demographics for {{latestMonth}} to {{thisYearMonth}}.', // Log message will not be translated.
      'updatedYouTubeVideoAnalyticsLog': 'Success: updated YouTube video analytics for {{startDate}} to {{endDate}}.', // Log message will not be translated.
      'noUpdatesForYouTubeVideoAnalyticsLog': 'Success: no updates for YouTube video analytics.', // Log message will not be translated.
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > YouTube > Authorize',
      'errorNoTextEnteredForChannelId': 'No text entered for channel ID.',
      'errorInvalidChannelId': 'Invalid Channel ID: "{{targetChannelId}}".\nMake sure to enter the ID of the YouTube channel that you own.',
      'errorNoTextEnteredForReportMonth': 'No text entered for report month.',
      'errorInvalidReportMonth': 'Invalid Report Month: "{{reportMonth}}".\nMake sure the report month is expressed in "yyyy-MM", e.g., enter "2020-03" for getting report for March 2020.',
      'reportCreated': 'Report for {{reportMonth}} of YouTube channel "{{targetChannelName}}" created.\nScript Time: {{scriptExeTime}} secs.'
    },
    'facebook': {
      'authorizeFacebookAPI': 'Facebook Graph APIを認証',
      'alreadyAuthorized': '[Facebook API] すでに認証済みです。',
      'authorizationSuccessful': '[Facebook API] 認証成功。このタブを閉じても大丈夫です。',
      'authorizationDenied': '[Facebook API] 認証に失敗しました。このタブは閉じてください。',
      'errorUnauthorized': 'Unauthorized. Get authorized by Menu > Facebook > Authorize',
      'updatedPageListLog': 'Success: updated page list.', // Log message will not be translated.
      'updatedPageListAlert': 'Updated summary page list.',
      'updatedPagePostListLog': 'Success: updated page post list.', // Log message will not be translated.
      'updatedPagePostListAlert': 'Updated page post list.'
    }
  }
};

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