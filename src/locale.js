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
      'updateYouTubeAnalyticsDataMailTemplate': 'This is an automatic mail that can be stopped or modified at:\n{{spreadsheetUrl}}\n\nFor more information on the Google Apps Script behind the spreadsheet, see https://github.com/ttsukagoshi/sns-report',
      'newYouTubeSpreadsheetCreatedAlert': 'New YouTube spreadsheet created for {{year}}:\n{{newFileUrl}}'
    },
    'facebook': {
      'authorizeFacebookAPI': 'Authorize Facebook Graph API',
      'alreadyAuthorized': '[Facebook API] You are already authorized.',
      'authorizationSuccessful': '[Facebook API] Success! You can close this tab.',
      'authorizationDenied': '[Facebook API] Denied. You can close this tab'
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
      'updateYouTubeAnalyticsDataMailTemplate': 'This is an automatic mail that can be stopped or modified at:\n{{spreadsheetUrl}}\n\nFor more information on the Google Apps Script behind the spreadsheet, see https://github.com/ttsukagoshi/sns-report',
      'newYouTubeSpreadsheetCreatedAlert': 'New YouTube spreadsheet created for {{year}}:\n{{newFileUrl}}'
    },
    'facebook': {
      'authorizeFacebookAPI': 'Facebook Graph APIを認証',
      'alreadyAuthorized': '[Facebook API] すでに認証済みです。',
      'authorizationSuccessful': '[Facebook API] 認証成功。このタブを閉じても大丈夫です。',
      'authorizationDenied': '[Facebook API] 認証に失敗しました。このタブは閉じてください。'
    }
  }
};
class LocalizedMessage {
  constructor(userLocale = 'en_US') {
    this.locale = userLocale;
    this.messageList = (MESSAGE[this.locale] ? MESSAGE[this.locale] : MESSAGE.en_US);
  }
  /**
   * 
   * @param {string} text 
   * @param {array} placeholderValues Array of objects containing a pair of placeholder string (key) and its corresponding value.
   * @returns {string} The replaced text.
   */
  replacePlaceholders_(text, placeholderValues = []) {
    let replacedText = placeholderValues.reduce((acc, cur) => acc.replace(cur.regex, cur.value), text);
    return replacedText;
  }

  /**
   * Replace placeholder string in this.messageList.youtube.updateYouTubeAnalyticsDataMailTemplate
   * @param {string} spreadsheetUrl Text to replace the placeholder.
   */
  replaceUpdateYouTubeAnalyticsDataMailTemplate(spreadsheetUrl) {
    let text = this.messageList.youtube.updateYouTubeAnalyticsDataMailTemplate.slice();
    text = text.replace(/\{\{spreadsheetUrl\}\}/g, spreadsheetUrl);
    this.messageList.youtube.updateYouTubeAnalyticsDataMailTemplate = text.slice();
    return text;
  }
  /**
   * 
   * @param {number|string} year 
   * @param {string} newFileUrl 
   */
  replaceNewYouTubeSpreadsheetCreatedAlert(year, newFileUrl) {
    let text = this.messageList.youtube.newYouTubeSpreadsheetCreatedAlert.slice();
    text = text.replace(/\{\{year\}\}/g, year)
      .replace(/\{\{newFileUrl\}\}/g, newFileUrl);
    this.messageList.youtube.newYouTubeSpreadsheetCreatedAlert = text.slice();
    return text;
  }
}