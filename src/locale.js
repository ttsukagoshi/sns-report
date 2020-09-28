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
      }
    },
    'youtube': {},
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
      }
    },
    'youtube': {},
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
}