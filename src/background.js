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

/* exported weeklyAnalyticsUpdate */
/* global updateYouTubeAnalyticsData */

/**
 * Periodically update analytics data using Google Apps Script's trigger.
 */
function weeklyAnalyticsUpdate() {
  console.log('[weeklyAnalyticsUpdate] Initiating: A periodical task to update analytics data using Google Apps Script\'s trigger...'); // log
  var muteUiAlert = true;
  var muteMailNotification = false;
  var yearLimit = false;
  try {
    updateYouTubeAnalyticsData(muteUiAlert, muteMailNotification, yearLimit);
    console.log('[weeklyAnalyticsUpdate] Complete: Periodical task to update analytics data using Google Apps Script\'s trigger is complete.'); // log
  } catch (error) {
    console.log('[weeklyAnalyticsUpdate] Terminated.'); // log
  }
}
