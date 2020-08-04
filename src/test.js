function archiveCatchup() {
  var targetYears = [2019];
  var yearLimit = true;
  // targetYears.forEach(value => youtubeAnalyticsVideo(value, yearLimit))
}

function test() {
  var channelIds = ['UCoCQUNvO3Dx8z919YMoNZ0A', 'UCZXk0hDQWjtcKWcXYvqLuyA'] // [owner, non-owner] channels, respectively
  var results = youtubeChannels_(channelIds, false);
  console.log(results);
}

function test_videoAnalytics() {
  var startDate = '2020-07-01';
  var endDate = '2020-07-31';
  var metrics = 'views,likes,dislikes,subscribersGained,subscribersLost';
  var ids = 'channel==MINE';
  var videoList = youtubeMyVideoList_();
  var videoAnalytics = videoList.map(function(video) {
    let videoId = video.id.videoId;
    let analytics = {
      'videoId': videoId,
      'channelId': video.snippet.channelId,
      'data': JSON.parse(youtubeAnalyticsReportsQuery_(startDate,endDate,metrics,ids,{
        'dimensions': 'day',
        'filters': `video==${videoId}`
      })).rows
    };
    return analytics;
  });
  console.log(videoAnalytics);/////////////////////
  
}

function test_1() {
  var startDate = '2020-08-01';
  var endDate = '2020-09-01';
  var metrics = 'viewerPercentage';
  var ids = 'channel==MINE';
  var options = {
    dimensions: 'ageGroup,gender'
  }

  // Get analytics data
  var reports = JSON.parse(youtubeAnalyticsReportsQuery_(startDate, endDate, metrics, ids, options));
  console.log(reports);
}
