function fetchYouTubeRevenueWithTitles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Dashboard');
  const rawVal = sheet.getRange('B1').getValue();
  const threshold = (rawVal === '' || rawVal === null) ? 0.30 : Number(rawVal);
  const channelId = 'YOUR_CHANNEL_ID'; // Replace with your UC... Channel ID

  // Calculate the 28-day window (accounting for YouTube's 2-day reporting lag)
  const endDate = new Date();
  endDate.setDate(endDate.getDate() - 2);
  const startDate = new Date();
  startDate.setDate(endDate.getDate() - 28);
  const formatDate = (date) => Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const startStr = formatDate(startDate);
  const endStr = formatDate(endDate);

  // 1. Fetch total channel-wide revenue and display date range
  const channelTotalRes = YouTubeAnalytics.Reports.query({
    ids: 'channel==' + channelId,
    startDate: startStr,
    endDate: endStr,
    metrics: 'estimatedRevenue'
  });
  const fullChannelRevenue = (channelTotalRes.rows && channelTotalRes.rows[0]) ? channelTotalRes.rows[0][0] : 0;

  sheet.getRange('D1').setValue('Full Channel Revenue ($)');
  sheet.getRange('E1').setValue(fullChannelRevenue);
  sheet.getRange('D2').setValue('Date Range:');
  sheet.getRange('E2').setValue(startStr + ' to ' + endStr);

  // 2. Get your channel's Uploads playlist ID
  const channelRes = YouTube.Channels.list('contentDetails', { id: channelId });
  if (!channelRes.items || channelRes.items.length === 0) {
    throw new Error('Channel not found.');
  }
  const uploadsPlaylistId = channelRes.items[0].contentDetails.relatedPlaylists.uploads;

  // 3. Fetch all Video IDs and Titles from your Uploads playlist
  const allVideos = [];
  const titleMap = {};
  let pageToken = '';

  do {
    const playlistRes = YouTube.PlaylistItems.list('snippet', {
      playlistId: uploadsPlaylistId,
      maxResults: 50,
      pageToken: pageToken
    });

    if (playlistRes.items) {
      playlistRes.items.forEach(item => {
        const vid = item.snippet.resourceId.videoId;
        const title = item.snippet.title;
        allVideos.push(vid);
        titleMap[vid] = title;
      });
    }
    pageToken = playlistRes.nextPageToken || '';
  } while (pageToken);

  // 4. Query YouTube Analytics API in batches of 200 videos
  const qualifyingRows = [];

  for (let i = 0; i < allVideos.length; i += 200) {
    const batchIds = allVideos.slice(i, i + 200).join(',');
    const analyticsRes = YouTubeAnalytics.Reports.query({
      ids: 'channel==' + channelId,
      startDate: startStr,
      endDate: endStr,
      metrics: 'views,estimatedRevenue',
      dimensions: 'video',
      filters: 'video==' + batchIds,
      maxResults: 200
    });

    if (analyticsRes.rows) {
      analyticsRes.rows.forEach(row => {
        const vid = row[0];
        const views = Number(row[1]);
        const rev = Number(row[2]);
        const rpm = views > 0 ? Number(((rev / views) * 1000).toFixed(2)) : 0;
        if (rev > threshold) {
          qualifyingRows.push([vid, titleMap[vid] || '(Title unavailable)', views, rev, rpm]);
        }
      });
    }
  }

  // 5. Sort highest-earning videos to the top
  qualifyingRows.sort((a, b) => b[3] - a[3]);

  // 6. Clear old data and write new results to Dashboard
  const lastRow = sheet.getLastRow();
  if (lastRow >= 5) {
    sheet.getRange(5, 1, lastRow - 4, 5).clearContent();
  }

  if (qualifyingRows.length > 0) {
    sheet.getRange(5, 1, qualifyingRows.length, 5).setValues(qualifyingRows);
  }
}
