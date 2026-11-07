/**
 * Google Sheet API Client for PineApple Attendance Bot
 * Connects to Google Apps Script Web App Endpoint
 */
function getApiUrl() {
  return process.env.GOOGLE_SHEET_API_URL;
}

function isGoogleSheetEnabled() {
  const url = getApiUrl();
  return Boolean(url && url.startsWith('http'));
}

async function callGoogleSheet(action, payload = {}, params = {}) {
  if (!isGoogleSheetEnabled()) return null;

  try {
    const apiUrl = getApiUrl();
    const query = new URLSearchParams({ action, ...params }).toString();
    const url = `${apiUrl}?${query}`;

    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 4000);

    const options = {
      method: payload && Object.keys(payload).length > 0 ? 'POST' : 'GET',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: payload && Object.keys(payload).length > 0 ? JSON.stringify({ action, ...payload }) : undefined,
      signal: controller.signal
    };

    let response;
    try {
      response = await fetch(url, options);
    } finally {
      clearTimeout(timeoutId);
    }
    if (!response.ok) {
      console.warn(`[GoogleSheet] HTTP Error: ${response.status}`);
      return null;
    }
    const data = await response.json();
    return data;
  } catch (err) {
    console.error(`[GoogleSheet] Error calling ${action}:`, err.message);
    return null;
  }
}

// Upload initial roster to Google Sheet
async function syncLocalToGoogleSheet(students, timetable, subjects) {
  if (!isGoogleSheetEnabled()) return;
  console.log('🔄 Syncing local data to Google Sheet...');
  if (students && students.length > 0) {
    await callGoogleSheet('saveStudents', { students });
  }
  if (timetable && timetable.length > 0) {
    await callGoogleSheet('saveTimetable', { timetable });
  }
  if (subjects && Object.keys(subjects).length > 0) {
    await callGoogleSheet('saveSubjects', { subjects });
  }
  console.log('✅ Local data synced to Google Sheet successfully!');
}

module.exports = {
  isGoogleSheetEnabled,
  callGoogleSheet,
  syncLocalToGoogleSheet
};
