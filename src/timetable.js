const dayjs = require('dayjs');
const { readJson, writeJson } = require('./config');
const { isGoogleSheetEnabled, callGoogleSheet } = require('./googleSheet');

const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const KHMER_DAYS = {
  'Sunday': 'អាទិត្យ',
  'Monday': 'ច័ន្ទ',
  'Tuesday': 'អង្គារ',
  'Wednesday': 'ពុធ',
  'Thursday': 'ព្រហស្បតិ៍',
  'Friday': 'សុក្រ',
  'Saturday': 'សៅរ៍'
};

let memoryTimetable = null;
let hasSyncedTimetableFromGoogleSheet = false;

function getTimetable() {
  if (Array.isArray(memoryTimetable)) {
    return memoryTimetable;
  }
  const fileData = readJson('timetable.json', []);
  memoryTimetable = fileData;
  return memoryTimetable;
}

function saveTimetable(timetable) {
  memoryTimetable = timetable;
  writeJson('timetable.json', timetable);
}

async function syncTimetableFromGoogleSheet() {
  if (!isGoogleSheetEnabled()) return;

  try {
    const res = await callGoogleSheet('getTimetable');
    if (res && res.status === 'success' && Array.isArray(res.timetable)) {
      hasSyncedTimetableFromGoogleSheet = true;
      if (res.timetable.length > 0) {
        memoryTimetable = res.timetable;
      } else {
        // Sheet exists but is empty -> populate with default template from local timetable.json
        const defaultTimetable = readJson('timetable.json', []);
        if (defaultTimetable.length > 0) {
          await callGoogleSheet('saveTimetable', { timetable: defaultTimetable });
          memoryTimetable = defaultTimetable;
        } else {
          memoryTimetable = [];
        }
      }
    }
  } catch (err) {
    console.error('[Timetable] Error syncing timetable from Google Sheet:', err.message);
  }
}

function getSubjectsDict() {
  return readJson('subjects.json', {});
}

function saveSubjectsDict(dict) {
  writeJson('subjects.json', dict);
}

/**
 * Update the Khmer subject name for a code (e.g. M2 -> គណិតវិទ្យា)
 * This automatically syncs all entries in timetable as well.
 */
async function updateSubjectName(code, newName) {
  const upperCode = code.trim().toUpperCase();
  const dict = getSubjectsDict();
  dict[upperCode] = newName.trim();
  saveSubjectsDict(dict);

  // Sync timetable
  const timetable = getTimetable();
  timetable.forEach(slot => {
    if (slot.code && slot.code.toUpperCase() === upperCode) {
      slot.subject = newName.trim();
    }
  });
  saveTimetable(timetable);

  if (isGoogleSheetEnabled()) {
    try {
      await callGoogleSheet('saveTimetable', { timetable });
      await callGoogleSheet('saveSubjects', { subjects: dict });
    } catch (e) {}
  }

  return dict;
}

/**
 * Returns current class slot based on system / local time.
 */
function getCurrentSlot(now = dayjs()) {
  const timetable = getTimetable();
  const dayName = DAYS[now.day()];
  const currentTime = now.format('HH:mm');

  // Find exact slot where start <= currentTime < end
  const current = timetable.find(slot => {
    return slot.day.toLowerCase() === dayName.toLowerCase() &&
           currentTime >= slot.start &&
           currentTime <= slot.end;
  });

  if (current) return { ...current, isExactNow: true, khmerDay: KHMER_DAYS[dayName] };

  // If outside exact class hours, look for slots today
  const todaySlots = timetable
    .filter(slot => slot.day.toLowerCase() === dayName.toLowerCase())
    .sort((a, b) => a.start.localeCompare(b.start));

  if (todaySlots.length === 0) return null;

  // Check if upcoming today
  const upcoming = todaySlots.find(slot => slot.start > currentTime);
  if (upcoming) return { ...upcoming, isUpcoming: true, khmerDay: KHMER_DAYS[dayName] };

  // Return the last class of today if day's classes have finished
  return { ...todaySlots[todaySlots.length - 1], isLastOfToday: true, khmerDay: KHMER_DAYS[dayName] };
}

function getTodaySlots(now = dayjs()) {
  const timetable = getTimetable();
  const dayName = DAYS[now.day()];
  return timetable
    .filter(slot => slot.day.toLowerCase() === dayName.toLowerCase())
    .sort((a, b) => a.start.localeCompare(b.start))
    .map(slot => ({ ...slot, khmerDay: KHMER_DAYS[dayName] }));
}

function getAllSubjects() {
  const timetable = getTimetable();
  const subjects = new Set();
  timetable.forEach(s => subjects.add(s.subject));
  return Array.from(subjects);
}

module.exports = {
  getTimetable,
  saveTimetable,
  syncTimetableFromGoogleSheet,
  getSubjectsDict,
  saveSubjectsDict,
  updateSubjectName,
  getCurrentSlot,
  getTodaySlots,
  getAllSubjects,
  DAYS,
  KHMER_DAYS
};
