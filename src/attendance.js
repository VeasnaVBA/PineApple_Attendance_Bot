const fs = require('fs');
const path = require('path');
const dayjs = require('dayjs');
const { readJson, writeJson, RECORDS_DIR } = require('./config');
const { isGoogleSheetEnabled, callGoogleSheet } = require('./googleSheet');

function getStudents() {
  return readJson('students.json', []);
}

function saveStudents(students) {
  writeJson('students.json', students);
}

function normalizeStr(str) {
  if (!str) return '';
  // Remove zero-width spaces and normalize whitespace
  return str.replace(/[\u200B-\u200D\uFEFF]/g, '').trim().toLowerCase();
}

function addStudent(id, name, gender = 'Unknown', dob = '') {
  const students = getStudents();
  // Pad single digit id e.g. 1 -> 01
  const cleanId = String(id).padStart(2, '0');
  const existing = students.find(s => s.id === cleanId || s.id === String(id));
  if (existing) {
    existing.name = name;
    if (gender) existing.gender = gender;
    if (dob) existing.dob = dob;
  } else {
    students.push({
      id: cleanId,
      name,
      gender: gender || 'Unknown',
      dob: dob || ''
    });
    // Sort students by ID
    students.sort((a, b) => {
      const numA = parseInt(a.id, 10);
      const numB = parseInt(b.id, 10);
      if (!isNaN(numA) && !isNaN(numB)) return numA - numB;
      return a.id.localeCompare(b.id);
    });
  }
  saveStudents(students);
  return students;
}

function updateStudent(id, newName, gender, dob) {
  const students = getStudents();
  const cleanId = String(id).padStart(2, '0');
  const stu = students.find(s => s.id === cleanId || s.id === String(id));
  if (!stu) return null;

  if (newName) stu.name = newName;
  if (gender) stu.gender = gender;
  if (dob) stu.dob = dob;

  saveStudents(students);
  return stu;
}

function removeStudent(query) {
  let students = getStudents();
  const initialLen = students.length;
  const q = normalizeStr(query);
  const qPadded = q.padStart(2, '0');

  students = students.filter(s => {
    const sId = normalizeStr(s.id);
    const sName = normalizeStr(s.name);
    return sId !== q && sId !== qPadded && sName !== q;
  });

  if (students.length !== initialLen) {
    saveStudents(students);
    return true;
  }
  return false;
}

/**
 * Match a raw name string to an existing student in students.json.
 */
function findStudentMatch(nameQuery, students) {
  const q = normalizeStr(nameQuery);
  if (!q) return null;

  const qPadded = q.padStart(2, '0');

  // 1. Exact ID match (e.g. "04" or "4")
  const idMatch = students.find(s => normalizeStr(s.id) === q || normalizeStr(s.id) === qPadded);
  if (idMatch) return idMatch;

  // 2. Exact full name match
  const exactName = students.find(s => normalizeStr(s.name) === q);
  if (exactName) return exactName;

  // 3. Word boundary / part match
  // e.g. "ចិន" in "ឆោម ចិន"
  const wordMatch = students.find(s => {
    const sName = normalizeStr(s.name);
    const parts = sName.split(/\s+/);
    return parts.includes(q) || sName.startsWith(q) || sName.endsWith(q);
  });
  if (wordMatch) return wordMatch;

  // 4. Substring match
  const subMatch = students.find(s => normalizeStr(s.name).includes(q));
  if (subMatch) return subMatch;

  return null;
}

/**
 * Searches for all students matching a query string (name or number/ID).
 */
function searchStudents(query) {
  const q = normalizeStr(query);
  if (!q) return [];
  const students = getStudents();
  const qPadded = q.padStart(2, '0');

  return students.filter(s => {
    const sId = normalizeStr(s.id);
    const sName = normalizeStr(s.name);
    return sId === q || sId === qPadded || sName.includes(q);
  });
}

/**
 * Parses attendance text from user.
 * Supports patterns like:
 *   ឆោម ចិន A
 *   ចិន A
 *   04 A
 *   Keo Sopheap P (Sick)
 *   1. ឆោម ចិន A
 *   - Sokha P
 */
function parseAttendanceInput(text) {
  const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
  const parsedItems = [];
  const unrecognized = [];

  // Regex matches: optional numbering/bullet, then name or id, then status code (P|A|L), then optional note
  const lineRegex = /^(?:[\d\.\-\*\•\)\s]+)?(.+?)[\s\:\=\-]+([PALpal])(?:\s*[\:\-]?\s*(.*))?$/;

  for (const line of lines) {
    if (line.startsWith('/')) continue;

    const match = line.match(lineRegex);
    if (match) {
      const rawName = match[1].trim();
      const code = match[2].toUpperCase();
      const note = match[3] ? match[3].trim() : '';

      let status = 'Present';
      if (code === 'A') status = 'Absent';
      else if (code === 'P') status = 'Permission';
      else if (code === 'L') status = 'Late';

      parsedItems.push({
        rawName,
        code,
        status,
        note
      });
    } else {
      unrecognized.push(line);
    }
  }

  return { parsedItems, unrecognized };
}

function getRecordFilePath(dateStr) {
  return path.join(RECORDS_DIR, `attendance-${dateStr}.json`);
}

function getDailyRecords(dateStr) {
  const filePath = getRecordFilePath(dateStr);
  if (fs.existsSync(filePath)) {
    try {
      return JSON.parse(fs.readFileSync(filePath, 'utf8'));
    } catch (e) {
      console.error(`Error reading ${filePath}`, e);
    }
  }
  return {};
}

function saveDailyRecords(dateStr, records) {
  try {
    const filePath = getRecordFilePath(dateStr);
    fs.writeFileSync(filePath, JSON.stringify(records, null, 2), 'utf8');
  } catch (err) {
    console.warn(`[attendance] Note: Could not write daily records to local disk:`, err.message);
  }
}

/**
 * Record attendance for a given subject and date.
 * All registered students not in parsedItems default to Present!
 */
function recordAttendance({ dateStr, subject, timeSlot, recordedBy, parsedItems }) {
  const students = getStudents();
  const daily = getDailyRecords(dateStr);

  const existingSubjectRecord = daily[subject] || {
    date: dateStr,
    subject,
    timeSlot: timeSlot || 'Class Time',
    recordedBy: recordedBy || 'Class Monitor',
    createdAt: new Date().toISOString(),
    records: []
  };

  const matchedStudentIds = new Set();
  const notFoundNames = [];

  const statusMap = {};
  for (const item of parsedItems) {
    const matched = findStudentMatch(item.rawName, students);
    if (matched) {
      matchedStudentIds.add(matched.id);
      statusMap[matched.id] = {
        code: item.code,
        status: item.status,
        note: item.note
      };
    } else {
      notFoundNames.push(item.rawName);
    }
  }

  // Build complete list of all 30 students
  const fullRecords = students.map((stu, index) => {
    const mark = statusMap[stu.id];
    if (mark) {
      return {
        no: index + 1,
        id: stu.id,
        name: stu.name,
        gender: stu.gender || 'Unknown',
        dob: stu.dob || '',
        status: mark.status,
        code: mark.code,
        note: mark.note || ''
      };
    }
    // Default to Present
    return {
      no: index + 1,
      id: stu.id,
      name: stu.name,
      gender: stu.gender || 'Unknown',
      dob: stu.dob || '',
      status: 'Present',
      code: '✓',
      note: ''
    };
  });

  const summary = {
    total: fullRecords.length,
    present: fullRecords.filter(r => r.status === 'Present').length,
    absent: fullRecords.filter(r => r.status === 'Absent').length,
    permission: fullRecords.filter(r => r.status === 'Permission').length,
    late: fullRecords.filter(r => r.status === 'Late').length
  };

  const updatedSubjectRecord = {
    ...existingSubjectRecord,
    subject,
    timeSlot: timeSlot || existingSubjectRecord.timeSlot,
    recordedBy: recordedBy || existingSubjectRecord.recordedBy,
    updatedAt: new Date().toISOString(),
    summary,
    records: fullRecords
  };

  // Key by timeSlot so multiple sessions of the same subject on the same day don't overwrite
  const slotKey = timeSlot ? timeSlot.trim() : subject;
  daily[slotKey] = updatedSubjectRecord;
  daily[subject] = updatedSubjectRecord; // fallback reference
  saveDailyRecords(dateStr, daily);

  // Sync to Google Sheet if configured
  if (isGoogleSheetEnabled()) {
    callGoogleSheet('recordAttendance', {
      dateStr,
      timeSlot,
      subject,
      recordedBy,
      fullRecords
    }).catch(e => console.warn('[GoogleSheet Sync Error]:', e.message));
  }

  return {
    subjectRecord: updatedSubjectRecord,
    summary,
    matchedCount: matchedStudentIds.size,
    notFoundNames
  };
}

/**
 * Returns complete day status including all scheduled sessions,
 * whether they were recorded or forgotten by the class monitor.
 */
function getDayAttendanceStatus(dateStr, timetableHelper) {
  const targetDate = dayjs(dateStr);
  const { DAYS, KHMER_DAYS, getTodaySlots } = timetableHelper;
  const dayName = DAYS[targetDate.day()];
  const khmerDay = KHMER_DAYS[dayName];
  const slots = getTodaySlots(targetDate);
  const daily = getDailyRecords(dateStr);
  const students = getStudents();

  if (slots.length === 0) {
    return {
      dateStr,
      dayName,
      khmerDay,
      hasTimetable: false,
      scheduledSessions: [],
      totalSessions: 0,
      recordedCount: 0,
      missedCount: 0,
      students
    };
  }

  const sessions = slots.map((slot, idx) => {
    const slotTime = `${slot.start} - ${slot.end}`;
    // Find record by exact slotTime or fallback
    let rec = daily[slotTime] || daily[`${slot.start}-${slot.end}`];
    if (!rec) {
      const foundKey = Object.keys(daily).find(k => {
        const item = daily[k];
        return item && (item.timeSlot === slotTime || item.subject === slot.subject);
      });
      if (foundKey) rec = daily[foundKey];
    }

    if (rec && rec.records && rec.records.length > 0) {
      return {
        sessionNo: idx + 1,
        timeSlot: slotTime,
        subject: slot.subject,
        code: slot.code || '',
        isRecorded: true,
        recordedBy: rec.recordedBy || 'ប្រធានថ្នាក់',
        summary: rec.summary,
        records: rec.records
      };
    } else {
      // Missed / forgotten by class monitor!
      return {
        sessionNo: idx + 1,
        timeSlot: slotTime,
        subject: slot.subject,
        code: slot.code || '',
        isRecorded: false,
        recordedBy: null,
        summary: null,
        records: []
      };
    }
  });

  const recordedCount = sessions.filter(s => s.isRecorded).length;
  const missedCount = sessions.filter(s => !s.isRecorded).length;

  return {
    dateStr,
    dayName,
    khmerDay,
    hasTimetable: true,
    scheduledSessions: sessions,
    totalSessions: sessions.length,
    recordedCount,
    missedCount,
    students
  };
}

/**
 * Computes attendance summary for a custom date range (From startDate to endDate)
 */
function getRangeAttendanceStatus(startDateStr, endDateStr, timetableHelper) {
  const start = dayjs(startDateStr);
  const end = dayjs(endDateStr);
  const students = getStudents();

  const daysList = [];
  let curr = start;

  while (curr.isBefore(end) || curr.isSame(end, 'day')) {
    const dStr = curr.format('YYYY-MM-DD');
    const dayStat = getDayAttendanceStatus(dStr, timetableHelper);
    if (dayStat.hasTimetable) {
      daysList.push(dayStat);
    }
    curr = curr.add(1, 'day');
  }

  const studentMap = {};
  students.forEach((stu, idx) => {
    studentMap[stu.id] = {
      no: idx + 1,
      id: stu.id,
      name: stu.name,
      gender: stu.gender || 'Unknown',
      dob: stu.dob || '',
      present: 0,
      absent: 0,
      permission: 0,
      late: 0,
      totalSessions: 0
    };
  });

  const subjectStats = {};
  let totalScheduledSessions = 0;
  let totalRecordedSessions = 0;
  let totalMissedSessions = 0;

  daysList.forEach(day => {
    totalScheduledSessions += day.totalSessions;
    totalRecordedSessions += day.recordedCount;
    totalMissedSessions += day.missedCount;

    day.scheduledSessions.forEach(session => {
      if (!subjectStats[session.subject]) {
        subjectStats[session.subject] = {
          subject: session.subject,
          code: session.code,
          scheduled: 0,
          recorded: 0,
          missed: 0,
          absent: 0,
          permission: 0,
          late: 0
        };
      }
      subjectStats[session.subject].scheduled++;
      if (session.isRecorded) {
        subjectStats[session.subject].recorded++;
      } else {
        subjectStats[session.subject].missed++;
      }

      if (session.isRecorded && session.records) {
        session.records.forEach(rec => {
          const st = studentMap[rec.id];
          if (st) {
            st.totalSessions++;
            if (rec.status === 'Absent' || rec.code === 'A') {
              st.absent++;
              subjectStats[session.subject].absent++;
            } else if (rec.status === 'Permission' || rec.code === 'P') {
              st.permission++;
              subjectStats[session.subject].permission++;
            } else if (rec.status === 'Late' || rec.code === 'L') {
              st.late++;
              subjectStats[session.subject].late++;
            } else {
              st.present++;
            }
          }
        });
      }
    });
  });

  const studentSummaries = Object.values(studentMap).map(st => {
    const rate = st.totalSessions > 0
      ? ((st.present / st.totalSessions) * 100).toFixed(1)
      : '100.0';

    let rating = 'ល្អណាស់ (Excellent)';
    if (st.absent >= 5) rating = 'ត្រូវកែលម្អ (Needs Attention)';
    else if (st.absent >= 2 || st.permission >= 4) rating = 'មធ្យម (Average)';
    else if (st.late >= 3) rating = 'ល្អបង្គួរ (Good)';

    return {
      ...st,
      attendanceRate: parseFloat(rate),
      rating
    };
  });

  return {
    startDateStr,
    endDateStr,
    totalSchoolDays: daysList.length,
    daysList,
    totalScheduledSessions,
    totalRecordedSessions,
    totalMissedSessions,
    studentSummaries,
    subjectStats: Object.values(subjectStats)
  };
}

module.exports = {
  getStudents,
  saveStudents,
  addStudent,
  updateStudent,
  removeStudent,
  findStudentMatch,
  searchStudents,
  parseAttendanceInput,
  getDailyRecords,
  saveDailyRecords,
  recordAttendance,
  getDayAttendanceStatus,
  getRangeAttendanceStatus
};
