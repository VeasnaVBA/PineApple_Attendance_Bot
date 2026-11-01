/**
 * =========================================================================
 * PineApple Attendance Bot - Google Apps Script Database Engine
 * =========================================================================
 * របៀបដំឡើង (Setup Instructions):
 * 1. ចូលទៅកាន់ Google Drive (https://drive.google.com)
 * 2. បង្កើត Folder មួយឈ្មោះថា "PineApple_Attendance"
 * 3. ចូលទៅក្នុង Folder នោះ រួចបង្កើត Google Sheet មួយឈ្មោះថា "PineApple_Attendance_DB"
 * 4. នៅលើ menu ខាងលើ ចុច: Extensions (ផ្នែកបន្ថែម) -> Apps Script
 * 5. លុបកូដចាស់ៗចោល រួច Copy កូដទាំងអស់ក្នុងឯកសារនេះទៅ Paste ចូល
 * 6. ចុច Save (រូបថាស)
 * 7. ចុចប៊ូតុង "Deploy" (ពណ៌ខៀវខាងលើស្តាំ) -> ជ្រើសរើស "New deployment"
 * 8. ចុចរូបកង់ធ្មេញ (Select type) -> ជ្រើសរើស "Web app"
 *    - Description: PineApple Attendance API
 *    - Execute as: Me (អាសយដ្ឋានអ៊ីមែលរបស់អ្នក)
 *    - Who has access: Anyone (អ្នកណាក៏ដោយ)
 * 9. ចុច "Deploy" -> Authorize access (អនុញ្ញាត)
 * 10. Copy យក "Web app URL" (មានទម្រង់ https://script.google.com/macros/s/.../exec)
 *     រួចយកទៅដាក់ក្នុង .env ឬ Vercel Environment Variables:
 *     GOOGLE_SHEET_API_URL=https://script.google.com/macros/s/.../exec
 * =========================================================================
 */

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    var params = e && e.parameter ? e.parameter : {};
    var postData = {};
    if (e && e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (err) {
        postData = {};
      }
    }

    var action = params.action || postData.action || 'ping';
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var result = { status: 'success', action: action };

    initSheets(ss);

    if (action === 'ping') {
      result.message = 'PineApple Attendance Database is ready!';
    } else if (action === 'getStudents') {
      result.students = getStudentsData(ss);
    } else if (action === 'saveStudents') {
      saveStudentsData(ss, postData.students || []);
      result.message = 'Students saved successfully';
    } else if (action === 'recordAttendance') {
      recordAttendanceData(ss, postData);
      result.message = 'Attendance recorded successfully';
    } else if (action === 'getAttendance') {
      result.records = getAttendanceData(ss, params.date || postData.date, params.startDate || postData.startDate, params.endDate || postData.endDate);
    } else if (action === 'getSettings') {
      result.settings = getSettingsData(ss);
    } else if (action === 'saveSettings') {
      saveSettingsData(ss, postData.settings || {});
      result.message = 'Settings saved successfully';
    } else if (action === 'getSubjects') {
      result.subjects = getSubjectsData(ss);
    } else if (action === 'saveSubjects') {
      saveSubjectsData(ss, postData.subjects || {});
      result.message = 'Subjects saved successfully';
    } else if (action === 'clearAttendance') {
      clearAttendanceData(ss);
      result.message = 'Attendance data cleared successfully';
    } else {
      result = { status: 'error', message: 'Unknown action: ' + action };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function initSheets(ss) {
  var sheetNames = ['Students', 'Attendance', 'Settings', 'Subjects'];
  sheetNames.forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (!s) {
      s = ss.insertSheet(name);
      if (name === 'Students') {
        s.appendRow(['No', 'ID', 'Name', 'Gender', 'DOB']);
        s.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      } else if (name === 'Attendance') {
        s.appendRow(['Date', 'TimeSlot', 'Subject', 'StudentID', 'StudentName', 'Status', 'Code', 'Note', 'RecordedBy', 'Timestamp']);
        s.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      } else if (name === 'Settings') {
        s.appendRow(['Key', 'Value']);
        s.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      } else if (name === 'Subjects') {
        s.appendRow(['Code', 'Name']);
        s.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
      }
    }
  });
}

function getStudentsData(ss) {
  var s = ss.getSheetByName('Students');
  var data = s.getDataRange().getValues();
  var students = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][1]) {
      students.push({
        no: Number(data[i][0]) || i,
        id: String(data[i][1]).padStart(2, '0'),
        name: String(data[i][2]),
        gender: String(data[i][3]),
        dob: String(data[i][4])
      });
    }
  }
  return students;
}

function saveStudentsData(ss, students) {
  var s = ss.getSheetByName('Students');
  s.clearContents();
  s.appendRow(['No', 'ID', 'Name', 'Gender', 'DOB']);
  s.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  var rows = [];
  students.forEach(function(st, idx) {
    rows.push([idx + 1, st.id, st.name, st.gender, st.dob || '']);
  });
  if (rows.length > 0) {
    s.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

function formatDateIso(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, 'Asia/Phnom_Penh', 'yyyy-MM-dd');
  }
  var str = String(val).replace(/\s*\(.*\)/, '').trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.substring(0, 10);
  try {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, 'Asia/Phnom_Penh', 'yyyy-MM-dd');
    }
  } catch (e) {}
  return String(val).substring(0, 10);
}

function recordAttendanceData(ss, payload) {
  var s = ss.getSheetByName('Attendance');
  var dateStr = payload.dateStr;
  var timeSlot = String(payload.timeSlot || '').trim();
  var subject = String(payload.subject || '').trim();
  var recordedBy = payload.recordedBy;
  var fullRecords = payload.fullRecords || [];
  var timestamp = new Date().toISOString();

  // Delete existing records for same date + timeSlot + subject to avoid duplicate overwrites
  var data = s.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var rowDate = formatDateIso(data[i][0]);
    var rowTime = String(data[i][1]).trim();
    var rowSubj = String(data[i][2]).trim();
    if (rowDate === dateStr && rowTime === timeSlot && rowSubj === subject) {
      s.deleteRow(i + 1);
    }
  }

  // Append new records
  var newRows = [];
  fullRecords.forEach(function(rec) {
    newRows.push([
      dateStr,
      timeSlot,
      subject,
      rec.id,
      rec.name,
      rec.status,
      rec.code,
      rec.note || '',
      recordedBy,
      timestamp
    ]);
  });

  if (newRows.length > 0) {
    s.getRange(s.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
  }
}

function getAttendanceData(ss, singleDate, startDate, endDate) {
  var s = ss.getSheetByName('Attendance');
  var data = s.getDataRange().getValues();
  var records = [];

  for (var i = 1; i < data.length; i++) {
    var rowDate = formatDateIso(data[i][0]);
    var match = false;
    if (singleDate) {
      if (rowDate === singleDate) match = true;
    } else if (startDate && endDate) {
      if (rowDate >= startDate && rowDate <= endDate) match = true;
    } else {
      match = true;
    }

    if (match && data[i][1] && data[i][3]) {
      records.push({
        date: rowDate,
        timeSlot: String(data[i][1]).trim(),
        subject: String(data[i][2]).trim(),
        id: String(data[i][3]),
        name: String(data[i][4]),
        status: String(data[i][5]),
        code: String(data[i][6]),
        note: String(data[i][7]),
        recordedBy: String(data[i][8]),
        timestamp: String(data[i][9])
      });
    }
  }
  return records;
}

function getSettingsData(ss) {
  var s = ss.getSheetByName('Settings');
  var data = s.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      settings[String(data[i][0])] = String(data[i][1]);
    }
  }
  return settings;
}

function saveSettingsData(ss, newSettings) {
  var s = ss.getSheetByName('Settings');
  var current = getSettingsData(ss);
  for (var k in newSettings) {
    current[k] = newSettings[k];
  }
  s.clearContents();
  s.appendRow(['Key', 'Value']);
  s.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  var rows = [];
  for (var key in current) {
    rows.push([key, current[key]]);
  }
  if (rows.length > 0) {
    s.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function getSubjectsData(ss) {
  var s = ss.getSheetByName('Subjects');
  var data = s.getDataRange().getValues();
  var subjects = {};
  for (var i = 1; i < data.length; i++) {
    if (data[i][0]) {
      subjects[String(data[i][0])] = String(data[i][1]);
    }
  }
  return subjects;
}

function saveSubjectsData(ss, subjects) {
  var s = ss.getSheetByName('Subjects');
  s.clearContents();
  s.appendRow(['Code', 'Name']);
  s.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  var rows = [];
  for (var code in subjects) {
    rows.push([code, subjects[code]]);
  }
  if (rows.length > 0) {
    s.getRange(2, 1, rows.length, 2).setValues(rows);
  }
}

function clearAttendanceData(ss) {
  var s = ss.getSheetByName('Attendance');
  if (s) {
    s.clearContents();
    s.appendRow(['Date', 'TimeSlot', 'Subject', 'StudentID', 'StudentName', 'Status', 'Code', 'Note', 'RecordedBy', 'Timestamp']);
    s.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  }
}

