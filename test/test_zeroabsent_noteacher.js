const assert = require('assert');
const dayjs = require('dayjs');
const {
  recordAttendance,
  getDayAttendanceStatus,
  getRangeAttendanceStatus,
  getStudents
} = require('../src/attendance');
const {
  generateFullDayExcel,
  generateRangeExcelReport
} = require('../src/excelReport');
const { DAYS, KHMER_DAYS, getTodaySlots } = require('../src/timetable');

async function testAll() {
  console.log('🧪 Starting Zero Absent & No Teacher Attendance Tests...\n');

  const testDate = '2026-11-09'; // A Monday
  const timetableHelper = { DAYS, KHMER_DAYS, getTodaySlots };

  // 1. Test record attendance: Zero Absent
  console.log('--- 1. Testing recordAttendance with isZeroAbsent: true ---');
  const zeroResult = await recordAttendance({
    dateStr: testDate,
    subject: 'គណិតវិទ្យា',
    timeSlot: '07:00 - 08:00',
    recordedBy: 'ប្រធានថ្នាក់',
    parsedItems: [],
    isZeroAbsent: true,
    noTeacher: false
  });

  assert.strictEqual(zeroResult.summary.absent, 0, 'Absent should be 0');
  assert.strictEqual(zeroResult.summary.permission, 0, 'Permission should be 0');
  assert.strictEqual(zeroResult.summary.late, 0, 'Late should be 0');
  assert.strictEqual(zeroResult.summary.present, zeroResult.summary.total, 'All students present');
  assert.strictEqual(zeroResult.subjectRecord.noTeacher, false, 'noTeacher should be false');
  console.log('✅ Zero Absent recorded successfully: present =', zeroResult.summary.present, ', absent =', zeroResult.summary.absent);

  // 2. Test record attendance: No Teacher
  console.log('\n--- 2. Testing recordAttendance with noTeacher: true ---');
  const noTeacherResult = await recordAttendance({
    dateStr: testDate,
    subject: 'ភាសាខ្មែរ',
    timeSlot: '08:00 - 09:00',
    recordedBy: 'ប្រធានថ្នាក់',
    parsedItems: [],
    isZeroAbsent: true,
    noTeacher: true
  });

  assert.strictEqual(noTeacherResult.summary.absent, 0, 'Absent should be 0');
  assert.strictEqual(noTeacherResult.summary.permission, 0, 'Permission should be 0');
  assert.strictEqual(noTeacherResult.summary.late, 0, 'Late should be 0');
  assert.strictEqual(noTeacherResult.summary.noTeacher, true, 'noTeacher flag should be true');
  assert.strictEqual(noTeacherResult.subjectRecord.noTeacher, true, 'Record noTeacher flag should be true');
  console.log('✅ No Teacher recorded successfully: absent =', noTeacherResult.summary.absent, ', noTeacher =', noTeacherResult.summary.noTeacher);

  // 3. Test getDayAttendanceStatus
  console.log('\n--- 3. Testing getDayAttendanceStatus ---');
  const dayStatus = await getDayAttendanceStatus(testDate, timetableHelper);
  const ntSession = dayStatus.scheduledSessions.find(s => s.noTeacher);
  assert(ntSession, 'Should have at least one session with noTeacher = true');
  console.log('✅ Found session with noTeacher in dayStatus:', ntSession.timeSlot, ntSession.subject);

  // 4. Test generateFullDayExcel
  console.log('\n--- 4. Testing generateFullDayExcel ---');
  const fullDayExcel = await generateFullDayExcel(dayStatus);
  console.log('✅ Full Day Excel created:', fullDayExcel.filename);

  // 5. Test getRangeAttendanceStatus (Weekly / Monthly range)
  console.log('\n--- 5. Testing getRangeAttendanceStatus ---');
  const rangeData = await getRangeAttendanceStatus('2026-11-09', '2026-11-14', timetableHelper);
  console.log('Total scheduled sessions:', rangeData.totalScheduledSessions);
  console.log('Total recorded sessions:', rangeData.totalRecordedSessions);
  console.log('Total no-teacher sessions:', rangeData.totalNoTeacherSessions);
  assert(rangeData.totalNoTeacherSessions >= 1, 'Should have recorded at least 1 no-teacher session');

  // Verify that students have 0 absences for that no teacher session
  const sampleStudent = rangeData.studentSummaries[0];
  console.log('Sample student:', sampleStudent.name, '- absent:', sampleStudent.absent, '- present:', sampleStudent.present);
  assert.strictEqual(sampleStudent.absent, 0, 'Sample student absent count should be 0');

  // 6. Test generateRangeExcelReport for Weekly
  console.log('\n--- 6. Testing generateRangeExcelReport (Weekly) ---');
  const weeklyExcel = await generateRangeExcelReport({
    ...rangeData,
    customTitle: 'របាយការណ៍វត្តមានសិស្សប្រចាំសប្តាហ៍',
    customFileName: 'Attendance_Weekly_2026-11-09_to_2026-11-14.xlsx'
  });
  console.log('✅ Weekly Excel created:', weeklyExcel.filename);

  // 7. Test generateRangeExcelReport for Monthly
  console.log('\n--- 7. Testing generateRangeExcelReport (Monthly) ---');
  const monthlyExcel = await generateRangeExcelReport({
    ...rangeData,
    customTitle: 'របាយការណ៍វត្តមានសិស្សប្រចាំខែវិច្ឆិកា 2026',
    customFileName: 'Attendance_Monthly_2026-11-01_to_2026-11-30.xlsx'
  });
  console.log('✅ Monthly Excel created:', monthlyExcel.filename);

  console.log('\n🎉 ALL TESTS PASSED SUCCESSFULLY! 🎉');
}

testAll().catch(err => {
  console.error('❌ Test failed:', err);
  process.exit(1);
});
