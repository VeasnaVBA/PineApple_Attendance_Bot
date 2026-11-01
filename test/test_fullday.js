const { getDayAttendanceStatus } = require('../src/attendance');
const { generateFullDayExcel } = require('../src/excelReport');
const { DAYS, KHMER_DAYS, getTodaySlots } = require('../src/timetable');

async function testFullDay() {
  console.log('Testing Saturday 2026-10-31...');
  const dayStatus = await getDayAttendanceStatus('2026-10-31', { DAYS, KHMER_DAYS, getTodaySlots });
  console.log('Day Status:', {
    dateStr: dayStatus.dateStr,
    dayName: dayStatus.dayName,
    hasTimetable: dayStatus.hasTimetable,
    totalSessions: dayStatus.totalSessions,
    recordedCount: dayStatus.recordedCount,
    missedCount: dayStatus.missedCount
  });

  const { filePath, filename } = await generateFullDayExcel(dayStatus);
  console.log('✅ Generated Full Day Excel:', filePath);
  console.log('Filename:', filename);
}

testFullDay().catch(console.error);
