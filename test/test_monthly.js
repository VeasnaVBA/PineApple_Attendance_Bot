const { getRangeAttendanceStatus } = require('../src/attendance');
const { generateRangeExcelReport } = require('../src/excelReport');
const { DAYS, KHMER_DAYS, getTodaySlots } = require('../src/timetable');

async function testMonthly() {
  console.log('Testing Monthly Report from 2026-10-01 to 2026-10-31...');
  const rangeData = await getRangeAttendanceStatus('2026-10-01', '2026-10-31', { DAYS, KHMER_DAYS, getTodaySlots });
  console.log('Total school days:', rangeData.totalSchoolDays);
  console.log('Total scheduled sessions:', rangeData.totalScheduledSessions);
  console.log('Total students:', rangeData.studentSummaries.length);

  const { filePath, filename } = await generateRangeExcelReport(rangeData);
  console.log('✅ Generated Monthly Excel successfully:', filePath);
  console.log('Filename:', filename);
}

testMonthly().catch(console.error);
