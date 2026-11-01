const { parseAttendanceInput, recordAttendance, getStudents } = require('../src/attendance');
const { generateAttendanceExcel } = require('../src/excelReport');
const { getCurrentSlot, getTodaySlots } = require('../src/timetable');
const dayjs = require('dayjs');

async function runTest() {
  console.log('--- 1. Testing Timetable ---');
  const now = dayjs();
  console.log('Today slots:', getTodaySlots(now));
  console.log('Current / Nearest slot:', getCurrentSlot(now));

  console.log('\n--- 2. Testing Message Parsing ---');
  const testMessage = `
    Dara A
    Keo Sopheap P (Sick leave)
    Meas Sokha L Late 10 minutes
  `;
  const parsed = parseAttendanceInput(testMessage);
  console.log('Parsed items:', parsed.parsedItems);

  console.log('\n--- 3. Testing Attendance Recording ---');
  const dateStr = now.format('YYYY-MM-DD');
  const recordResult = recordAttendance({
    dateStr,
    subject: 'Mathematics',
    timeSlot: '07:30 - 09:00',
    recordedBy: 'Class Monitor',
    parsedItems: parsed.parsedItems
  });

  console.log('Summary:', recordResult.summary);
  console.log('Total students marked:', recordResult.subjectRecord.records.length);

  console.log('\n--- 4. Testing Excel Generation ---');
  const { filePath, filename } = await generateAttendanceExcel({
    dateStr,
    subject: recordResult.subjectRecord.subject,
    timeSlot: recordResult.subjectRecord.timeSlot,
    summary: recordResult.summary,
    records: recordResult.subjectRecord.records,
    recordedBy: recordResult.subjectRecord.recordedBy
  });

  console.log('✅ Excel generated successfully:', filePath);
  console.log('Filename:', filename);
}

runTest().catch(console.error);
