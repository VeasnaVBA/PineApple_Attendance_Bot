const ExcelJS = require('exceljs');
const path = require('path');
const dayjs = require('dayjs');
const { TEMP_DIR, CLASS_NAME, TEACHER_NAME } = require('./config');

/**
 * Generates an Excel report for a SINGLE subject session.
 */
async function generateAttendanceExcel({ dateStr, subject, timeSlot, summary, records, recordedBy }) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PineApple Attendance Bot';
  workbook.lastModifiedBy = TEACHER_NAME;
  workbook.created = new Date();
  workbook.modified = new Date();

  const safeSubject = (subject || 'Attendance').replace(/[\/\\?*\[\]:]/g, '_').substring(0, 30);
  const worksheet = workbook.addWorksheet(safeSubject, {
    views: [{ showGridLines: true }]
  });

  const THEME_HEADER_FILL = '1F4E78';
  const THEME_HEADER_TEXT = 'FFFFFF';
  const BORDER_COLOR = 'D9D9D9';
  const FONT_FAMILY = 'Calibri';

  // Title Row
  worksheet.mergeCells('A1:G1');
  const titleCell = worksheet.getCell('A1');
  titleCell.value = 'បញ្ជីវត្តមានសិស្ស (STUDENT ATTENDANCE REPORT)';
  titleCell.font = { name: FONT_FAMILY, size: 15, bold: true, color: { argb: 'FFFFFF' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
  worksheet.getRow(1).height = 36;

  // Metadata block
  const meta = [
    ['គ្រឹះស្ថាន/ថ្នាក់:', CLASS_NAME, 'កាលបរិច្ឆេទ:', dayjs(dateStr).format('dddd, DD MMMM YYYY')],
    ['មុខវិជ្ជា:', subject, 'ម៉ោងសិក្សា:', timeSlot || 'Regular Session'],
    ['គ្រូបន្ទុកថ្នាក់:', TEACHER_NAME, 'អ្នកកត់ត្រា:', recordedBy || 'ប្រធានថ្នាក់']
  ];

  meta.forEach((row, idx) => {
    const rowNum = 3 + idx;
    worksheet.getCell(`A${rowNum}`).value = row[0];
    worksheet.getCell(`A${rowNum}`).font = { name: FONT_FAMILY, bold: true, color: { argb: '333333' } };
    worksheet.getCell(`B${rowNum}`).value = row[1];
    worksheet.getCell(`B${rowNum}`).font = { name: FONT_FAMILY, bold: false };
    worksheet.mergeCells(`B${rowNum}:C${rowNum}`);

    worksheet.getCell(`E${rowNum}`).value = row[2];
    worksheet.getCell(`E${rowNum}`).font = { name: FONT_FAMILY, bold: true, color: { argb: '333333' } };
    worksheet.getCell(`F${rowNum}`).value = row[3];
    worksheet.getCell(`F${rowNum}`).font = { name: FONT_FAMILY, bold: false };
    worksheet.mergeCells(`F${rowNum}:G${rowNum}`);
  });

  // Summary KPI Cards
  const total = summary.total || records.length || 0;
  const present = summary.present || 0;
  const absent = summary.absent || 0;
  const permission = summary.permission || 0;
  const late = summary.late || 0;
  const attendanceRate = total > 0 ? ((present / total) * 100).toFixed(1) + '%' : '0%';

  const kpis = [
    { label: 'សរុប (Total)', val: total, bg: 'E9ECEF', fg: '212529', col: 'A' },
    { label: 'វត្តមាន (Present)', val: present, bg: 'D4EDDA', fg: '155724', col: 'B' },
    { label: 'អវត្តមាន (Absent - A)', val: absent, bg: 'F8D7DA', fg: '721C24', col: 'C' },
    { label: 'ច្បាប់ (Permission - P)', val: permission, bg: 'FFF3CD', fg: '856404', col: 'D' },
    { label: 'យឺត (Late - L)', val: late, bg: 'E2E3E5', fg: '383D41', col: 'E' },
    { label: 'ភាគរយវត្តមាន (%)', val: attendanceRate, bg: 'CCE5FF', fg: '004085', col: 'F' }
  ];

  worksheet.getRow(7).height = 18;
  worksheet.getRow(8).height = 24;

  kpis.forEach(kpi => {
    const headerCell = worksheet.getCell(`${kpi.col}7`);
    headerCell.value = kpi.label;
    headerCell.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: kpi.fg } };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
    headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };

    const valCell = worksheet.getCell(`${kpi.col}8`);
    valCell.value = kpi.val;
    valCell.font = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: kpi.fg } };
    valCell.alignment = { horizontal: 'center', vertical: 'middle' };
    valCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };
  });

  // Table Headers
  const headers = ['ល.រ', 'អត្តលេខ', 'នាមត្រកូល-នាមខ្លួន', 'ភេទ', 'ថ្ងៃខែឆ្នាំកំណើត', 'ស្ថានភាព', 'សម្គាល់ / Note'];
  const tableHeaderRow = worksheet.getRow(10);
  tableHeaderRow.height = 26;

  headers.forEach((h, i) => {
    const cell = tableHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: FONT_FAMILY, size: 11, bold: true, color: { argb: THEME_HEADER_TEXT } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
    cell.alignment = { horizontal: i === 2 || i === 6 ? 'left' : 'center', vertical: 'middle' };
    cell.border = {
      top: { style: 'thin', color: { argb: '000000' } },
      bottom: { style: 'medium', color: { argb: '000000' } }
    };
  });

  // Table Rows
  let curRow = 11;
  for (const stu of records) {
    const row = worksheet.getRow(curRow);
    row.height = 22;

    row.getCell(1).value = stu.no;
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(2).value = stu.id;
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(3).value = stu.name;
    row.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };

    row.getCell(4).value = stu.gender || '-';
    row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(5).value = stu.dob || '-';
    row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    const statusCell = row.getCell(6);
    let statusText = 'មានវត្តមាន (Present)';
    if (stu.status === 'Absent') statusText = 'អវត្តមាន (Absent)';
    else if (stu.status === 'Permission') statusText = 'ច្បាប់ (Permission)';
    else if (stu.status === 'Late') statusText = 'យឺត (Late)';

    statusCell.value = statusText;
    statusCell.alignment = { horizontal: 'center', vertical: 'middle' };
    statusCell.font = { bold: true };

    if (stu.status === 'Present') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D4EDDA' } };
      statusCell.font = { bold: true, color: { argb: '155724' } };
    } else if (stu.status === 'Absent') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
      statusCell.font = { bold: true, color: { argb: '721C24' } };
    } else if (stu.status === 'Permission') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
      statusCell.font = { bold: true, color: { argb: '856404' } };
    } else if (stu.status === 'Late') {
      statusCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2E3E5' } };
      statusCell.font = { bold: true, color: { argb: '383D41' } };
    }

    const noteCell = row.getCell(7);
    noteCell.value = stu.note || '';
    noteCell.alignment = { horizontal: 'left', vertical: 'middle' };

    for (let c = 1; c <= 7; c++) {
      row.getCell(c).border = {
        top: { style: 'thin', color: { argb: BORDER_COLOR } },
        bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
        left: { style: 'thin', color: { argb: BORDER_COLOR } },
        right: { style: 'thin', color: { argb: BORDER_COLOR } }
      };
    }
    curRow++;
  }

  worksheet.columns = [
    { width: 6 },
    { width: 10 },
    { width: 28 },
    { width: 10 },
    { width: 16 },
    { width: 22 },
    { width: 28 }
  ];

  const safeSubjFile = (subject || 'Subject').replace(/[^a-zA-Z0-9_\-\u1780-\u17FF]/g, '_');
  const filename = `Attendance_${safeSubjFile}_${dateStr}.xlsx`;
  const filePath = path.join(TEMP_DIR, filename);

  await workbook.xlsx.writeFile(filePath);
  return { filePath, filename };
}

/**
 * Generates a COMPREHENSIVE FULL-DAY Excel report containing ALL subjects of that day!
 * Flags sessions that were recorded vs. forgotten by the class monitor.
 */
async function generateFullDayExcel(dayStatus) {
  const { dateStr, khmerDay, scheduledSessions, totalSessions, recordedCount, missedCount, students } = dayStatus;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PineApple Attendance Bot';
  workbook.lastModifiedBy = TEACHER_NAME;
  workbook.created = new Date();
  workbook.modified = new Date();

  const THEME_HEADER_FILL = '1F4E78';
  const BORDER_COLOR = 'D9D9D9';
  const FONT_FAMILY = 'Calibri';

  // 1. MASTER SUMMARY SHEET
  const matrixSheet = workbook.addWorksheet('តារាងរួមប្រចាំថ្ងៃ', {
    views: [{ showGridLines: true }]
  });

  const totalCols = 5 + scheduledSessions.length + 3; // Info cols (5) + Session cols + Summary cols (3)

  // Title
  matrixSheet.mergeCells(1, 1, 1, totalCols);
  const titleCell = matrixSheet.getCell(1, 1);
  titleCell.value = `របាយការណ៍វត្តមានសិស្សពេញមួយថ្ងៃ (FULL DAY ATTENDANCE REPORT)`;
  titleCell.font = { name: FONT_FAMILY, size: 15, bold: true, color: { argb: 'FFFFFF' } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
  matrixSheet.getRow(1).height = 36;

  // Metadata block
  matrixSheet.getCell('A3').value = 'គ្រឹះស្ថាន/ថ្នាក់:';
  matrixSheet.getCell('A3').font = { bold: true };
  matrixSheet.getCell('B3').value = CLASS_NAME;

  matrixSheet.getCell('D3').value = 'កាលបរិច្ឆេទ:';
  matrixSheet.getCell('D3').font = { bold: true };
  matrixSheet.getCell('E3').value = `ថ្ងៃ${khmerDay}, ${dayjs(dateStr).format('DD MMMM YYYY')}`;

  matrixSheet.getCell('A4').value = 'គ្រូបន្ទុកថ្នាក់:';
  matrixSheet.getCell('A4').font = { bold: true };
  matrixSheet.getCell('B4').value = TEACHER_NAME;

  matrixSheet.getCell('D4').value = 'ស្ថានភាពកត់ត្រា:';
  matrixSheet.getCell('D4').font = { bold: true };
  const statusStr = missedCount === 0
    ? `✅ កត់គ្រប់ ${totalSessions} ម៉ោង`
    : `⚠️ កត់បាន ${recordedCount}/${totalSessions} ម៉ោង (ភ្លេចកត់ ${missedCount} ម៉ោង)`;
  matrixSheet.getCell('E4').value = statusStr;
  matrixSheet.getCell('E4').font = { bold: true, color: { argb: missedCount > 0 ? 'C00000' : '008000' } };

  // Session Banner Header (Row 6)
  matrixSheet.getRow(6).height = 20;
  matrixSheet.getCell('A6').value = 'កាលវិភាគម៉ោងសិក្សាប្រចាំថ្ងៃ:';
  matrixSheet.getCell('A6').font = { bold: true, color: { argb: '1F4E78' } };

  // Session status tags (Row 7)
  matrixSheet.getRow(7).height = 24;
  scheduledSessions.forEach((s, idx) => {
    const colIdx = 6 + idx;
    const cell = matrixSheet.getCell(7, colIdx);
    if (s.isRecorded) {
      cell.value = `✅ បានកត់ (${s.recordedBy})`;
      cell.font = { size: 9, bold: true, color: { argb: '155724' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D4EDDA' } };
    } else {
      cell.value = `⚠️ ភ្លេចកត់ត្រា!`;
      cell.font = { size: 9, bold: true, color: { argb: '721C24' } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
    }
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Table Headers (Row 8)
  const headerRow = matrixSheet.getRow(8);
  headerRow.height = 36;

  const fixedHeaders = ['ល.រ', 'អត្តលេខ', 'នាមត្រកូល-នាមខ្លួន', 'ភេទ', 'ថ្ងៃកំណើត'];
  fixedHeaders.forEach((h, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // Session columns headers
  scheduledSessions.forEach((s, i) => {
    const cell = headerRow.getCell(6 + i);
    cell.value = `${s.timeSlot}\n${s.code || s.subject}`;
    cell.font = { bold: true, color: { argb: 'FFFFFF' }, size: 9 };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.isRecorded ? '1F4E78' : '808080' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // Summary headers
  const summaryHeaders = ['អវត្តមាន (A)', 'ច្បាប់ (P)', 'យឺត (L)'];
  summaryHeaders.forEach((h, i) => {
    const cell = headerRow.getCell(6 + scheduledSessions.length + i);
    cell.value = h;
    cell.font = { bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '002060' } };
    cell.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  // Populate Student Rows (Row 9 onwards)
  let curRow = 9;
  for (const stu of students) {
    const row = matrixSheet.getRow(curRow);
    row.height = 20;

    row.getCell(1).value = parseInt(stu.id, 10) || curRow - 8;
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(2).value = stu.id;
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(3).value = stu.name;
    row.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };

    row.getCell(4).value = stu.gender || '-';
    row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(5).value = stu.dob || '-';
    row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    let totalA = 0;
    let totalP = 0;
    let totalL = 0;

    // Session Status cells
    scheduledSessions.forEach((s, i) => {
      const cell = row.getCell(6 + i);
      if (!s.isRecorded) {
        // Class monitor forgot to record!
        cell.value = '⚠️ ភ្លេចកត់';
        cell.font = { size: 9, italic: true, color: { argb: '6C757D' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F0F0F0' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        const sId = String(stu.id || '').padStart(2, '0');
        const stuRec = s.records.find(r => {
          const rId = String(r.id || '').padStart(2, '0');
          return (rId && rId === sId) || (r.name && r.name === stu.name);
        });
        const code = stuRec ? stuRec.code : '✓';
        const status = stuRec ? stuRec.status : 'Present';

        if (status === 'Absent' || code === 'A') {
          cell.value = 'A';
          cell.font = { bold: true, color: { argb: '721C24' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
          totalA++;
        } else if (status === 'Permission' || code === 'P') {
          cell.value = 'P';
          cell.font = { bold: true, color: { argb: '856404' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
          totalP++;
        } else if (status === 'Late' || code === 'L') {
          cell.value = 'L';
          cell.font = { bold: true, color: { argb: '383D41' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2E3E5' } };
          totalL++;
        } else {
          // Present
          cell.value = '✓';
          cell.font = { bold: true, color: { argb: '155724' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D4EDDA' } };
        }
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });

    // Summary Totals
    const aCell = row.getCell(6 + scheduledSessions.length);
    aCell.value = totalA;
    aCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (totalA > 0) {
      aCell.font = { bold: true, color: { argb: '721C24' } };
      aCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
    }

    const pCell = row.getCell(7 + scheduledSessions.length);
    pCell.value = totalP;
    pCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (totalP > 0) {
      pCell.font = { bold: true, color: { argb: '856404' } };
      pCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
    }

    const lCell = row.getCell(8 + scheduledSessions.length);
    lCell.value = totalL;
    lCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (totalL > 0) {
      lCell.font = { bold: true, color: { argb: '383D41' } };
      lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2E3E5' } };
    }

    // Cell borders
    for (let c = 1; c <= totalCols; c++) {
      row.getCell(c).border = {
        top: { style: 'thin', color: { argb: BORDER_COLOR } },
        bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
        left: { style: 'thin', color: { argb: BORDER_COLOR } },
        right: { style: 'thin', color: { argb: BORDER_COLOR } }
      };
    }
    curRow++;
  }

  // Adjust Matrix Column Widths
  matrixSheet.columns = [
    { width: 6 },
    { width: 10 },
    { width: 26 },
    { width: 8 },
    { width: 14 },
    ...scheduledSessions.map(() => ({ width: 16 })),
    { width: 14 },
    { width: 14 },
    { width: 14 }
  ];

  // 2. DETAIL SHEETS PER SESSION
  scheduledSessions.forEach((s) => {
    const safeSheetTitle = `${s.timeSlot.replace(/[\s\:]/g, '')}_${(s.code || s.subject).slice(0, 10)}`;
    const detailSheet = workbook.addWorksheet(safeSheetTitle, { views: [{ showGridLines: true }] });

    detailSheet.mergeCells('A1:G1');
    const dTitle = detailSheet.getCell('A1');
    dTitle.value = `វត្តមានម៉ោង: ${s.timeSlot} - ${s.subject}`;
    dTitle.font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
    dTitle.alignment = { horizontal: 'center', vertical: 'middle' };
    dTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: s.isRecorded ? '1F4E78' : 'C00000' } };
    detailSheet.getRow(1).height = 32;

    if (!s.isRecorded) {
      detailSheet.mergeCells('A3:G4');
      const warnCell = detailSheet.getCell('A3');
      warnCell.value = '⚠️ ម៉ោងនេះប្រធានថ្នាក់មិនបានកត់ត្រាវត្តមាននៅឡើយទេ (Class Monitor Forgot to Record!)';
      warnCell.font = { bold: true, size: 13, color: { argb: '721C24' } };
      warnCell.alignment = { horizontal: 'center', vertical: 'middle' };
      warnCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
    } else {
      // Show details
      const headers = ['ល.រ', 'អត្តលេខ', 'នាមត្រកូល-នាមខ្លួន', 'ភេទ', 'ថ្ងៃកំណើត', 'ស្ថានភាព', 'សម្គាល់/មូលហេតុ'];
      const hRow = detailSheet.getRow(3);
      hRow.height = 24;
      headers.forEach((h, idx) => {
        const cell = hRow.getCell(idx + 1);
        cell.value = h;
        cell.font = { bold: true, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F4E78' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      });

      let rIdx = 4;
      s.records.forEach(r => {
        const rRow = detailSheet.getRow(rIdx);
        rRow.height = 20;
        rRow.getCell(1).value = r.no;
        rRow.getCell(2).value = r.id;
        rRow.getCell(3).value = r.name;
        rRow.getCell(4).value = r.gender;
        rRow.getCell(5).value = r.dob || '-';

        const stCell = rRow.getCell(6);
        stCell.value = r.status;
        stCell.alignment = { horizontal: 'center', vertical: 'middle' };
        if (r.status === 'Present') {
          stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D4EDDA' } };
          stCell.font = { bold: true, color: { argb: '155724' } };
        } else if (r.status === 'Absent') {
          stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
          stCell.font = { bold: true, color: { argb: '721C24' } };
        } else if (r.status === 'Permission') {
          stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
          stCell.font = { bold: true, color: { argb: '856404' } };
        } else if (r.status === 'Late') {
          stCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2E3E5' } };
          stCell.font = { bold: true, color: { argb: '383D41' } };
        }

        rRow.getCell(7).value = r.note || '';

        for (let c = 1; c <= 7; c++) {
          rRow.getCell(c).border = {
            top: { style: 'thin', color: { argb: BORDER_COLOR } },
            bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
            left: { style: 'thin', color: { argb: BORDER_COLOR } },
            right: { style: 'thin', color: { argb: BORDER_COLOR } }
          };
        }
        rIdx++;
      });

      detailSheet.columns = [
        { width: 6 },
        { width: 10 },
        { width: 26 },
        { width: 10 },
        { width: 16 },
        { width: 18 },
        { width: 30 }
      ];
    }
  });

  const filename = `Attendance_FullDay_${dayStatus.dayName}_${dateStr}.xlsx`;
  const filePath = path.join(TEMP_DIR, filename);

  await workbook.xlsx.writeFile(filePath);
  return { filePath, filename };
}

/**
 * Generates a MONTHLY / DATE RANGE Excel report
 * (from startDate to endDate)
 */
async function generateRangeExcelReport(rangeData) {
  const {
    startDateStr,
    endDateStr,
    totalSchoolDays,
    daysList,
    totalScheduledSessions,
    totalRecordedSessions,
    totalMissedSessions,
    studentSummaries,
    subjectStats,
    customTitle,
    customFileName
  } = rangeData;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'PineApple Attendance Bot';
  workbook.lastModifiedBy = TEACHER_NAME;
  workbook.created = new Date();
  workbook.modified = new Date();

  const THEME_HEADER_FILL = '1F4E78';
  const BORDER_COLOR = 'D9D9D9';
  const FONT_FAMILY = 'Calibri';

  // --- SHEET 1: សង្ខេបប្រចាំខែ (MONTHLY SUMMARY & RANKING) ---
  const summarySheet = workbook.addWorksheet('សង្ខេបប្រចាំខែ', {
    views: [{ showGridLines: true }]
  });

  // Title
  summarySheet.mergeCells('A1:K1');
  const tCell = summarySheet.getCell('A1');
  tCell.value = customTitle || `របាយការណ៍វត្តមានសិស្សសរុប (ATTENDANCE REPORT)`;
  tCell.font = { name: FONT_FAMILY, size: 15, bold: true, color: { argb: 'FFFFFF' } };
  tCell.alignment = { horizontal: 'center', vertical: 'middle' };
  tCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
  summarySheet.getRow(1).height = 36;

  // Metadata
  summarySheet.getCell('A3').value = 'គ្រឹះស្ថាន/ថ្នាក់:';
  summarySheet.getCell('A3').font = { bold: true };
  summarySheet.getCell('B3').value = CLASS_NAME;

  summarySheet.getCell('D3').value = 'ចន្លោះកាលបរិច្ឆេទ:';
  summarySheet.getCell('D3').font = { bold: true };
  summarySheet.getCell('E3').value = `ពី ${dayjs(startDateStr).format('DD/MM/YYYY')} ដល់ ${dayjs(endDateStr).format('DD/MM/YYYY')}`;

  summarySheet.getCell('A4').value = 'គ្រូបន្ទុកថ្នាក់:';
  summarySheet.getCell('A4').font = { bold: true };
  summarySheet.getCell('B4').value = TEACHER_NAME;

  summarySheet.getCell('D4').value = 'ទិន្នន័យម៉ោងសិក្សា:';
  summarySheet.getCell('D4').font = { bold: true };
  summarySheet.getCell('E4').value = `${totalSchoolDays} ថ្ងៃរៀន | សរុប ${totalScheduledSessions} ម៉ោង (កត់ ${totalRecordedSessions}, ភ្លេច ${totalMissedSessions})`;

  // Summary KPI Cards (Row 6 & 7)
  let totalClassAbsences = 0;
  let totalClassPermissions = 0;
  let totalClassLates = 0;
  studentSummaries.forEach(st => {
    totalClassAbsences += st.absent;
    totalClassPermissions += st.permission;
    totalClassLates += st.late;
  });

  const avgAttendance = studentSummaries.length > 0
    ? (studentSummaries.reduce((acc, st) => acc + st.attendanceRate, 0) / studentSummaries.length).toFixed(1) + '%'
    : '100%';

  const kpis = [
    { label: 'ថ្ងៃរៀនសរុប (Days)', val: totalSchoolDays, bg: 'E9ECEF', fg: '212529', col: 'A' },
    { label: 'ម៉ោងបង្រៀនសរុប', val: totalScheduledSessions, bg: 'CCE5FF', fg: '004085', col: 'B' },
    { label: 'អវត្តមានសរុប (A)', val: totalClassAbsences, bg: 'F8D7DA', fg: '721C24', col: 'C' },
    { label: 'ច្បាប់សរុប (P)', val: totalClassPermissions, bg: 'FFF3CD', fg: '856404', col: 'D' },
    { label: 'យឺតសរុប (L)', val: totalClassLates, bg: 'E2E3E5', fg: '383D41', col: 'E' },
    { label: 'មធ្យមភាគវត្តមាន', val: avgAttendance, bg: 'D4EDDA', fg: '155724', col: 'F' }
  ];

  summarySheet.getRow(6).height = 18;
  summarySheet.getRow(7).height = 24;

  kpis.forEach(kpi => {
    const headerCell = summarySheet.getCell(`${kpi.col}6`);
    headerCell.value = kpi.label;
    headerCell.font = { name: FONT_FAMILY, size: 9, bold: true, color: { argb: kpi.fg } };
    headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
    headerCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };

    const valCell = summarySheet.getCell(`${kpi.col}7`);
    valCell.value = kpi.val;
    valCell.font = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: kpi.fg } };
    valCell.alignment = { horizontal: 'center', vertical: 'middle' };
    valCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: kpi.bg } };
  });

  // Table Headers (Row 9)
  const headers = [
    'ល.រ',
    'អត្តលេខ',
    'គោត្តនាម-នាម',
    'ភេទ',
    'ថ្ងៃកំណើត',
    'វត្តមាន (ម៉ោង)',
    'អវត្តមាន (A)',
    'ច្បាប់ (P)',
    'យឺត (L)',
    'ភាគរយ (%)',
    'ការវាយតម្លៃវិន័យ'
  ];

  const tableHeaderRow = summarySheet.getRow(9);
  tableHeaderRow.height = 28;

  headers.forEach((h, i) => {
    const cell = tableHeaderRow.getCell(i + 1);
    cell.value = h;
    cell.font = { name: FONT_FAMILY, size: 10, bold: true, color: { argb: 'FFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
    cell.alignment = { horizontal: i === 2 || i === 10 ? 'left' : 'center', vertical: 'middle' };
  });

  // Table Rows (Row 10 onwards)
  let curRow = 10;
  for (const st of studentSummaries) {
    const row = summarySheet.getRow(curRow);
    row.height = 20;

    row.getCell(1).value = st.no;
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(2).value = st.id;
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(3).value = st.name;
    row.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };

    row.getCell(4).value = st.gender || '-';
    row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(5).value = st.dob || '-';
    row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    row.getCell(6).value = st.present;
    row.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };

    const aCell = row.getCell(7);
    aCell.value = st.absent;
    aCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (st.absent > 0) {
      aCell.font = { bold: true, color: { argb: '721C24' } };
      aCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
    }

    const pCell = row.getCell(8);
    pCell.value = st.permission;
    pCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (st.permission > 0) {
      pCell.font = { bold: true, color: { argb: '856404' } };
      pCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
    }

    const lCell = row.getCell(9);
    lCell.value = st.late;
    lCell.alignment = { horizontal: 'center', vertical: 'middle' };
    if (st.late > 0) {
      lCell.font = { bold: true, color: { argb: '383D41' } };
      lCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2E3E5' } };
    }

    const rateCell = row.getCell(10);
    rateCell.value = `${st.attendanceRate}%`;
    rateCell.alignment = { horizontal: 'center', vertical: 'middle' };
    rateCell.font = { bold: true };
    if (st.attendanceRate < 90) {
      rateCell.font = { bold: true, color: { argb: '721C24' } };
    } else {
      rateCell.font = { bold: true, color: { argb: '155724' } };
    }

    const evalCell = row.getCell(11);
    evalCell.value = st.rating;
    evalCell.alignment = { horizontal: 'left', vertical: 'middle' };

    for (let c = 1; c <= 11; c++) {
      row.getCell(c).border = {
        top: { style: 'thin', color: { argb: BORDER_COLOR } },
        bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
        left: { style: 'thin', color: { argb: BORDER_COLOR } },
        right: { style: 'thin', color: { argb: BORDER_COLOR } }
      };
    }
    curRow++;
  }

  summarySheet.columns = [
    { width: 6 },
    { width: 10 },
    { width: 26 },
    { width: 8 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 14 },
    { width: 24 }
  ];

  // --- SHEET 2: តារាងវត្តមានតាមថ្ងៃ (DAILY CALENDAR MATRIX) ---
  const calSheet = workbook.addWorksheet('តារាងវត្តមានតាមថ្ងៃ', {
    views: [{ showGridLines: true }]
  });

  const totalCalCols = 5 + daysList.length;

  calSheet.mergeCells(1, 1, 1, totalCalCols);
  const calTitle = calSheet.getCell(1, 1);
  calTitle.value = `តារាងវត្តមានសិស្សតាមថ្ងៃ (DAILY ATTENDANCE LOG)`;
  calTitle.font = { name: FONT_FAMILY, size: 14, bold: true, color: { argb: 'FFFFFF' } };
  calTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  calTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
  calSheet.getRow(1).height = 32;

  // Header row (Row 3)
  const calHeader = calSheet.getRow(3);
  calHeader.height = 30;
  const fixedCal = ['ល.រ', 'អត្តលេខ', 'នាមត្រកូល-នាមខ្លួន', 'ភេទ', 'ថ្ងៃកំណើត'];
  fixedCal.forEach((h, i) => {
    const c = calHeader.getCell(i + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: 'FFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
    c.alignment = { horizontal: 'center', vertical: 'middle' };
  });

  daysList.forEach((d, i) => {
    const c = calHeader.getCell(6 + i);
    c.value = `${dayjs(d.dateStr).format('DD/MM')}\n${d.khmerDay}`;
    c.font = { bold: true, color: { argb: 'FFFFFF' }, size: 9 };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '002060' } };
    c.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  });

  // Populate student rows in Daily Matrix
  let cRowIdx = 4;
  for (const st of studentSummaries) {
    const row = calSheet.getRow(cRowIdx);
    row.height = 20;

    row.getCell(1).value = st.no;
    row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(2).value = st.id;
    row.getCell(2).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(3).value = st.name;
    row.getCell(3).alignment = { horizontal: 'left', vertical: 'middle' };
    row.getCell(4).value = st.gender || '-';
    row.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
    row.getCell(5).value = st.dob || '-';
    row.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };

    daysList.forEach((d, i) => {
      const cell = row.getCell(6 + i);

      let hasAbsent = false;
      let hasPermission = false;
      let hasLate = false;
      let anyRecorded = false;

      d.scheduledSessions.forEach(session => {
        if (session.isRecorded) {
          anyRecorded = true;
          const rec = session.records.find(r => r.id === st.id);
          if (rec) {
            if (rec.status === 'Absent' || rec.code === 'A') hasAbsent = true;
            else if (rec.status === 'Permission' || rec.code === 'P') hasPermission = true;
            else if (rec.status === 'Late' || rec.code === 'L') hasLate = true;
          }
        }
      });

      if (!anyRecorded) {
        cell.value = '-';
        cell.font = { color: { argb: '999999' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (hasAbsent) {
        cell.value = 'A';
        cell.font = { bold: true, color: { argb: '721C24' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (hasPermission) {
        cell.value = 'P';
        cell.font = { bold: true, color: { argb: '856404' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (hasLate) {
        cell.value = 'L';
        cell.font = { bold: true, color: { argb: '383D41' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E2E3E5' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else {
        cell.value = '✓';
        cell.font = { bold: true, color: { argb: '155724' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D4EDDA' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    });

    for (let c = 1; c <= totalCalCols; c++) {
      row.getCell(c).border = {
        top: { style: 'thin', color: { argb: BORDER_COLOR } },
        bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
        left: { style: 'thin', color: { argb: BORDER_COLOR } },
        right: { style: 'thin', color: { argb: BORDER_COLOR } }
      };
    }
    cRowIdx++;
  }

  calSheet.columns = [
    { width: 6 },
    { width: 10 },
    { width: 26 },
    { width: 8 },
    { width: 14 },
    ...daysList.map(() => ({ width: 11 }))
  ];

  // --- SHEET 3: សង្ខេបតាមមុខវិជ្ជា (SUBJECT STATS) ---
  const subjSheet = workbook.addWorksheet('សង្ខេបតាមមុខវិជ្ជា', {
    views: [{ showGridLines: true }]
  });

  subjSheet.mergeCells('A1:G1');
  const sTitle = subjSheet.getCell('A1');
  sTitle.value = 'ស្ថិតិវត្តមានតាមមុខវិជ្ជានីមួយៗ (SUBJECT ATTENDANCE SUMMARY)';
  sTitle.font = { bold: true, size: 14, color: { argb: 'FFFFFF' } };
  sTitle.alignment = { horizontal: 'center', vertical: 'middle' };
  sTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
  subjSheet.getRow(1).height = 32;

  const sHeaders = ['ល.រ', 'មុខវិជ្ជា', 'កូដ', 'ម៉ោងបង្រៀនសរុប', 'បានកត់ត្រា', 'អវត្តមានសរុប (A)', 'ច្បាប់សរុប (P)'];
  const sHRow = subjSheet.getRow(3);
  sHRow.height = 24;
  sHeaders.forEach((h, idx) => {
    const c = sHRow.getCell(idx + 1);
    c.value = h;
    c.font = { bold: true, color: { argb: 'FFFFFF' } };
    c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: THEME_HEADER_FILL } };
    c.alignment = { horizontal: idx === 1 ? 'left' : 'center', vertical: 'middle' };
  });

  let sRowIdx = 4;
  subjectStats.forEach((sb, i) => {
    const r = subjSheet.getRow(sRowIdx);
    r.height = 20;
    r.getCell(1).value = i + 1;
    r.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(2).value = sb.subject;
    r.getCell(3).value = sb.code || '-';
    r.getCell(3).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(4).value = sb.scheduled;
    r.getCell(4).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(5).value = sb.recorded;
    r.getCell(5).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(6).value = sb.absent;
    r.getCell(6).alignment = { horizontal: 'center', vertical: 'middle' };
    r.getCell(7).value = sb.permission;
    r.getCell(7).alignment = { horizontal: 'center', vertical: 'middle' };

    for (let c = 1; c <= 7; c++) {
      r.getCell(c).border = {
        top: { style: 'thin', color: { argb: BORDER_COLOR } },
        bottom: { style: 'thin', color: { argb: BORDER_COLOR } },
        left: { style: 'thin', color: { argb: BORDER_COLOR } },
        right: { style: 'thin', color: { argb: BORDER_COLOR } }
      };
    }
    sRowIdx++;
  });

  subjSheet.columns = [
    { width: 6 },
    { width: 28 },
    { width: 12 },
    { width: 18 },
    { width: 16 },
    { width: 18 },
    { width: 18 }
  ];

  let filename;
  if (customFileName && customFileName.trim()) {
    const safeName = customFileName.trim().replace(/[\/\\:\*\?"<>\|]/g, '_');
    filename = safeName.endsWith('.xlsx') ? safeName : `${safeName}.xlsx`;
  } else {
    filename = `Attendance_Monthly_${startDateStr}_to_${endDateStr}.xlsx`;
  }
  const filePath = path.join(TEMP_DIR, filename);

  await workbook.xlsx.writeFile(filePath);
  return { filePath, filename };
}

module.exports = {
  generateAttendanceExcel,
  generateFullDayExcel,
  generateRangeExcelReport
};
