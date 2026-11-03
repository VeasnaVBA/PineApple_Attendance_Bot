const { Bot, InputFile, InlineKeyboard } = require('grammy');
const dayjs = require('dayjs');
const fs = require('fs');
const path = require('path');

const {
  BOT_TOKEN,
  CLASS_NAME,
  TEACHER_NAME,
  getTeacherId,
  setTeacherId,
  getTeacherName,
  setTeacherName,
  getClearPassword,
  setClearPassword,
  getSettings
} = require('./config');

const { isGoogleSheetEnabled, callGoogleSheet } = require('./googleSheet');

const {
  getCurrentSlot,
  getTodaySlots,
  getTimetable,
  saveTimetable,
  syncTimetableFromGoogleSheet,
  getSubjectsDict,
  updateSubjectName,
  DAYS,
  KHMER_DAYS
} = require('./timetable');

const {
  getStudents,
  saveStudents,
  searchStudents,
  addStudent,
  updateStudent,
  removeStudent,
  parseAttendanceInput,
  getDailyRecords,
  recordAttendance,
  getDayAttendanceStatus,
  getRangeAttendanceStatus,
  clearAttendanceDatabase,
  clearStudentsDatabase,
  clearAllDatabase,
  normalizeTimeSlot,
  syncFromGoogleSheet
} = require('./attendance');

const {
  generateAttendanceExcel,
  generateFullDayExcel,
  generateRangeExcelReport
} = require('./excelReport');

if (!BOT_TOKEN || BOT_TOKEN === 'YOUR_TELEGRAM_BOT_TOKEN_HERE') {
  console.error('\n⚠️  ERROR: BOT_TOKEN is not set in .env!');
  process.exit(1);
}

const bot = new Bot(BOT_TOKEN);

// Active interactive sessions for recording attendance: userId -> { dateStr, timeSlot, subject, code, timestamp }
const pendingSessions = new Map();

// Active interactive sessions for naming custom date range reports: userId -> { startDateStr, endDateStr }
const pendingRangeRename = new Map();

// Helper: Check if caller is Teacher
function checkIsTeacher(ctx) {
  const teacherId = getTeacherId();
  if (!teacherId) return true;
  return String(ctx.from?.id) === String(teacherId);
}

// /setteacher command
bot.command('setteacher', async (ctx) => {
  const currentTeacherId = getTeacherId();
  const callerId = String(ctx.from.id);
  const parts = ctx.message.text.trim().split(/\s+/).slice(1);
  const customName = parts.join(' ').trim();
  const callerName = customName || ctx.from.first_name || 'លោកគ្រូ';

  if (currentTeacherId && currentTeacherId !== callerId) {
    return ctx.reply('⚠️ លោកគ្រូបន្ទុកថ្នាក់ត្រូវបានកំណត់រួចហើយ។ មិនអាចផ្លាស់ប្តូរបានទេ!');
  }

  setTeacherId(callerId, callerName);
  await ctx.reply(
    `✅ *បានកំណត់លោកគ្រូជោគជ័យ!*\n` +
    `👨‍🏫 *ឈ្មោះ:* ${callerName}\n` +
    `🆔 *User ID:* \`${callerId}\`\n\n` +
    `🔒 ចាប់ពីពេលនេះទៅ មានតែលោកគ្រូប៉ុណ្ណោះដែលអាច:\n` +
    `• ទាញយករបាយការណ៍ Excel (/attendance)\n` +
    `• បន្ថែម/កែប្រែ/លុបសិស្ស (/addstudent, /updatestudent, /removestudent)\n` +
    `• កែសម្រួលឈ្មោះមុខវិជ្ជា (/editsubject)\n` +
    `• ប្តូរឈ្មោះលោកគ្រូ (/setteachername <ឈ្មោះ>)\n\n` +
    `👉 ប្រធានថ្នាក់អាចធ្វើបានតែមួយគត់ គឺ *កត់ត្រាវត្តមាន*!`,
    { parse_mode: 'Markdown' }
  );
});

// /setteachername command
bot.command(['setteachername', 'teachername'], async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិកែប្រែឈ្មោះគ្រូបន្ទុកថ្នាក់ គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }
  const parts = ctx.message.text.trim().split(/\s+/).slice(1);
  const newName = parts.join(' ').trim();

  if (!newName) {
    return ctx.reply(
      `👨‍🏫 *ការកំណត់ឈ្មោះគ្រូបន្ទុកថ្នាក់ (Teacher Name)*\n\n` +
      `ឈ្មោះបច្ចុប្បន្ន: *${TEACHER_NAME}*\n\n` +
      `👉 ដើម្បីប្តូរឈ្មោះ សូមវាយ:\n` +
      `\`/setteachername <ឈ្មោះថ្មី>\`\n\n` +
      `*ឧទាហរណ៍:*\n` +
      `\`/setteachername សក់ វាសនា\`\n` +
      `\`/setteachername លោកគ្រូ សក់ វាសនា\``,
      { parse_mode: 'Markdown' }
    );
  }

  setTeacherName(newName);
  if (isGoogleSheetEnabled()) {
    try {
      await callGoogleSheet('saveSettings', { settings: { teacherName: newName } });
    } catch (e) {}
  }

  await ctx.reply(
    `✅ *បានផ្លាស់ប្តូរឈ្មោះគ្រូបន្ទុកថ្នាក់ជោគជ័យ!*\n\n` +
    `👨‍🏫 *ឈ្មោះថ្មី:* *${newName}*\n\n` +
    `📌 ឈ្មោះនេះនឹងបង្ហាញលើគ្រប់របាយការណ៍ Excel ទាំងអស់ចាប់ពីពេលនេះតទៅ!`,
    { parse_mode: 'Markdown' }
  );
});

// /teacher command
bot.command('teacher', async (ctx) => {
  const teacherId = getTeacherId();
  if (!teacherId) {
    return ctx.reply('⚠️ មិនទាន់មានលោកគ្រូចុះឈ្មោះទេ។ សូមវាយ `/setteacher` ដើម្បីចុះឈ្មោះជាលោកគ្រូ!', {
      parse_mode: 'Markdown'
    });
  }
  const isMe = String(ctx.from.id) === String(teacherId);
  await ctx.reply(
    `👨‍🏫 *ព័ត៌មានគ្រូបន្ទុកថ្នាក់:*\n` +
    `• ឈ្មោះ: *${TEACHER_NAME}*\n` +
    `• User ID: \`${teacherId}\`\n` +
    `• តួនាទីរបស់អ្នក: *${isMe ? 'លោកគ្រូ (Teacher ✅)' : 'ប្រធានថ្នាក់ (Class Monitor)'}*\n\n` +
    `💡 _លោកគ្រូអាចប្តូរឈ្មោះបានជានិច្ចដោយវាយ: \`/setteachername <ឈ្មោះថ្មី>\`_`,
    { parse_mode: 'Markdown' }
  );
});

// /start command
bot.command('start', async (ctx) => {
  const teacherId = getTeacherId();
  if (!teacherId && ctx.chat.type === 'private') {
    setTeacherId(ctx.from.id, ctx.from.first_name);
  }

  const welcomeText =
`👋 *សួស្តី! សូមស្វាគមន៍មកកាន់ ${CLASS_NAME}*
👨‍🏫 *គ្រូបន្ទុកថ្នាក់:* ${TEACHER_NAME}

🤖 *ជំនួយការកត់ត្រាវត្តមានសិស្ស និងបង្កើតរបាយការណ៍ Excel:*

📌 *របៀបប្រើប្រាស់:*
1️⃣ *សម្រាប់ប្រធានថ្នាក់ ឬលោកគ្រូ (កត់ត្រាវត្តមាន):*
• ចុចប៊ូតុង *[ ✍️ កត់ត្រាវត្តមាន ]* ខាងក្រោម
• ជ្រើសរើស **កាលបរិច្ឆេទ (Date)** ➡️ **ម៉ោង (Time)** ➡️ ឃើញ **មុខវិជ្ជា (Subject)**
• រួចវាយឈ្មោះ ឬលេខរៀងសិស្សដែល អវត្តមាន(A), ច្បាប់(P), ឬ យឺត(L) រួចចុច Send!
*(សិស្សផ្សេងទៀត គឺប្រព័ន្ធកត់ត្រាថា "មានវត្តមាន" ទាំងអស់ដោយស្វ័យប្រវត្តិ!)*

2️⃣ *សម្រាប់លោកគ្រូ (ទាញយករបាយការណ៍):*
វាយពាក្យ \`attendance\` ឬ /attendance នោះ Bot នឹងផ្ញើឯកសារ Excel (.xlsx) ជូនលោកគ្រូ!
*(🔒 របាយការណ៍ Excel ផ្ញើជូនតែលោកគ្រូប៉ុណ្ណោះ ប្រធានថ្នាក់មិនអាចទាញយកបានទេ)*`;

  const keyboard = new InlineKeyboard()
    .text('✍️ កត់ត្រាវត្តមាន (Record Attendance)', 'cmd_start_record')
    .row()
    .text('📥 របាយការណ៍ប្រចាំថ្ងៃ (Daily Excel)', 'cmd_attendance')
    .text('📆 របាយការណ៍ប្រចាំខែ (Monthly)', 'cmd_monthly')
    .row()
    .text('📅 កាលវិភាគថ្ងៃនេះ', 'cmd_today')
    .text('👥 បញ្ជីឈ្មោះសិស្ស', 'cmd_students')
    .row()
    .text('🧹 សម្អាតសារក្នុងឆាត (Clear Chat)', 'cmd_clean_menu');

  await ctx.reply(welcomeText, { parse_mode: 'Markdown', reply_markup: keyboard });
});

// /help command
bot.command('help', async (ctx) => {
  const isTeacher = checkIsTeacher(ctx);

  let helpText =
`📖 *ការណែនាំអំពីការប្រើប្រាស់ Bot:*

✍️ *១. របៀបកត់ត្រាវត្តមាន:*
• វាយបញ្ជា \`/record\` ឬចុចប៊ូតុង *[ ✍️ កត់ត្រាវត្តមាន ]*
• ជ្រើសរើស **ថ្ងៃ (Date)** ➡️ ជ្រើសរើស **ម៉ោង (Time)**
• បន្ទាប់មកវាយឈ្មោះ ឬលេខរៀងសិស្ស រួចចុច Send:
\`\`\`
ឆោម ចិន A
ខា សិនហួ P (ឈឺ)
06 L
\`\`\`

🔍 *២. របៀបស្វែងរកឈ្មោះសិស្ស (ពេលភ្លេចឈ្មោះ):*
• \`/find <ឈ្មោះ ឬ លេខរៀង>\` (ឧ. \`/find ចិន\` ឬ \`/find 04\`)
• \`/students\` (មើលបញ្ជីឈ្មោះសិស្សទាំងអស់)

🗓️ *៣. ពិនិត្យកាលវិភាគ:*
• \`/today\` (មើលម៉ោង និងមុខវិជ្ជាថ្ងៃនេះ)
• \`/timetable\` (មើលកាលវិភាគប្រចាំសប្តាហ៍)`;

  if (isTeacher) {
    helpText += `\n\n🔒 *៤. បញ្ជាផ្តាច់មុខសម្រាប់លោកគ្រូ (Teacher Only):*
• \`/attendance\` - របាយការណ៍ប្រចាំថ្ងៃ (ឬ \`/attendance 01-10-2026\`)
• \`/monthly\` - របាយការណ៍សង្ខេបប្រចាំខែ (Monthly Report)
• \`/range <ថ្ងៃចាប់ផ្តើម> to <ថ្ងៃបញ្ចប់>\`
  _ឧទាហរណ៍:_ \`/range 01-10-2026 to 31-10-2026\`
• \`/addstudent <លេខរៀង> <ឈ្មោះ> <ភេទ> [ថ្ងៃខែឆ្នាំកំណើត]\`
• \`/updatestudent <លេខរៀង> <ឈ្មោះកែសម្រួល>\`
• \`/removestudent <លេខរៀង ឬ ឈ្មោះ>\`
• \`/editsubject <CODE> <ឈ្មោះជាភាសាខ្មែរ>\`
• \`/setteachername <ឈ្មោះថ្មី>\` - កែប្រែឈ្មោះគ្រូបន្ទុកថ្នាក់លើរបាយការណ៍ Excel`;
  }

  helpText += `\n\n🧹 *៥. សម្អាតសារក្នុងឆាត (មិនប៉ះពាល់ទិន្នន័យ):*
• \`/clean\` ឬ \`/clear\` - សម្អាត ៣០ សារចុងក្រោយ
• \`/clean <ចំនួន>\` - សម្អាតតាមចំនួនកំណត់ (ឧ. \`/clean 50\`)
• ឬចុចប៊ូតុង *[ 🧹 សម្អាតសារក្នុងឆាត ]*
_🔒 ការសម្អាតសារ គឺលុបតែសារក្នុង Telegram ប៉ុណ្ណោះ ទិន្នន័យវត្តមាន និងសិស្សទាំងអស់ក្នុង Google Sheet / Database នៅតែមានពេញលេញ ១០០%!_

🗑 *៦. សម្អាតទិន្នន័យក្នុង Sheet (ទាមទារ Password):*
• \`/cleardata <password>\` ឬ \`/clearsheet <password>\`
  _លុបកំណត់ត្រាវត្តមានទាំងអស់ចេញពី Google Sheet និង Bot ទាំងស្រុង (ទុកបញ្ជីសិស្ស និងកាលវិភាគនៅដដែល)_
• \`/setclearpassword <password_ថ្មី>\` - ផ្លាស់ប្តូរ Password សម្អាត Sheet (សម្រាប់តែលោកគ្រូ)`;

  await ctx.reply(helpText, { parse_mode: 'Markdown' });
});

// /students command
bot.command('students', async (ctx) => {
  if (isGoogleSheetEnabled()) {
    await syncFromGoogleSheet();
  }
  const students = getStudents();
  if (students.length === 0) {
    return ctx.reply(
      '⚠️ *មិនទាន់មានទិន្នន័យសិស្សនៅឡើយទេ (០ នាក់)!*\n\n' +
      '👉 *របៀបបញ្ចូលឈ្មោះសិស្ស:*\n' +
      '1️⃣ បើក **Google Sheet** ➡️ ចូលសន្លឹក **"Students"** ➡️ បញ្ចូលឈ្មោះសិស្ស (No, ID, Name, Gender, DOB)\n' +
      '2️⃣ ឬវាយពាក្យបញ្ជាក្នុង Bot: `/addstudent <លេខរៀង> <ឈ្មោះ> <ភេទ>`\n' +
      '3️⃣ ឬប្រសិនបើលោកគ្រូចង់ទាញយកបញ្ជីសិស្សគំរូ ៣០ នាក់មកវិញ សូមវាយ: `/restorestudents`',
      { parse_mode: 'Markdown' }
    );
  }

  let text = `👥 *បញ្ជីឈ្មោះសិស្ស (${students.length} នាក់) - ${CLASS_NAME}:*\n\n`;
  students.forEach((s) => {
    const dob = s.dob ? ` (${s.dob})` : '';
    text += `${s.id}. *${s.name}* - ភេទ: ${s.gender}${dob}\n`;
  });

  await ctx.reply(text, { parse_mode: 'Markdown' });
});

// /restorestudents command (for Teacher to restore sample roster to Google Sheet)
bot.command('restorestudents', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិនេះសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }
  const defaultStudents = [
    { id: '01', name: 'ខា សិនហួ', gender: 'ប្រុស', dob: '02/01/2012' },
    { id: '02', name: 'ខេង គឹមឡាង', gender: 'ស្រី', dob: '07/01/2011' },
    { id: '03', name: 'ចំណាន លីនតុន', gender: 'ប្រុស', dob: '11/05/2012' },
    { id: '04', name: 'ឆោម ចិន', gender: 'ស្រី', dob: '25/06/2010' },
    { id: '05', name: 'ជួប តុនកែវ', gender: 'ស្រី', dob: '27/06/2012' },
    { id: '06', name: 'ឌៀម រតនវិសាល', gender: 'ប្រុស', dob: '06/10/2011' },
    { id: '07', name: 'ទេពមិន មុនីន', gender: 'ប្រុស', dob: '07/01/2012' },
    { id: '08', name: 'បូរី រស្មី', gender: 'ប្រុស', dob: '29/05/2012' },
    { id: '09', name: 'ផល្លា សូភីដានី', gender: 'ស្រី', dob: '13/01/2013' },
    { id: '10', name: 'ផាត សុផាន់ណា', gender: 'ស្រី', dob: '21/01/2011' },
    { id: '11', name: 'ផោន សុភ័ក្ត្រវឌ្ឍនា', gender: 'ស្រី', dob: '18/10/2012' },
    { id: '12', name: 'ភាន សុបញ្ញា', gender: 'ស្រី', dob: '02/08/2011' },
    { id: '13', name: 'យ៉ាំ សុធីម', gender: 'ស្រី', dob: '18/11/2011' },
    { id: '14', name: 'យ៉ុក ខិត្តាយ៉ា', gender: 'ស្រី', dob: '07/11/2011' },
    { id: '15', name: 'យ៉េត ករុណា', gender: 'ប្រុស', dob: '27/01/2013' },
    { id: '16', name: 'យុតពិសិដ្ឋសុគន្ធ សក្ខណា', gender: 'ស្រី', dob: '28/01/2012' },
    { id: '17', name: 'រិន រតនា', gender: 'ស្រី', dob: '29/06/2010' },
    { id: '18', name: 'រីម សុភ័ក្ត្រា', gender: 'ស្រី', dob: '09/04/2012' },
    { id: '19', name: 'លាង ហ៊ីយឡាង', gender: 'ស្រី', dob: '03/04/2012' },
    { id: '20', name: 'ស្រ៊ិន ពៅម៉ូនីកា', gender: 'ស្រី', dob: '11/01/2012' },
    { id: '21', name: 'សាន្ត សុខហែម', gender: 'ស្រី', dob: '18/11/2011' },
    { id: '22', name: 'សុជាតិ ហាណា', gender: 'ស្រី', dob: '31/01/2011' },
    { id: '23', name: 'សុក័ក្ត្រ ដាវីយ៉ា', gender: 'ស្រី', dob: '14/09/2011' },
    { id: '24', name: 'សួស្តី ឧស្សាហ៍', gender: 'ប្រុស', dob: '06/02/2012' },
    { id: '25', name: 'សៀង ចរិយា', gender: 'ស្រី', dob: '13/04/2011' },
    { id: '26', name: 'ហិន ដាលីស', gender: 'ស្រី', dob: '04/12/2012' },
    { id: '27', name: 'ហុង សុខជាង', gender: 'ប្រុស', dob: '22/07/2011' },
    { id: '28', name: 'ហេង ធាក់ទិន', gender: 'ប្រុស', dob: '27/02/2012' },
    { id: '29', name: 'ហេង ប៊ុនលាភ', gender: 'ប្រុស', dob: '07/03/2011' },
    { id: '30', name: 'ឡុង វិសិដ្ឋ', gender: 'ប្រុស', dob: '15/08/2012' }
  ];

  saveStudents(defaultStudents);
  if (isGoogleSheetEnabled()) {
    try {
      await callGoogleSheet('saveStudents', { students: defaultStudents });
    } catch (e) {}
  }
  await ctx.reply(
    `✅ *បានទាញយក និងរក្សាទុកបញ្ជីសិស្សគំរូ ៣០ នាក់ចូល Google Sheet រួចរាល់ហើយ!*`,
    { parse_mode: 'Markdown' }
  );
});

// /find or /search command
async function handleFindStudent(ctx) {
  const parts = ctx.message.text.split(/\s+/).slice(1);
  const query = parts.join(' ').trim();

  if (!query) {
    return ctx.reply(
      '🔍 *សូមបញ្ជាក់ឈ្មោះ ឬលេខរៀងសិស្សដែលចង់ស្វែងរក!*\n' +
      '• របៀបប្រើ: `/find <ឈ្មោះ ឬ លេខរៀង>`\n' +
      '• ឧទាហរណ៍:\n' +
      '  - `/find ចិន`\n' +
      '  - `/find ហេង`\n' +
      '  - `/find 04`\n\n' +
      '💡 ឬវាយ `/students` ដើម្បីមើលបញ្ជីសិស្សទាំងអស់។',
      { parse_mode: 'Markdown' }
    );
  }

  const results = searchStudents(query);

  if (results.length === 0) {
    return ctx.reply(
      `⚠️ រកមិនឃើញសិស្សដែលមានពាក្យគន្លឹះ "${query}" ទេ។\nសូមវាយ \`/students\` ដើម្បីផ្ទៀងផ្ទាត់បញ្ជីឈ្មោះ។`,
      { parse_mode: 'Markdown' }
    );
  }

  let text = `🔍 *លទ្ធផលស្វែងរក "${query}" (រកឃើញ ${results.length} នាក់):*\n\n`;
  results.forEach(s => {
    const dob = s.dob ? ` (កើត: ${s.dob})` : '';
    text += `• លេខរៀង \`${s.id}\` - *${s.name}* (${s.gender})${dob}\n`;
  });

  text += `\n✍️ *របៀបកត់វត្តមាន:*\n`;
  const first = results[0];
  text += `• \`${first.id} A\` ឬ \`${first.name} A\` (អវត្តមាន)\n`;
  text += `• \`${first.id} P\` ឬ \`${first.name} P\` (សុំច្បាប់)\n`;
  text += `• \`${first.id} L\` ឬ \`${first.name} L\` (មកយឺត)`;

  await ctx.reply(text, { parse_mode: 'Markdown' });
}

bot.command('find', handleFindStudent);
bot.command('search', handleFindStudent);
bot.command('s', handleFindStudent);

// --- INTERACTIVE ATTENDANCE RECORDING FLOW ---
// Step 1: Prompt Choose Date
async function promptChooseDate(ctx) {
  const now = dayjs();
  const todayStr = now.format('YYYY-MM-DD');
  const todayKhmer = KHMER_DAYS[DAYS[now.day()]];

  const yesterday = now.subtract(1, 'day');
  const yesterdayStr = yesterday.format('YYYY-MM-DD');
  const yesterdayKhmer = KHMER_DAYS[DAYS[yesterday.day()]];

  const tomorrow = now.add(1, 'day');
  const tomorrowStr = tomorrow.format('YYYY-MM-DD');
  const tomorrowKhmer = KHMER_DAYS[DAYS[tomorrow.day()]];

  const keyboard = new InlineKeyboard()
    .text(`📅 ថ្ងៃនេះ (${todayKhmer}, ${now.format('DD/MM')})`, `rec_date_${todayStr}`)
    .row()
    .text(`📅 ម្សិលមិញ (${yesterdayKhmer}, ${yesterday.format('DD/MM')})`, `rec_date_${yesterdayStr}`)
    .row()
    .text(`📅 ថ្ងៃស្អែក (${tomorrowKhmer}, ${tomorrow.format('DD/MM')})`, `rec_date_${tomorrowStr}`)
    .row()
    .text('❌ បោះបង់ (Cancel)', 'rec_cancel');

  const text =
`📋 *កត់ត្រាវត្តមានសិស្ស (Record Attendance)*
👉 *ជំហានទី ១:* សូមជ្រើសរើស **កាលបរិច្ឆេទ (Date)** ខាងក្រោម៖`;

  if (ctx.callbackQuery) {
    await ctx.editMessageText(text, { parse_mode: 'Markdown', reply_markup: keyboard });
  } else {
    await ctx.reply(text, { parse_mode: 'Markdown', reply_markup: keyboard });
  }
}

// Commands to trigger record flow
bot.command('record', promptChooseDate);
bot.command('mark', promptChooseDate);
bot.callbackQuery('cmd_start_record', async (ctx) => {
  await ctx.answerCallbackQuery();
  await promptChooseDate(ctx);
});

// Step 2: Handle Date chosen -> Prompt Choose Time Slot
bot.callbackQuery(/^rec_date_(.+)$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const dateStr = ctx.match[1];
  const targetDate = dayjs(dateStr);
  const dayName = DAYS[targetDate.day()];
  const khmerDay = KHMER_DAYS[dayName];

  // Fetch complete day status to see which slots have already been recorded
  const dayStatus = await getDayAttendanceStatus(dateStr, { DAYS, KHMER_DAYS, getTodaySlots });
  const slots = getTodaySlots(targetDate);

  if (slots.length === 0 && dayStatus.scheduledSessions.length === 0) {
    const keyboard = new InlineKeyboard()
      .text('🔙 ជ្រើសរើសថ្ងៃផ្សេង', 'cmd_start_record')
      .row()
      .text('❌ បោះបង់', 'rec_cancel');

    return ctx.editMessageText(
      `📅 *កាលបរិច្ឆេទ: ថ្ងៃ${khmerDay} (${targetDate.format('DD-MM-YYYY')})*\n\n` +
      `⚠️ មិនមានម៉ោងសិក្សាក្នុងកាលវិភាគសម្រាប់ថ្ងៃនេះទេ។`,
      { parse_mode: 'Markdown', reply_markup: keyboard }
    );
  }

  const keyboard = new InlineKeyboard();

  if (slots.length > 0) {
    slots.forEach((slot, idx) => {
      const slotTime = `${slot.start} - ${slot.end}`;
      const normSlot = normalizeTimeSlot(slotTime);
      const matchedSess = (dayStatus.scheduledSessions && dayStatus.scheduledSessions[idx]) ||
        dayStatus.scheduledSessions.find(s => normalizeTimeSlot(s.timeSlot) === normSlot);

      if (matchedSess && matchedSess.isRecorded) {
        const sum = matchedSess.summary || {};
        let noteStr = '✅ បានកត់រួច';
        if (sum.absent > 0 || sum.permission > 0 || sum.late > 0) {
          noteStr += ` (A:${sum.absent || 0}, P:${sum.permission || 0}${sum.late ? `, L:${sum.late}` : ''})`;
        } else {
          noteStr += ' (គ្រប់ 30)';
        }
        keyboard.text(
          `✅ ${slot.start}-${slot.end}: ${slot.subject} ${noteStr}`,
          `rec_slot_${dateStr}_${slot.start.replace(':', '')}_${slot.end.replace(':', '')}`
        ).row();
      } else {
        keyboard.text(
          `⚪ ${slot.start}-${slot.end}: ${slot.subject} (មិនទាន់កត់)`,
          `rec_slot_${dateStr}_${slot.start.replace(':', '')}_${slot.end.replace(':', '')}`
        ).row();
      }
    });
  } else {
    // Weekend / extra recorded sessions
    dayStatus.scheduledSessions.forEach(sess => {
      const sum = sess.summary || {};
      let noteStr = '✅ បានកត់រួច';
      if (sum.absent > 0 || sum.permission > 0) {
        noteStr += ` (A:${sum.absent || 0}, P:${sum.permission || 0})`;
      }
      keyboard.text(
        `✅ ${sess.timeSlot}: ${sess.subject} ${noteStr}`,
        `rec_slot_${dateStr}_${sess.timeSlot.replace(/[^0-9]/g, '').slice(0, 4) || '0700'}_${sess.timeSlot.replace(/[^0-9]/g, '').slice(4, 8) || '0800'}`
      ).row();
    });
  }

  keyboard.text('🔙 ថយក្រោយ (Back)', 'cmd_start_record').text('❌ បោះបង់', 'rec_cancel');

  const recordedCount = dayStatus.scheduledSessions.filter(s => s.isRecorded).length;
  const totalSlotsCount = slots.length > 0 ? slots.length : dayStatus.scheduledSessions.length;

  const text =
`📋 *កត់ត្រាវត្តមានសិស្ស (Record Attendance)*
📅 *កាលបរិច្ឆេទ:* ថ្ងៃ${khmerDay}, ${targetDate.format('DD MMMM YYYY')}
📊 *ស្ថានភាព:* បានកត់ត្រា *${recordedCount}/${totalSlotsCount}* ម៉ោង

🟢 = ម៉ោងបានកត់រួច (ចុចលើវាដើម្បី *កត់សិស្សបន្ថែម* ឬ *កែប្រែ*)
⚪ = ម៉ោងមិនទាន់បានកត់ត្រា (ចុចដើម្បីចាប់ផ្តើមកត់)

👉 *សូមជ្រើសរើស ម៉ោងសិក្សា (Time Slot) ដើម្បីបន្ត៖*`;

  await ctx.editMessageText(text, { parse_mode: 'Markdown', reply_markup: keyboard });
});

// Step 3: Handle Time Slot chosen -> Show confirmed subject and prompt to type student names!
bot.callbackQuery(/^rec_slot_([^_]+)_([^_]+)_(.+)$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const dateStr = ctx.match[1];
  const rawStart = ctx.match[2];
  const rawEnd = ctx.match[3];

  const startTime = rawStart.length === 4 ? `${rawStart.slice(0, 2)}:${rawStart.slice(2)}` : rawStart;
  const endTime = rawEnd.length === 4 ? `${rawEnd.slice(0, 2)}:${rawEnd.slice(2)}` : rawEnd;

  const targetDate = dayjs(dateStr);
  const dayName = DAYS[targetDate.day()];
  const khmerDay = KHMER_DAYS[dayName];

  const timetable = getTimetable();
  const slot = timetable.find(s =>
    s.day.toLowerCase() === dayName.toLowerCase() &&
    s.start === startTime &&
    s.end === endTime
  );

  const subject = slot ? slot.subject : 'មុខវិជ្ជាទូទៅ';
  const code = slot ? slot.code : '';

  // Save session for this user
  pendingSessions.set(ctx.from.id, {
    dateStr,
    timeSlot: `${startTime} - ${endTime}`,
    subject,
    code,
    khmerDay,
    chatId: ctx.chat.id,
    timestamp: Date.now()
  });

  // Check if slot has existing record
  const daily = getDailyRecords(dateStr);
  const slotTime = `${startTime} - ${endTime}`;
  const normSlot = normalizeTimeSlot(slotTime);
  const existingRecord = daily[normSlot] || daily[slotTime] || daily[`${startTime}-${endTime}`];
  const isRecorded = existingRecord && Array.isArray(existingRecord.records) && existingRecord.records.length > 0;

  let statusBanner = '';
  if (isRecorded) {
    const nonPresent = existingRecord.records.filter(r => r.status !== 'Present');
    statusBanner = `ℹ️ *ម៉ោងនេះបានកត់ត្រារួចម្តងហើយ!* (កត់ត្រាដោយ: ${existingRecord.recordedBy || 'ប្រធានថ្នាក់'})\n`;
    if (nonPresent.length > 0) {
      statusBanner += `📋 *សិស្សដែលបានកត់កន្លងមក (${nonPresent.length} នាក់):*\n`;
      statusBanner += nonPresent.map(r => `• \`${r.id}\` ${r.name} - *${r.code}* ${r.note ? `(${r.note})` : ''}`).join('\n') + `\n\n`;
      statusBanner += `➕ *លោកគ្រូ ឬប្រធានថ្នាក់ អាចវាយឈ្មោះ/លេខសិស្សបន្ថែម* (សិស្សចាស់ៗនៅតែរក្សាទុកដដែល) ឬវាយ \`01 V\` ដើម្បីប្តូរមក "មានវត្តមាន" វិញ។`;
    } else {
      statusBanner += `✅ កាលពីកត់លើកមុន គឺ *មានវត្តមានពេញលេញទាំង ${existingRecord.records.length} នាក់*។\n\n` +
        `➕ *លោកគ្រូ ឬប្រធានថ្នាក់ អាចវាយឈ្មោះ/លេខសិស្សបន្ថែម* ប្រសិនបើមានសិស្សអវត្តមាន (A), សុំច្បាប់ (P), ឬ មកយឺត (L)។`;
    }
  } else {
    statusBanner =
`👇 *ជំហានទី ៣:* សូមវាយឈ្មោះ ឬលេខរៀងសិស្សដែល អវត្តមាន(A), សុំច្បាប់(P), ឬ មកយឺត(L) រួចចុច **Send/Enter**:

*ឧទាហរណ៍:*
\`\`\`
ឆោម ចិន A
06 L
ខា សិនហួ P (ឈឺ)
\`\`\`
_(💡 សិស្សដែលមិនបានសរសេរ គឺប្រព័ន្ធកត់ត្រាថា "មានវត្តមាន" ទាំងអស់ដោយស្វ័យប្រវត្តិ)_`;
  }

  const promptText =
`📋 *កត់ត្រាវត្តមានសិស្ស (Record Attendance)*
📅 *កាលបរិច្ឆេទ:* ថ្ងៃ${khmerDay}, ${targetDate.format('DD-MM-YYYY')}
⏰ *ម៉ោងសិក្សា:* \`${startTime} - ${endTime}\`
📚 *មុខវិជ្ជា:* *${subject}*

${statusBanner}

_(វាយ \`/cancel\` ប្រសិនបើចង់បោះបង់)_`;

  const keyboard = new InlineKeyboard()
    .text('🔙 ជ្រើសរើសម៉ោងផ្សេង', `rec_date_${dateStr}`)
    .text('❌ បោះបង់ (Cancel)', 'rec_cancel');
  await ctx.editMessageText(promptText, { parse_mode: 'Markdown', reply_markup: keyboard });
});

// Cancel active recording session
bot.callbackQuery('rec_cancel', async (ctx) => {
  await ctx.answerCallbackQuery();
  pendingSessions.delete(ctx.from.id);
  await ctx.editMessageText('❌ បានបោះបង់ការកត់ត្រាវត្តមាន។');
});

// /today command
bot.command('today', async (ctx) => {
  if (isGoogleSheetEnabled()) {
    await syncTimetableFromGoogleSheet();
  }
  const now = dayjs();
  const dayName = DAYS[now.day()];
  const khmerDay = KHMER_DAYS[dayName];
  const slots = getTodaySlots(now);
  const currentSlot = getCurrentSlot(now);

  if (slots.length === 0) {
    return ctx.reply(`📅 *ថ្ងៃនេះជាថ្ងៃ ${khmerDay} (${dayName})*\nមិនមានម៉ោងសិក្សាទេ!`, { parse_mode: 'Markdown' });
  }

  let text = `📅 *កាលវិភាគថ្ងៃនេះ (ថ្ងៃ${khmerDay}, ${now.format('DD-MM-YYYY')}):*\n\n`;
  slots.forEach(s => {
    const isCurrent = currentSlot && currentSlot.start === s.start;
    const badge = isCurrent ? ' 👉 *(ម៉ោងបច្ចុប្បន្ន)*' : '';
    text += `⏰ \`${s.start} - ${s.end}\`: *${s.subject}* ${badge}\n`;
  });

  if (currentSlot) {
    text += `\n🎯 *មុខវិជ្ជាបច្ចុប្បន្ន:* *${currentSlot.subject}* (\`${currentSlot.start} - ${currentSlot.end}\`)`;
  }

  const keyboard = new InlineKeyboard().text('✍️ កត់ត្រាវត្តមានម៉ោងនេះ', 'cmd_start_record');
  await ctx.reply(text, { parse_mode: 'Markdown', reply_markup: keyboard });
});

// /timetable command
bot.command('timetable', async (ctx) => {
  if (isGoogleSheetEnabled()) {
    await syncTimetableFromGoogleSheet();
  }
  const timetable = getTimetable();
  if (timetable.length === 0) {
    return ctx.reply('មិនទាន់មានកាលវិភាគនៅឡើយទេ។');
  }

  const grouped = {};
  timetable.forEach(slot => {
    if (!grouped[slot.day]) grouped[slot.day] = [];
    grouped[slot.day].push(slot);
  });

  let text = `🗓️ *កាលវិភាគប្រចាំសប្តាហ៍ - ${CLASS_NAME}*\n\n`;
  for (const day of ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']) {
    if (grouped[day] && grouped[day].length > 0) {
      text += `📌 *ថ្ងៃ${KHMER_DAYS[day]} (${day}):*\n`;
      grouped[day].forEach(s => {
        text += `  • \`${s.start} - ${s.end}\`: *${s.subject}*\n`;
      });
      text += '\n';
    }
  }

  text += `\n💡 *ចំណាំ:* លោកគ្រូអាចកែប្រែកាលវិភាគនេះដោយផ្ទាល់ក្នុងសន្លឹក **"Timetable"** ក្នុង Google Sheet បានគ្រប់ពេលវេលា!`;

  await ctx.reply(text, { parse_mode: 'Markdown' });
});

// /synctimetable command
bot.command('synctimetable', async (ctx) => {
  await syncTimetableFromGoogleSheet();
  const timetable = getTimetable();
  await ctx.reply(`✅ *បានទាញយក និងធ្វើបច្ចុប្បន្នភាពកាលវិភាគពី Google Sheet រួចរាល់!* (សរុប ${timetable.length} ម៉ោងសិក្សា)`, { parse_mode: 'Markdown' });
});

// /restoretimetable command (Teacher only)
bot.command('restoretimetable', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិនេះគឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }
  const defaultTimetable = readJson('timetable.json', []);
  saveTimetable(defaultTimetable);
  if (isGoogleSheetEnabled()) {
    try {
      await callGoogleSheet('saveTimetable', { timetable: defaultTimetable });
    } catch (e) {}
  }
  await ctx.reply(
    `✅ *បានបញ្ចូលកាលវិភាគគំរូ ចូល Google Sheet ឡើងវិញដោយជោគជ័យ!* (សរុប ${defaultTimetable.length} ម៉ោង)`,
    { parse_mode: 'Markdown' }
  );
});

// /subjects command
bot.command('subjects', async (ctx) => {
  const dict = getSubjectsDict();
  let text = `📚 *បញ្ជីមុខវិជ្ជា និងកូដ:*\n\n`;
  for (const [code, name] of Object.entries(dict)) {
    text += `• \`${code}\` ➡️ *${name}*\n`;
  }
  if (checkIsTeacher(ctx)) {
    text += `\n💡 *របៀបកែសម្រួលឈ្មោះ:* វាយ \`/editsubject <CODE> <ឈ្មោះថ្មី>\`\n_ឧទាហរណ៍:_ \`/editsubject M2 គណិតវិទ្យា\``;
  }
  await ctx.reply(text, { parse_mode: 'Markdown' });
});

// /editsubject command (Teacher Only)
bot.command('editsubject', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិកែប្រែឈ្មោះមុខវិជ្ជា គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }

  const parts = ctx.message.text.split(/\s+/).slice(1);
  if (parts.length < 2) {
    return ctx.reply(
      '❌ *ទម្រង់មិនត្រឹមត្រូវ!*\nរបៀបប្រើ: `/editsubject <CODE> <ឈ្មោះមុខវិជ្ជាថ្មី>`\nឧទាហរណ៍: `/editsubject M2 គណិតវិទ្យា`',
      { parse_mode: 'Markdown' }
    );
  }

  const code = parts[0].toUpperCase();
  const newName = parts.slice(1).join(' ').trim();

  await updateSubjectName(code, newName);
  await ctx.reply(`✅ បានកែសម្រួលមុខវិជ្ជាកូដ \`${code}\` ទៅជា: *${newName}* ដោយជោគជ័យ!`, {
    parse_mode: 'Markdown'
  });
});

// /addstudent command (Teacher Only)
bot.command('addstudent', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិបន្ថែមសិស្ស គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }

  const parts = ctx.message.text.split(/\s+/).slice(1);
  if (parts.length < 2) {
    return ctx.reply(
      '❌ *ទម្រង់មិនត្រឹមត្រូវ!*\nរបៀបប្រើ: `/addstudent <លេខរៀង> <ឈ្មោះ> <ភេទ> [ថ្ងៃខែឆ្នាំកំណើត]`\nឧទាហរណ៍: `/addstudent 31 សុខ សាន ប្រុស 12/04/2012`',
      { parse_mode: 'Markdown' }
    );
  }

  const id = parts[0];
  let gender = 'មិនបញ្ជាក់';
  let dob = '';
  let nameParts = parts.slice(1);

  if (nameParts.length > 0 && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(nameParts[nameParts.length - 1])) {
    dob = nameParts.pop();
  }

  if (nameParts.length > 0) {
    const gCheck = nameParts[nameParts.length - 1];
    if (['ប្រុស', 'ស្រី', 'ប', 'ស', 'male', 'female', 'm', 'f'].includes(gCheck.toLowerCase())) {
      if (['ស្រី', 'ស', 'female', 'f'].includes(gCheck.toLowerCase())) gender = 'ស្រី';
      else gender = 'ប្រុស';
      nameParts.pop();
    }
  }

  const fullName = nameParts.join(' ').trim();
  addStudent(id, fullName, gender, dob);

  await ctx.reply(`✅ បានបញ្ចូលសិស្សថ្មី: *${fullName}* (លេខរៀង: \`${id}\`, ភេទ: ${gender}${dob ? `, ថ្ងៃខែ: ${dob}` : ''}) ដោយជោគជ័យ!`, {
    parse_mode: 'Markdown'
  });
});

// /updatestudent command (Teacher Only)
bot.command('updatestudent', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិកែប្រែព័ត៌មានសិស្ស គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }

  const parts = ctx.message.text.split(/\s+/).slice(1);
  if (parts.length < 2) {
    return ctx.reply('❌ របៀបប្រើ: `/updatestudent <លេខរៀង> <ឈ្មោះកែសម្រួល>`\nឧទាហរណ៍: `/updatestudent 04 ឆោម ចិនណា`', {
      parse_mode: 'Markdown'
    });
  }

  const id = parts[0];
  const newName = parts.slice(1).join(' ').trim();
  const updated = updateStudent(id, newName);

  if (updated) {
    await ctx.reply(`✅ បានកែសម្រួលឈ្មោះសិស្សលេខ \`${id}\` ទៅជា: *${newName}* រួចរាល់!`, {
      parse_mode: 'Markdown'
    });
  } else {
    await ctx.reply(`⚠️ រកមិនឃើញសិស្សដែលមានលេខរៀង \`${id}\` ទេ។`, { parse_mode: 'Markdown' });
  }
});

// /removestudent command (Teacher Only)
bot.command('removestudent', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិលុបសិស្ស គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }

  const query = ctx.message.text.split(/\s+/).slice(1).join(' ').trim();
  if (!query) {
    return ctx.reply('❌ សូមបញ្ជាក់លេខរៀង ឬឈ្មោះសិស្សដែលចង់លុប។\nឧទាហរណ៍: `/removestudent 30`', {
      parse_mode: 'Markdown'
    });
  }

  const removed = removeStudent(query);
  if (removed) {
    await ctx.reply(`✅ បានលុបសិស្ស "${query}" ចេញពីបញ្ជីរៀនរួចរាល់។`);
  } else {
    await ctx.reply(`⚠️ រកមិនឃើញសិស្សដែលមានឈ្មោះ ឬលេខរៀង "${query}" ទេ។`);
  }
});

// /id command
bot.command('id', async (ctx) => {
  const isMeTeacher = checkIsTeacher(ctx);
  await ctx.reply(
    `🆔 *ព័ត៌មានអត្តសញ្ញាណ (IDs):*\n` +
    `• *Group Chat ID:* \`${ctx.chat.id}\`\n` +
    `• *Your User ID:* \`${ctx.from.id}\`\n` +
    `• *Chat Type:* ${ctx.chat.type}\n` +
    `• *តួនាទីរបស់អ្នក:* ${isMeTeacher ? 'លោកគ្រូ (Teacher ✅)' : 'ប្រធានថ្នាក់ (Monitor)'}`,
    { parse_mode: 'Markdown' }
  );
});

// --- CHAT MESSAGE CLEANER (Does NOT touch Database / Google Sheet) ---
async function cleanChatMessages(ctx, count = 30) {
  const currentMsgId = ctx.message ? ctx.message.message_id : ctx.callbackQuery?.message?.message_id;
  if (!currentMsgId) return;

  const chatId = ctx.chat.id;
  const numToDelete = Math.min(Math.max(parseInt(count, 10) || 30, 1), 100);

  const messageIds = [];
  for (let i = 0; i < numToDelete; i++) {
    const id = currentMsgId - i;
    if (id > 0) messageIds.push(id);
  }

  let deletedCount = 0;
  try {
    await ctx.api.deleteMessages(chatId, messageIds);
    deletedCount = messageIds.length;
  } catch (batchErr) {
    for (const msgId of messageIds) {
      try {
        await ctx.api.deleteMessage(chatId, msgId);
        deletedCount++;
      } catch (singleErr) {}
    }
  }

  const dismissKeyboard = new InlineKeyboard().text('🗑️ បិទការជូនដំណឹងនេះ (Dismiss)', 'clean_dismiss');

  try {
    const notice = await ctx.api.sendMessage(
      chatId,
      `🧹 *បានសម្អាតសារក្នុងឆាតរួចរាល់! (${deletedCount} សារ)*\n\n` +
      `🔒 *សុវត្ថិភាពទិន្នន័យ:* ការសម្អាតនេះ គឺលុបតែសារដែលបានផ្ញើក្នុង Telegram ប៉ុណ្ណោះ។\n` +
      `**រាល់ទិន្នន័យវត្តមាន និងឈ្មោះសិស្សក្នុង Google Sheet / Database គឺនៅរក្សាទុកពេញលេញ ១០០% មិនបាត់បង់ឡើយ!**`,
      { parse_mode: 'Markdown', reply_markup: dismissKeyboard }
    );

    setTimeout(async () => {
      try {
        await ctx.api.deleteMessage(chatId, notice.message_id);
      } catch (e) {}
    }, 5000);
  } catch (err) {}
}

// Commands: /clean, /clear, /clearchat
bot.command(['clean', 'clear', 'clearchat'], async (ctx) => {
  const parts = ctx.message.text.trim().split(/\s+/).slice(1);
  const count = parts[0] ? parseInt(parts[0], 10) : 30;
  await cleanChatMessages(ctx, count);
});

bot.callbackQuery('cmd_clean_menu', async (ctx) => {
  await ctx.answerCallbackQuery();
  const keyboard = new InlineKeyboard()
    .text('🧹 សម្អាត ២០ សារ', 'clean_count_20')
    .text('🧹 សម្អាត ៥០ សារ', 'clean_count_50')
    .row()
    .text('🧹 សម្អាត ១០០ សារ', 'clean_count_100')
    .row()
    .text('🗑️ សម្អាតទិន្នន័យក្នុង Sheet (ទាមទារ Password)', 'cmd_open_clearsheet_help')
    .row()
    .text('❌ បោះបង់ (Cancel)', 'clean_dismiss');

  await ctx.reply(
    `🧹 *ជម្រើសសម្អាតសារក្នុងឆាត (Clear Messages)*\n\n` +
    `👉 សូមជ្រើសរើសចំនួនសារដែលលោកគ្រូចង់សម្អាត:\n\n` +
    `🔒 *ចំណាំសំខាន់:* ការសម្អាតនេះ គឺលុបតែសារក្នុង Telegram ដើម្បីឱ្យឆាតស្រឡះស្អាតប៉ុណ្ណោះ។ **ទិន្នន័យវត្តមាន និងសិស្សទាំងអស់ក្នុង Google Sheet មិនត្រូវបានប៉ះពាល់ឡើយ!**`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
});

bot.callbackQuery(/^clean_count_(\d+)$/, async (ctx) => {
  await ctx.answerCallbackQuery();
  const count = parseInt(ctx.match[1], 10) || 30;
  await cleanChatMessages(ctx, count);
});

bot.callbackQuery('clean_dismiss', async (ctx) => {
  await ctx.answerCallbackQuery();
  try {
    await ctx.deleteMessage();
  } catch (e) {}
});

// Help callback from clean menu
bot.callbackQuery('cmd_open_clearsheet_help', async (ctx) => {
  await ctx.answerCallbackQuery();
  await ctx.reply(
    `🔑 *ការសម្អាតទិន្នន័យក្នុង Google Sheet (ទាមទារ Password)*\n\n` +
    `ដើម្បីជ្រើសរើសទិន្នន័យដែលចង់លុប សូមវាយពាក្យបញ្ជា:\n` +
    `\`/cleardata <password>\`\n\n` +
    `👉 *ឧទាហរណ៍:* \`/cleardata pineapple123\`\n` +
    `_(Password លំនាំដើមគឺ \`pineapple123\`)_`,
    { parse_mode: 'Markdown' }
  );
});

// --- CLEAR SHEET DATA (Requires Token Password & Lets User Choose Data) ---
bot.command(['cleardata', 'clearsheet'], async (ctx) => {
  const text = ctx.message.text.trim();
  const parts = text.split(/\s+/).slice(1);
  const inputPassword = parts[0] ? parts[0].trim() : '';

  // Delete message to avoid showing password in chat history
  try {
    await ctx.deleteMessage();
  } catch (e) {}

  if (!inputPassword) {
    return ctx.reply(
      `⚠️ *ពាក្យបញ្ជាសម្អាតទិន្នន័យក្នុង Sheet (Clear Sheet Data)*\n\n` +
      `🔑 *ដើម្បីចូលទៅកាន់ផ្ទាំងជ្រើសរើសទិន្នន័យដែលត្រូវសម្អាត សូមបញ្ចូល Password:*\n` +
      `\`/cleardata <password>\`\n` +
      `ឬ\n` +
      `\`/clearsheet <password>\`\n\n` +
      `👉 *ឧទាហរណ៍:* \`/cleardata pineapple123\`\n\n` +
      `🔒 _ចំណាំ: Password លំនាំដើមគឺ \`pineapple123\` (លោកគ្រូអាចផ្លាស់ប្តូរបានក្នុង .env ឬប្រើ /setclearpassword)_`,
      { parse_mode: 'Markdown' }
    );
  }

  const expectedPassword = getClearPassword();
  if (inputPassword !== expectedPassword) {
    return ctx.reply(
      `❌ *Password មិនត្រឹមត្រូវទេ! (Access Denied)*\n` +
      `ប្រតិបត្តិការសម្អាតទិន្នន័យក្នុង Google Sheet ត្រូវបានបដិសេធ ដើម្បីសុវត្ថិភាព។`,
      { parse_mode: 'Markdown' }
    );
  }

  const todayStr = dayjs().format('YYYY-MM-DD');
  const studentCount = getStudents().length;
  const keyboard = new InlineKeyboard()
    .text(`📅 លុបវត្តមានតែថ្ងៃនេះ (${todayStr})`, 'cls_req_today')
    .row()
    .text('🗓️ លុបវត្តមានទាំងអស់ (All Attendance)', 'cls_req_all_att')
    .row()
    .text('👥 សម្អាតបញ្ជីសិស្ស (Clear Students)', 'cls_req_students')
    .row()
    .text('💣 លុប Reset ទាំងអស់ (Reset Everything)', 'cls_req_reset_all')
    .row()
    .text('❌ បោះបង់ (Cancel)', 'clean_dismiss');

  await ctx.reply(
    `🗂️ *ផ្ទាំងជ្រើសរើសទិន្នន័យដែលត្រូវសម្អាត (Clear Sheet Data Menu)*\n\n` +
    `🔑 *Password បានផ្ទៀងផ្ទាត់ជោគជ័យ!* ✅\n\n` +
    `👉 សូមជ្រើសរើសប្រភេទនៃទិន្នន័យដែលលោកគ្រូចង់លុបចេញពី Google Sheet:\n\n` +
    `1️⃣ *លុបវត្តមានតែថ្ងៃនេះ (${todayStr}):* លុបតែកំណត់ត្រាក្នុងសន្លឹក Attendance ថ្ងៃនេះ ទុកថ្ងៃផ្សេងៗ និងសិស្សនៅដដែល\n` +
    `2️⃣ *លុបវត្តមានទាំងអស់ (All Attendance):* លុបកំណត់ត្រាក្នុងសន្លឹក Attendance គ្រប់ថ្ងៃ (ទុកបញ្ជីសិស្ស ${studentCount} នាក់ក្នុងសន្លឹក Students នៅដដែល)\n` +
    `3️⃣ *សម្អាតបញ្ជីសិស្ស (Clear Students):* សម្អាតតារាងបញ្ជីឈ្មោះសិស្សក្នុងសន្លឹក Students\n` +
    `4️⃣ *លុប Reset ទាំងអស់ (Reset Everything):* លុបទាំងកំណត់ត្រាវត្តមាន និងបញ្ជីសិស្សទាំងស្រុង`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
});

// Confirmation handlers
bot.callbackQuery('cls_req_today', async (ctx) => {
  await ctx.answerCallbackQuery();
  const todayStr = dayjs().format('YYYY-MM-DD');
  const keyboard = new InlineKeyboard()
    .text('✅ បាទ/ចាស! លុបវត្តមានថ្ងៃនេះ', 'cls_do_today')
    .text('❌ ទេ! បោះបង់', 'clean_dismiss');

  await ctx.editMessageText(
    `⚠️ *ការបញ្ជាក់លុបទិន្នន័យ (Confirmation)*\n\n` +
    `តើលោកគ្រូពិតជាចង់ *លុបកំណត់ត្រាវត្តមានតែថ្ងៃនេះ (${todayStr})* ចេញពីសន្លឹក Attendance ក្នុង Google Sheet មែនទេ?\n` +
    `*(កំណត់ត្រាថ្ងៃមុនៗ និងបញ្ជីឈ្មោះសិស្ស គឺមិនត្រូវបានប៉ះពាល់ឡើយ)*`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
});

bot.callbackQuery('cls_req_all_att', async (ctx) => {
  await ctx.answerCallbackQuery();
  const studentCount = getStudents().length;
  const keyboard = new InlineKeyboard()
    .text('✅ បាទ/ចាស! លុបវត្តមានទាំងអស់', 'cls_do_all_att')
    .text('❌ ទេ! បោះបង់', 'clean_dismiss');

  await ctx.editMessageText(
    `⚠️ *ការបញ្ជាក់លុបទិន្នន័យ (Confirmation)*\n\n` +
    `តើលោកគ្រូពិតជាចង់ *លុបកំណត់ត្រាវត្តមានទាំងអស់ (All Attendance)* ចេញពីសន្លឹក Attendance ក្នុង Google Sheet មែនទេ?\n` +
    `*(បញ្ជីឈ្មោះសិស្សទាំង ${studentCount} នាក់ក្នុងសន្លឹក Students និងកាលវិភាគនៅរក្សាទុកដដែល)*`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
});

bot.callbackQuery('cls_req_students', async (ctx) => {
  await ctx.answerCallbackQuery();
  const keyboard = new InlineKeyboard()
    .text('⚠️ បាទ/ចាស! សម្អាតបញ្ជីសិស្ស', 'cls_do_students')
    .text('❌ ទេ! បោះបង់', 'clean_dismiss');

  await ctx.editMessageText(
    `🚨 *ការបញ្ជាក់សម្អាតបញ្ជីសិស្ស (Confirmation)*\n\n` +
    `តើលោកគ្រូពិតជាចង់ *សម្អាតបញ្ជីឈ្មោះសិស្ស (Students Sheet)* មែនទេ?\n` +
    `*(ឈ្មោះសិស្សទាំងអស់នឹងត្រូវបានលុបចេញពី Sheet)*`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
});

bot.callbackQuery('cls_req_reset_all', async (ctx) => {
  await ctx.answerCallbackQuery();
  const keyboard = new InlineKeyboard()
    .text('💣 បាទ/ចាស! Reset ទាំងអស់', 'cls_do_reset_all')
    .text('❌ ទេ! បោះបង់', 'clean_dismiss');

  await ctx.editMessageText(
    `🚨 *ការបញ្ជាក់ Reset ទិន្នន័យទាំងអស់ (Confirmation)*\n\n` +
    `តើលោកគ្រូពិតជាចង់ *លុបទិន្នន័យទាំងអស់ (Reset Everything)* មែនទេ?\n` +
    `*(ទាំងកំណត់ត្រាវត្តមាន និងបញ្ជីឈ្មោះសិស្ស នឹងត្រូវបានលុបចេញទាំងស្រុង)*`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
});

// Execution handlers
bot.callbackQuery('cls_do_today', async (ctx) => {
  await ctx.answerCallbackQuery('⏳ កំពុងលុបវត្តមានថ្ងៃនេះ...');
  const todayStr = dayjs().format('YYYY-MM-DD');
  ctx.editMessageText(`⏳ *កំពុងលុបកំណត់ត្រាវត្តមានថ្ងៃនេះ (${todayStr}) ក្នុង Google Sheet...*`, {
    parse_mode: 'Markdown'
  }).catch(() => {});

  try {
    const res = await clearAttendanceDatabase(todayStr);
    if (res && res.status === 'error') {
      return ctx.editMessageText(
        `⚠️ *មិនអាចលុបក្នុង Google Sheet បានទេ:*\n` +
        `\`${res.message}\`\n\n` +
        `👉 *មូលហេតុ:* Google Apps Script មិនទាន់ស្គាល់មុខងារនេះទេ។ សូមលោកគ្រូចម្លងកូដពី \`google_apps_script.js\` ទៅ Paste ក្នុង Google Sheet ➔ Deploy ➔ New version សិនបាទ!`,
        { parse_mode: 'Markdown' }
      ).catch(() => {});
    }
    await ctx.editMessageText(
      `🗑 *បានលុបកំណត់ត្រាវត្តមានថ្ងៃនេះ (${todayStr}) ជោគជ័យ!* \n\n` +
      `✅ កំណត់ត្រាវត្តមានថ្ងៃនេះត្រូវបានលុបចេញពី Google Sheet។\n` +
      `✅ ទិន្នន័យថ្ងៃផ្សេងៗ និងបញ្ជីឈ្មោះសិស្សនៅរក្សាទុកដដែល។`,
      { parse_mode: 'Markdown' }
    ).catch(() => {});
  } catch (err) {
    await ctx.editMessageText(`❌ មានបញ្ហា: ${err.message}`, { parse_mode: 'Markdown' }).catch(() => {});
  }
});

bot.callbackQuery('cls_do_all_att', async (ctx) => {
  await ctx.answerCallbackQuery('⏳ កំពុងលុបកំណត់ត្រាវត្តមានទាំងអស់...');
  ctx.editMessageText('⏳ *កំពុងលុបកំណត់ត្រាវត្តមានទាំងអស់ក្នុង Google Sheet...*', {
    parse_mode: 'Markdown'
  }).catch(() => {});

  try {
    const res = await clearAttendanceDatabase();
    if (res && res.status === 'error') {
      return ctx.editMessageText(
        `⚠️ *មិនអាចលុបក្នុង Google Sheet បានទេ:*\n` +
        `\`${res.message}\`\n\n` +
        `👉 *មូលហេតុ:* Google Apps Script មិនទាន់ស្គាល់មុខងារនេះទេ។ សូមលោកគ្រូចម្លងកូដពី \`google_apps_script.js\` ទៅ Paste ក្នុង Google Sheet ➔ Deploy ➔ New version សិនបាទ!`,
        { parse_mode: 'Markdown' }
      ).catch(() => {});
    }
    const studentCount = getStudents().length;
    await ctx.editMessageText(
      `🗑 *បានលុបកំណត់ត្រាវត្តមានទាំងអស់ជោគជ័យ! (All Attendance Cleared)*\n\n` +
      `✅ សន្លឹកវត្តមាន (Attendance Tab ក្នុង Google Sheet) ត្រូវបានលុបចេញស្អាត (០ កំណត់ត្រា)។\n` +
      `✅ Memory Cache ក្នុង Bot ត្រូវបាន Reset។\n` +
      `👥 បញ្ជីឈ្មោះសិស្ស (${studentCount} នាក់ ក្នុងសន្លឹក Students) និងកាលវិភាគនៅរក្សាទុកពេញលេញ ១០០% ដដែល។`,
      { parse_mode: 'Markdown' }
    ).catch(() => {});
  } catch (err) {
    await ctx.editMessageText(`❌ មានបញ្ហា: ${err.message}`, { parse_mode: 'Markdown' }).catch(() => {});
  }
});

bot.callbackQuery('cls_do_students', async (ctx) => {
  await ctx.answerCallbackQuery('⏳ កំពុងសម្អាតបញ្ជីសិស្ស...');
  ctx.editMessageText('⏳ *កំពុងសម្អាតបញ្ជីឈ្មោះសិស្សក្នុង Google Sheet...*', {
    parse_mode: 'Markdown'
  }).catch(() => {});

  try {
    const res = await clearStudentsDatabase();
    if (res && res.status === 'error') {
      return ctx.editMessageText(
        `⚠️ *មិនអាចលុបក្នុង Google Sheet បានទេ:*\n` +
        `\`${res.message}\`\n\n` +
        `👉 *មូលហេតុ:* Google Apps Script មិនទាន់ស្គាល់មុខងារនេះទេ។ សូមលោកគ្រូចម្លងកូដពី \`google_apps_script.js\` ទៅ Paste ក្នុង Google Sheet ➔ Deploy ➔ New version សិនបាទ!`,
        { parse_mode: 'Markdown' }
      ).catch(() => {});
    }
    await ctx.editMessageText(
      `👥 *បានសម្អាតបញ្ជីឈ្មោះសិស្សក្នុង Google Sheet ជោគជ័យ! (Students Cleared)*\n\n` +
      `✅ តារាងសិស្សក្នុង Google Sheet ត្រូវបានសម្អាត។\n` +
      `👉 លោកគ្រូអាចប្រើ /addstudent ដើម្បីបញ្ចូលសិស្សថ្មីឡើងវិញ។`,
      { parse_mode: 'Markdown' }
    ).catch(() => {});
  } catch (err) {
    await ctx.editMessageText(`❌ មានបញ្ហា: ${err.message}`, { parse_mode: 'Markdown' }).catch(() => {});
  }
});

bot.callbackQuery('cls_do_reset_all', async (ctx) => {
  await ctx.answerCallbackQuery('⏳ កំពុង Reset ទាំងអស់...');
  ctx.editMessageText('⏳ *កំពុង Reset ទិន្នន័យទាំងអស់ក្នុង Google Sheet...*', {
    parse_mode: 'Markdown'
  }).catch(() => {});

  try {
    const res = await clearAllDatabase();
    if (res && res.status === 'error') {
      return ctx.editMessageText(
        `⚠️ *មិនអាច Reset Google Sheet បានទេ:*\n` +
        `\`${res.message}\`\n\n` +
        `👉 *មូលហេតុ:* Google Apps Script មិនទាន់ស្គាល់មុខងារនេះទេ។ សូមលោកគ្រូចម្លងកូដពី \`google_apps_script.js\` ទៅ Paste ក្នុង Google Sheet ➔ Deploy ➔ New version សិនបាទ!`,
        { parse_mode: 'Markdown' }
      ).catch(() => {});
    }
    await ctx.editMessageText(
      `💣 *បាន Reset ទិន្នន័យទាំងអស់ក្នុង Google Sheet ជោគជ័យ! (Reset Everything)*\n\n` +
      `✅ រាល់កំណត់ត្រាវត្តមាន និងបញ្ជីឈ្មោះសិស្សត្រូវបានសម្អាតទាំងស្រុង។`,
      { parse_mode: 'Markdown' }
    ).catch(() => {});
  } catch (err) {
    await ctx.editMessageText(`❌ មានបញ្ហា: ${err.message}`, { parse_mode: 'Markdown' }).catch(() => {});
  }
});

// /setclearpassword command (Teacher Only)
bot.command('setclearpassword', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ សិទ្ធិកំណត់ Password សម្អាត Sheet គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!');
  }

  // Delete message to hide new password
  try {
    await ctx.deleteMessage();
  } catch (e) {}

  const parts = ctx.message.text.trim().split(/\s+/).slice(1);
  const newPass = parts[0] ? parts[0].trim() : '';

  if (!newPass || newPass.length < 4) {
    return ctx.reply(
      `❌ សូមបញ្ចូល Password ថ្មី (យ៉ាងតិច ៤ តួអក្សរ)!\n` +
      `ទម្រង់: \`/setclearpassword <password_ថ្មី>\`\n` +
      `ឧទាហរណ៍: \`/setclearpassword mypass2026\``,
      { parse_mode: 'Markdown' }
    );
  }

  setClearPassword(newPass);
  await ctx.reply(
    `✅ *បានផ្លាស់ប្តូរ Password សម្អាត Sheet ជោគជ័យ!*\n` +
    `🔑 *Password ថ្មីរបស់អ្នក:* \`${newPass}\`\n\n` +
    `👉 ពេលចង់សម្អាតទិន្នន័យ សូមប្រើ: \`/cleardata ${newPass}\``,
    { parse_mode: 'Markdown' }
  );
});

// Generate and send Excel report (STRICT TEACHER ONLY)
// Prompt Date Selection for Full-Day Attendance Report
async function promptChooseReportDate(ctx) {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!* ប្រធានថ្នាក់មិនអាចទាញយកបានទេ។', {
      parse_mode: 'Markdown'
    });
  }

  const now = dayjs();
  const dayName = DAYS[now.day()];
  const todayStr = now.format('YYYY-MM-DD');
  const todayKhmer = KHMER_DAYS[dayName] || 'អាទិត្យ';

  const yesterday = now.subtract(1, 'day');
  const yesterdayStr = yesterday.format('YYYY-MM-DD');
  const yesterdayDayName = DAYS[yesterday.day()];
  const yesterdayKhmer = KHMER_DAYS[yesterdayDayName] || 'សៅរ៍';

  const keyboard = new InlineKeyboard()
    .text(`📅 របាយការណ៍ថ្ងៃនេះ (${todayKhmer}, ${now.format('DD/MM')})`, `get_fullday_${todayStr}`)
    .row()
    .text(`📅 របាយការណ៍ម្សិលមិញ (${yesterdayKhmer}, ${yesterday.format('DD/MM')})`, `get_fullday_${yesterdayStr}`)
    .row()
    .text('📅 ជ្រើសរើសថ្ងៃមុនៗ... (Past Days)', 'list_past_report_days');

  const text =
`📊 *ជ្រើសរើសរបាយការណ៍វត្តមានពេញមួយថ្ងៃ (Full Day Report):*
👉 សូមជ្រើសរើសកាលបរិច្ឆេទដែលលោកគ្រូចង់ទាញយកឯកសារ Excel ខាងក្រោម៖`;

  if (ctx.callbackQuery) {
    return ctx.editMessageText(text, { parse_mode: 'Markdown', reply_markup: keyboard });
  }
  return ctx.reply(text, { parse_mode: 'Markdown', reply_markup: keyboard });
}

// Generate and send FULL-DAY Excel report
async function sendFullDayReport(ctx, dateStr) {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!*', {
      parse_mode: 'Markdown'
    });
  }

  const dayStatus = await getDayAttendanceStatus(dateStr, { DAYS, KHMER_DAYS, getTodaySlots });

  if (!dayStatus.hasTimetable) {
    return ctx.reply(`⚠️ *ថ្ងៃ${dayStatus.khmerDay} (${dayjs(dateStr).format('DD-MM-YYYY')}) មិនមានកាលវិភាគសិក្សាទេ!*`, {
      parse_mode: 'Markdown'
    });
  }

  try {
    const { filePath, filename } = await generateFullDayExcel(dayStatus);

    // Build session status list for caption
    let sessionLines = '';
    dayStatus.scheduledSessions.forEach(s => {
      const icon = s.isRecorded ? '✅' : '⚠️';
      const statusNote = s.isRecorded
        ? `បានកត់ត្រា (ដោយ ${s.recordedBy})`
        : `*ប្រធានថ្នាក់ភ្លេចកត់ត្រា!*`;
      sessionLines += `• ⏰ \`${s.timeSlot}\` (${s.code || s.subject}): ${icon} ${statusNote}\n`;
    });

    const caption =
`📊 *របាយការណ៍វត្តមានសិស្សពេញមួយថ្ងៃ (Full Day Report)*
🏫 *ថ្នាក់:* ${CLASS_NAME}
📅 *កាលបរិច្ឆេទ:* ថ្ងៃ${dayStatus.khmerDay}, ${dayjs(dateStr).format('DD MMMM YYYY')}
👨‍🏫 *គ្រូបន្ទុកថ្នាក់:* ${TEACHER_NAME}

⏰ *ស្ថានភាពម៉ោងសិក្សា (${dayStatus.totalSessions} ម៉ោងសរុប):*
${sessionLines}
📌 *សរុប:* បានកត់ត្រា *${dayStatus.recordedCount}* ម៉ោង | ⚠️ ភ្លេចកត់ត្រា *${dayStatus.missedCount}* ម៉ោង
👥 *សិស្សសរុប:* ${dayStatus.students.length} នាក់

📁 _ឯកសារ Excel មានតារាងរួមគ្រប់មុខវិជ្ជា និងសន្លឹកកិច្ចការលម្អិតតាមម៉ោងនីមួយៗ!_`;

    const isGroup = ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
    const teacherId = getTeacherId();

    if (isGroup && teacherId) {
      try {
        await bot.api.sendDocument(teacherId, new InputFile(filePath, filename), {
          caption: `🔒 *[ផ្ញើដោយសម្ងាត់ពីគ្រុប]*\n\n` + caption,
          parse_mode: 'Markdown'
        });

        await ctx.reply(
          `🔒 *របាយការណ៍ Excel ពេញមួយថ្ងៃ ត្រូវបានផ្ញើចូលក្នុងឆាតផ្ទាល់ខ្លួន (Private Chat) របស់លោកគ្រូរួចរាល់ហើយ!*\n_ប្រធានថ្នាក់ក្នុងគ្រុបនេះ មិនអាចមើល ឬទាញយកឯកសារនេះបានទេ។_`,
          { parse_mode: 'Markdown' }
        );
        return;
      } catch (dmErr) {
        console.warn('Could not send DM to teacher:', dmErr.message);
      }
    }

    await ctx.replyWithDocument(new InputFile(filePath, filename), {
      caption,
      parse_mode: 'Markdown'
    });
  } catch (err) {
    console.error('Error generating Full Day Excel:', err);
    await ctx.reply(`❌ បរាជ័យក្នុងការបង្កើតរបាយការណ៍: ${err.message}`);
  }
}

// Helper: Parse user date input in DD-MM-YYYY, DD/MM/YYYY, or YYYY-MM-DD
function parseUserDateInput(input) {
  if (!input) return null;
  const str = input.trim().toLowerCase();
  const now = dayjs();

  if (/^(yesterday|ម្សិលមិញ)$/.test(str)) {
    return now.subtract(1, 'day').format('YYYY-MM-DD');
  }
  if (/^(today|ថ្ងៃនេះ)$/.test(str)) {
    return now.format('YYYY-MM-DD');
  }

  // DD-MM-YYYY or DD/MM/YYYY (e.g. 01-10-2026, 31/10/2026)
  const ddmmyyyy = str.match(/^(\d{1,2})[\-\/\.](\d{1,2})[\-\/\.](\d{4})$/);
  if (ddmmyyyy) {
    const day = ddmmyyyy[1].padStart(2, '0');
    const month = ddmmyyyy[2].padStart(2, '0');
    const year = ddmmyyyy[3];
    return `${year}-${month}-${day}`;
  }

  // YYYY-MM-DD or YYYY/MM/DD (e.g. 2026-10-01)
  const yyyymmdd = str.match(/^(\d{4})[\-\/\.](\d{1,2})[\-\/\.](\d{1,2})$/);
  if (yyyymmdd) {
    const year = yyyymmdd[1];
    const month = yyyymmdd[2].padStart(2, '0');
    const day = yyyymmdd[3].padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  return null;
}

// Command: /attendance [optional date e.g. 01-10-2026 or yesterday]
bot.command('attendance', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!*', {
      parse_mode: 'Markdown'
    });
  }

  const parts = ctx.message.text.trim().split(/\s+/).slice(1);
  const arg = parts.join(' ');

  if (arg) {
    const parsedDate = parseUserDateInput(arg);
    if (parsedDate) {
      return sendFullDayReport(ctx, parsedDate);
    } else {
      await ctx.reply(
        `⚠️ *ទម្រង់កាលបរិច្ឆេទមិនត្រឹមត្រូវ!*\n` +
        `សូមវាយតាមទម្រង់: \`DD-MM-YYYY\` (ឧ. \`/attendance 01-10-2026\`) ឬ \`yesterday\`\n` +
        `Bot នឹងបើកផ្ទាំងជ្រើសរើសថ្ងៃជូនលោកគ្រូស្វ័យប្រវត្តិ៖`,
        { parse_mode: 'Markdown' }
      );
    }
  }

  // Default: Open interactive date picker!
  await promptChooseReportDate(ctx);
});

// Callbacks for report
bot.callbackQuery('cmd_attendance', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.answerCallbackQuery({
      text: '⚠️ សិទ្ធិទាញយករបាយការណ៍ សម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!',
      show_alert: true
    });
  }
  await ctx.answerCallbackQuery();
  await promptChooseReportDate(ctx);
});

bot.callbackQuery(/^get_fullday_(.+)$/, async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.answerCallbackQuery({
      text: '⚠️ សិទ្ធិទាញយករបាយការណ៍ សម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!',
      show_alert: true
    });
  }
  await ctx.answerCallbackQuery();
  const dateStr = ctx.match[1];
  await sendFullDayReport(ctx, dateStr);
});

bot.callbackQuery('list_past_report_days', async (ctx) => {
  await ctx.answerCallbackQuery();
  const now = dayjs();
  const keyboard = new InlineKeyboard();
  for (let i = 1; i <= 6; i++) {
    const pastD = now.subtract(i, 'day');
    const pStr = pastD.format('YYYY-MM-DD');
    const pDayName = DAYS[pastD.day()];
    const pKhmer = KHMER_DAYS[pDayName];
    if (pastD.day() !== 0) { // skip Sunday
      keyboard.text(`📅 ថ្ងៃ${pKhmer} (${pastD.format('DD/MM/YYYY')})`, `get_fullday_${pStr}`).row();
    }
  }
  keyboard.text('🔙 ថយក្រោយ (Back)', 'cmd_attendance');
  await ctx.editMessageText('📅 *សូមជ្រើសរើសថ្ងៃដែលចង់ទាញយករបាយការណ៍:*', {
    parse_mode: 'Markdown',
    reply_markup: keyboard
  });
});

// --- MONTHLY & CUSTOM DATE RANGE REPORT FUNCTIONS ---

async function sendRangeReport(ctx, startDateStr, endDateStr, customTitle = null, customFileName = null) {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!*', {
      parse_mode: 'Markdown'
    });
  }

  const rangeData = await getRangeAttendanceStatus(startDateStr, endDateStr, { DAYS, KHMER_DAYS, getTodaySlots });

  if (rangeData.totalSchoolDays === 0) {
    return ctx.reply(
      `⚠️ *មិនមានកាលវិភាគសិក្សាក្នុងចន្លោះកាលបរិច្ឆេទនេះទេ!* (${dayjs(startDateStr).format('DD/MM/YYYY')} ដល់ ${dayjs(endDateStr).format('DD/MM/YYYY')})`,
      { parse_mode: 'Markdown' }
    );
  }

  try {
    await ctx.replyWithChatAction?.('upload_document');
    const { filePath, filename } = await generateRangeExcelReport({
      ...rangeData,
      customTitle,
      customFileName
    });

    let totalClassA = 0;
    let totalClassP = 0;
    let totalClassL = 0;
    rangeData.studentSummaries.forEach(st => {
      totalClassA += st.absent;
      totalClassP += st.permission;
      totalClassL += st.late;
    });

    const displayTitle = customTitle || 'របាយការណ៍វត្តមានសិស្សសរុប (Monthly Report)';
    const caption =
`📊 *${displayTitle}*
🏫 *ថ្នាក់:* ${CLASS_NAME}
📁 *ឈ្មោះឯកសារ:* \`${filename}\`
📅 *ចន្លោះកាលបរិច្ឆេទ:* ពី \`${dayjs(startDateStr).format('DD/MM/YYYY')}\` ដល់ \`${dayjs(endDateStr).format('DD/MM/YYYY')}\`
👨‍🏫 *គ្រូបន្ទុកថ្នាក់:* ${TEACHER_NAME}

📈 *ស្ថិតិសរុបទូទាំងថ្នាក់:*
• 🏫 ថ្ងៃរៀនសរុប: *${rangeData.totalSchoolDays} ថ្ងៃ*
• ⏰ ម៉ោងបង្រៀនសរុប: *${rangeData.totalScheduledSessions} ម៉ោង* (កត់ត្រាបាន ${rangeData.totalRecordedSessions}, ភ្លេច ${rangeData.totalMissedSessions})
• 👥 សិស្សសរុប: *${rangeData.studentSummaries.length} នាក់*
• ❌ អវត្តមានសរុប (A): *${totalClassA} លើក*
• 📝 ច្បាប់សរុប (P): *${totalClassP} លើក*
• ⏳ យឺតសរុប (L): *${totalClassL} លើក*

📁 *ឯកសារ Excel មាន ៣ សន្លឹកកិច្ចការ:*
1️⃣ សង្ខេបប្រចាំខែ និងការវាយតម្លៃវិន័យសិស្សម្នាក់ៗ
2️⃣ តារាងវត្តមានតាមថ្ងៃពេញមួយខែ (Daily Attendance Grid)
3️⃣ ស្ថិតិវត្តមានតាមមុខវិជ្ជានីមួយៗ`;

    const isGroup = ctx.chat.type === 'group' || ctx.chat.type === 'supergroup';
    const teacherId = getTeacherId();

    if (isGroup && teacherId) {
      try {
        await bot.api.sendDocument(teacherId, new InputFile(filePath, filename), {
          caption: `🔒 *[ផ្ញើដោយសម្ងាត់ពីគ្រុប]*\n\n` + caption,
          parse_mode: 'Markdown'
        });

        await ctx.reply(
          `🔒 *របាយការណ៍ប្រចាំខែ Excel ត្រូវបានផ្ញើចូលក្នុងឆាតផ្ទាល់ខ្លួន (Private Chat) របស់លោកគ្រូរួចរាល់ហើយ!*\n_ប្រធានថ្នាក់ក្នុងគ្រុបនេះ មិនអាចមើល ឬទាញយកឯកសារនេះបានទេ។_`,
          { parse_mode: 'Markdown' }
        );
        return;
      } catch (dmErr) {
        console.warn('Could not send DM to teacher:', dmErr.message);
      }
    }

    await ctx.replyWithDocument(new InputFile(filePath, filename), {
      caption,
      parse_mode: 'Markdown'
    });
  } catch (err) {
    console.error('Error generating Range Excel:', err);
    await ctx.reply(`❌ បរាជ័យក្នុងការបង្កើតរបាយការណ៍ប្រចាំខែ: ${err.message}`);
  }
}

// Prompt teacher to rename file when selecting date-to-date
async function promptRangeRename(ctx, startDateStr, endDateStr) {
  pendingRangeRename.set(ctx.from.id, { startDateStr, endDateStr });
  const startFmt = dayjs(startDateStr).format('DD-MM-YYYY');
  const endFmt = dayjs(endDateStr).format('DD-MM-YYYY');

  const keyboard = new InlineKeyboard()
    .text('⚡ ប្រើឈ្មោះស្វ័យប្រវត្តិ (Skip)', 'range_skip_rename')
    .row()
    .text('❌ បោះបង់ (Cancel)', 'range_cancel_rename');

  await ctx.reply(
    `📅 *បានកំណត់ចន្លោះថ្ងៃ:* ពី \`${startFmt}\` ដល់ \`${endFmt}\`\n\n` +
    `✍️ *សូមវាយ "ឈ្មោះឯកសារ" ឬ "ចំណងជើងរបាយការណ៍" ដែលលោកគ្រូចង់ដាក់:*\n` +
    `_(ឧទាហរណ៍: \`របាយការណ៍ពាក់កណ្តាលខែ១០_១១\` ឬ \`វត្តមាន_ខែវិច្ឆិកា_2026\`)_\n\n` +
    `👉 ឬចុចប៊ូតុង *[ប្រើឈ្មោះស្វ័យប្រវត្តិ]* ប្រសិនបើចង់ប្រើឈ្មោះស្វ័យប្រវត្តិតាមកាលបរិច្ឆេទ។`,
    { parse_mode: 'Markdown', reply_markup: keyboard }
  );
}

async function promptMonthlyReport(ctx) {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!*', {
      parse_mode: 'Markdown'
    });
  }

  const now = dayjs();
  const thisMonthStart = now.startOf('month').format('YYYY-MM-DD');
  const thisMonthEnd = now.endOf('month').format('YYYY-MM-DD');
  const thisMonthLabel = `ខែនេះ (ខែ ${now.format('MM/YYYY')})`;

  const lastMonth = now.subtract(1, 'month');
  const lastMonthStart = lastMonth.startOf('month').format('YYYY-MM-DD');
  const lastMonthEnd = lastMonth.endOf('month').format('YYYY-MM-DD');
  const lastMonthLabel = `ខែមុន (ខែ ${lastMonth.format('MM/YYYY')})`;

  const keyboard = new InlineKeyboard()
    .text(`📆 ${thisMonthLabel}`, `get_range_${thisMonthStart}_${thisMonthEnd}`)
    .row()
    .text(`📆 ${lastMonthLabel}`, `get_range_${lastMonthStart}_${lastMonthEnd}`)
    .row()
    .text('📅 កំណត់ចន្លោះថ្ងៃដោយខ្លួនឯង (Custom Range)', 'prompt_custom_range')
    .row()
    .text('❌ បោះបង់', 'rec_cancel');

  const text =
`📆 *របាយការណ៍វត្តមានសង្ខេបប្រចាំខែ (Monthly Report)*

👉 សូមជ្រើសរើសខែដែលលោកគ្រូចង់ទាញយករបាយការណ៍ Excel ខាងក្រោម៖
_(ឬវាយបញ្ជាកំណត់ថ្ងៃដោយខ្លួនឯង: \`/range 01-10-2026 to 31-10-2026\`)_`;

  if (ctx.callbackQuery) {
    return ctx.editMessageText(text, { parse_mode: 'Markdown', reply_markup: keyboard });
  }
  return ctx.reply(text, { parse_mode: 'Markdown', reply_markup: keyboard });
}

// Command: /monthly
bot.command('monthly', promptMonthlyReport);
bot.command('month', promptMonthlyReport);

// Command: /range <start> to <end> [custom name]
bot.command('range', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!*', {
      parse_mode: 'Markdown'
    });
  }

  const text = ctx.message.text.trim();
  // Match: /range DD-MM-YYYY to DD-MM-YYYY [custom name]
  const match = text.match(/^\/range\s+(\d{1,2}[\-\/\.]\d{1,2}[\-\/\.]\d{4})\s*(?:to|\-|ដល់)\s*(\d{1,2}[\-\/\.]\d{1,2}[\-\/\.]\d{4})(?:\s+(.+))?$/i);

  if (!match) {
    return ctx.reply(
      `❌ *ទម្រង់មិនត្រឹមត្រូវ!*\n` +
      `របៀបប្រើ: \`/range <ថ្ងៃចាប់ផ្តើម> to <ថ្ងៃបញ្ចប់> [ឈ្មោះឯកសារ]\`\n` +
      `*ឧទាហរណ៍:*\n` +
      `• \`/range 15-10-2026 to 14-11-2026\`\n` +
      `• \`/range 15-10-2026 to 14-11-2026 របាយការណ៍ខែវិច្ឆិកា\``,
      { parse_mode: 'Markdown' }
    );
  }

  const startParsed = parseUserDateInput(match[1]);
  const endParsed = parseUserDateInput(match[2]);
  const customName = match[3] ? match[3].trim() : null;

  if (!startParsed || !endParsed) {
    return ctx.reply(
      `❌ *ថ្ងៃខែមិនត្រឹមត្រូវ!*\nសូមវាយតាមទម្រង់: \`DD-MM-YYYY\` (ឧ. \`/range 15-10-2026 to 14-11-2026\`)`,
      { parse_mode: 'Markdown' }
    );
  }

  if (dayjs(startParsed).isAfter(dayjs(endParsed))) {
    return ctx.reply('❌ ថ្ងៃចាប់ផ្តើមមិនអាចនៅក្រោយថ្ងៃបញ្ចប់បានទេ!', { parse_mode: 'Markdown' });
  }

  if (customName) {
    return sendRangeReport(ctx, startParsed, endParsed, customName, customName);
  }

  await promptRangeRename(ctx, startParsed, endParsed);
});

// Callbacks for monthly / range
bot.callbackQuery('cmd_monthly', async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.answerCallbackQuery({
      text: '⚠️ សិទ្ធិទាញយករបាយការណ៍ សម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!',
      show_alert: true
    });
  }
  await ctx.answerCallbackQuery();
  await promptMonthlyReport(ctx);
});

bot.callbackQuery(/^get_range_([^_]+)_(.+)$/, async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.answerCallbackQuery({
      text: '⚠️ សិទ្ធិទាញយករបាយការណ៍ សម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!',
      show_alert: true
    });
  }
  await ctx.answerCallbackQuery();
  const startStr = ctx.match[1];
  const endStr = ctx.match[2];
  const monthTitle = `របាយការណ៍វត្តមានសិស្សប្រចាំខែ ${dayjs(startStr).format('MM/YYYY')}`;
  const monthFileName = `Attendance_Report_Month_${dayjs(startStr).format('MM_YYYY')}.xlsx`;
  await sendRangeReport(ctx, startStr, endStr, monthTitle, monthFileName);
});

bot.callbackQuery('range_skip_rename', async (ctx) => {
  if (!checkIsTeacher(ctx)) return ctx.answerCallbackQuery();
  await ctx.answerCallbackQuery();
  const pending = pendingRangeRename.get(ctx.from.id);
  if (!pending) {
    return ctx.reply('⚠️ មិនមានប្រតិបត្តិការកំណត់ឈ្មោះដែលកំពុងរង់ចាំទេ។');
  }
  pendingRangeRename.delete(ctx.from.id);
  await sendRangeReport(ctx, pending.startDateStr, pending.endDateStr, null, null);
});

bot.callbackQuery('range_cancel_rename', async (ctx) => {
  await ctx.answerCallbackQuery();
  pendingRangeRename.delete(ctx.from.id);
  await ctx.editMessageText('❌ បានបោះបង់ការទាញយករបាយការណ៍។');
});

bot.callbackQuery('prompt_custom_range', async (ctx) => {
  await ctx.answerCallbackQuery();
  await ctx.reply(
    `📅 *របៀបទាញយករបាយការណ៍តាមចន្លោះថ្ងៃ (Custom Date to Date):*\n\n` +
    `លោកគ្រូអាចវាយកាលបរិច្ឆេទ ពីថ្ងៃណា ដល់ថ្ងៃណា ឧទាហរណ៍៖\n` +
    `• \`15-10-2026 to 14-11-2026\`\n` +
    `• \`/range 15-10-2026 to 14-11-2026\`\n` +
    `• \`/range 15-10-2026 to 14-11-2026 របាយការណ៍ខែវិច្ឆិកា\`\n\n` +
    `💡 _បន្ទាប់ពីវាយចន្លោះថ្ងៃ Bot នឹងសួររក "ឈ្មោះឯកសារ" ដែលលោកគ្រូចង់ដាក់ភ្លាមៗ!_`,
    { parse_mode: 'Markdown' }
  );
});

bot.callbackQuery('cmd_today', async (ctx) => {
  await ctx.answerCallbackQuery();
  const now = dayjs();
  const dayName = DAYS[now.day()];
  const khmerDay = KHMER_DAYS[dayName];
  const slots = getTodaySlots(now);
  let text = `📅 *កាលវិភាគថ្ងៃ${khmerDay} (${dayName}):*\n\n`;
  if (slots.length === 0) text += 'មិនមានម៉ោងសិក្សាទេ។';
  else {
    slots.forEach(s => text += `⏰ \`${s.start} - ${s.end}\`: *${s.subject}*\n`);
  }
  await ctx.reply(text, { parse_mode: 'Markdown' });
});

bot.callbackQuery('cmd_students', async (ctx) => {
  await ctx.answerCallbackQuery();
  const students = getStudents();
  let text = `👥 *បញ្ជីឈ្មោះសិស្ស (${students.length} នាក់):*\n`;
  students.forEach((s) => {
    text += `${s.id}. ${s.name} (${s.gender})\n`;
  });
  await ctx.reply(text, { parse_mode: 'Markdown' });
});

bot.callbackQuery(/^get_report_(.+)$/, async (ctx) => {
  if (!checkIsTeacher(ctx)) {
    return ctx.answerCallbackQuery({
      text: '⚠️ សិទ្ធិទាញយករបាយការណ៍ សម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!',
      show_alert: true
    });
  }
  await ctx.answerCallbackQuery();
  const todayStr = dayjs().format('YYYY-MM-DD');
  await sendFullDayReport(ctx, todayStr);
});

// Message listener for attendance or keywords
bot.on('message:text', async (ctx, next) => {
  const text = ctx.message.text.trim();

  // Handle pending range rename for teacher
  if (pendingRangeRename.has(ctx.from.id)) {
    const pending = pendingRangeRename.get(ctx.from.id);
    if (/^\/(cancel|exit|stop)$/i.test(text) || text === 'បោះបង់') {
      pendingRangeRename.delete(ctx.from.id);
      return ctx.reply('❌ បានបោះបង់ការទាញយករបាយការណ៍។');
    }
    if (/^\/(skip|default)$/i.test(text) || /^(skip|រំលង|ធម្មតា)$/i.test(text)) {
      pendingRangeRename.delete(ctx.from.id);
      return sendRangeReport(ctx, pending.startDateStr, pending.endDateStr, null, null);
    }
    pendingRangeRename.delete(ctx.from.id);
    return sendRangeReport(ctx, pending.startDateStr, pending.endDateStr, text, text);
  }

  // Cancel command for active session
  if (/^\/(cancel|exit|stop)$/i.test(text) || text === 'បោះបង់') {
    if (pendingSessions.has(ctx.from.id)) {
      pendingSessions.delete(ctx.from.id);
      return ctx.reply('❌ បានបោះបង់ការកត់ត្រាវត្តមាន។');
    }
  }

  // Keyword check: attendance / report / វត្តមាន [optional date]
  const matchAtt = text.match(/^(?:attendance|report|វត្តមាន|\/attendance|\/report|\/វត្តមាន)(?:\s+(.+))?$/i);
  if (matchAtt) {
    const rawArg = matchAtt[1];
    if (rawArg) {
      const parsedDate = parseUserDateInput(rawArg);
      if (parsedDate) {
        return sendFullDayReport(ctx, parsedDate);
      }
    }
    return promptChooseReportDate(ctx);
  }

  // Clear chat text keywords: e.g. "clean", "clear", "សម្អាតឆាត", "លុបសារ"
  const cleanMatch = text.match(/^(?:\/clean|\/clear|\/clearchat|clean|clear|clearchat|សម្អាត|សម្អាតឆាត|លុបសារ)(?:\s+(\d+))?$/i);
  if (cleanMatch) {
    const count = cleanMatch[1] ? parseInt(cleanMatch[1], 10) : 30;
    return cleanChatMessages(ctx, count);
  }

  // Monthly keyword check
  if (/^(monthly|\/monthly|របាយការណ៍ខែ|ខែនេះ|ខែមុន)$/i.test(text)) {
    return promptMonthlyReport(ctx);
  }

  // Cross-month or custom date range matching: e.g. "15-10-2026 to 14-11-2026" or "15-10-2026 to 14-11-2026 របាយការណ៍ខែវិច្ឆិកា"
  const rangeMatch = text.match(/^(?:\/range\s+|range\s+)?(\d{1,2}[\-\/\.]\d{1,2}[\-\/\.]\d{4})\s*(?:to|\-|ដល់)\s*(\d{1,2}[\-\/\.]\d{1,2}[\-\/\.]\d{4})(?:\s+(.+))?$/i);
  if (rangeMatch) {
    if (!checkIsTeacher(ctx)) {
      return ctx.reply('⚠️ *សិទ្ធិទាញយករបាយការណ៍វត្តមាន គឺសម្រាប់តែលោកគ្រូប៉ុណ្ណោះ!*', {
        parse_mode: 'Markdown'
      });
    }
    const startParsed = parseUserDateInput(rangeMatch[1]);
    const endParsed = parseUserDateInput(rangeMatch[2]);
    const customName = rangeMatch[3] ? rangeMatch[3].trim() : null;

    if (startParsed && endParsed) {
      if (dayjs(startParsed).isAfter(dayjs(endParsed))) {
        return ctx.reply('❌ ថ្ងៃចាប់ផ្តើមមិនអាចនៅក្រោយថ្ងៃបញ្ចប់បានទេ!', { parse_mode: 'Markdown' });
      }
      if (customName) {
        return sendRangeReport(ctx, startParsed, endParsed, customName, customName);
      }
      return promptRangeRename(ctx, startParsed, endParsed);
    }
  }

  // Check if sender has an active session from interactive button selection
  const session = pendingSessions.get(ctx.from.id);

  // Parse attendance input: Name P / Name A / Name L
  const { parsedItems } = parseAttendanceInput(text);

  if (parsedItems.length > 0) {
    const now = dayjs();
    let dateStr, subject, timeSlot;

    if (session) {
      // Use chosen date, time slot, and subject from interactive button session!
      dateStr = session.dateStr;
      subject = session.subject;
      timeSlot = session.timeSlot;
      pendingSessions.delete(ctx.from.id); // Clear session
    } else {
      // Direct message fallback: Auto-detect from real-time timetable
      const currentSlot = getCurrentSlot(now);
      dateStr = now.format('YYYY-MM-DD');
      subject = currentSlot ? currentSlot.subject : 'មុខវិជ្ជាទូទៅ';
      timeSlot = currentSlot ? `${currentSlot.start} - ${currentSlot.end}` : now.format('HH:mm');
    }

    const recordedBy = ctx.from?.first_name || 'ប្រធានថ្នាក់';

    const result = await recordAttendance({
      dateStr,
      subject,
      timeSlot,
      recordedBy,
      parsedItems
    });

    if (result.status === 'no_students') {
      return ctx.reply(
        `⚠️ *បច្ចុប្បន្នមិនទាន់មានឈ្មោះសិស្សក្នុងប្រព័ន្ធទេ (០ នាក់)!*\n\n` +
        `ដោយសារលោកគ្រូបាន Reset / Clear ទិន្នន័យបញ្ជីសិស្សក្នុង Google Sheet រួចរាល់ ដូច្នេះមិនទាន់អាចកត់ត្រាវត្តមានបានឡើយ។\n\n` +
        `👉 *ដើម្បីចាប់ផ្តើមកត់វត្តមាន សូមលោកគ្រូ:*\n` +
        `1️⃣ បើក **Google Sheet** ➡️ ចូលសន្លឹក **"Students"** ➡️ បញ្ចូលឈ្មោះសិស្សពិតប្រាកដ (No, ID, Name, Gender, DOB)\n` +
        `2️⃣ ឬវាយពាក្យបញ្ជាក្នុង Bot: \`/addstudent <លេខរៀង> <ឈ្មោះ> <ភេទ>\`\n` +
        `3️⃣ ឬប្រសិនបើលោកគ្រូចង់ទាញយកបញ្ជីសិស្សគំរូ ៣០ នាក់មកវិញ សូមវាយ: \`/restorestudents\``,
        { parse_mode: 'Markdown' }
      );
    }

    const sum = result.summary;
    let replyMsg =
`✅ *បានកត់ត្រាវត្តមានដោយជោគជ័យ!*
🏫 *ថ្នាក់:* ${CLASS_NAME}
📚 *មុខវិជ្ជា:* ${subject} (\`${timeSlot}\`)
📅 *កាលបរិច្ឆេទ:* ${dayjs(dateStr).format('dddd, DD MMMM YYYY')}
👤 *អ្នកកត់ត្រា:* ${recordedBy}

📊 *ស្ថិតិវត្តមាន:*
• 👥 សិស្សសរុប: *${sum.total} នាក់*
• ✅ មានវត្តមាន: *${sum.present} នាក់*
• ❌ អវត្តមាន (A): *${sum.absent} នាក់*
• 📝 ច្បាប់ (P): *${sum.permission} នាក់*
• ⏳ មកយឺត (L): *${sum.late} នាក់*`;

    if (result.notFoundNames.length > 0) {
      replyMsg += `\n\n⚠️ *រកមិនឃើញឈ្មោះខាងក្រោមក្នុងបញ្ជី:*
${result.notFoundNames.map(n => `• _${n}_`).join('\n')}`;
    }

    const keyboard = new InlineKeyboard()
      .text('➕ កត់ម៉ោងផ្សេងទៀតក្នុងថ្ងៃនេះ', `rec_date_${dateStr}`);
    if (checkIsTeacher(ctx)) {
      keyboard.row().text('📥 ទាញយករបាយការណ៍ (Excel)', `get_fullday_${dateStr}`);
    }

    await ctx.reply(replyMsg, {
      parse_mode: 'Markdown',
      reply_markup: keyboard
    });
    return;
  }

  // If in active session but message didn't contain attendance format
  if (session) {
    return ctx.reply(
      `⚠️ *ទម្រង់មិនត្រឹមត្រូវ!*\n` +
      `សូមវាយឈ្មោះ ឬលេខរៀងសិស្សជាមួយ \`A\` (អវត្តមាន), \`P\` (ច្បាប់), ឬ \`L\` (យឺត)។\n\n` +
      `*ឧទាហរណ៍:*\n` +
      `\`ឆោម ចិន A\`\n` +
      `\`06 L\`\n` +
      `\`ខា សិនហួ P (ឈឺ)\`\n\n` +
      `_(វាយ \`/cancel\` ដើម្បីបោះបង់)_`,
      { parse_mode: 'Markdown' }
    );
  }

  await next();
});

bot.catch((err) => {
  console.error('Error in bot operation:', err);
});

module.exports = bot;

if (!process.env.VERCEL && require.main === module && process.env.POLLING === 'true') {
  console.log('🤖 Starting PineApple Attendance Bot...');
  bot.start({
    onStart: (botInfo) => {
      console.log(`✅ Bot @${botInfo.username} is running with Interactive Button Attendance Flow!`);
    }
  });
}

