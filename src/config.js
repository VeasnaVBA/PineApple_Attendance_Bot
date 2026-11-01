const path = require('path');
const fs = require('fs');
require('dotenv').config();

const DATA_DIR = path.join(__dirname, '..', 'data');
const RECORDS_DIR = path.join(DATA_DIR, 'records');
const TEMP_DIR = path.join(__dirname, '..', 'temp');

if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
if (!fs.existsSync(RECORDS_DIR)) fs.mkdirSync(RECORDS_DIR, { recursive: true });
if (!fs.existsSync(TEMP_DIR)) fs.mkdirSync(TEMP_DIR, { recursive: true });

function readJson(filename, defaultVal = []) {
  const filePath = path.join(DATA_DIR, filename);
  try {
    if (fs.existsSync(filePath)) {
      const data = fs.readFileSync(filePath, 'utf8');
      return JSON.parse(data);
    }
  } catch (err) {
    console.error(`Error reading ${filename}:`, err);
  }
  return defaultVal;
}

function writeJson(filename, data) {
  const filePath = path.join(DATA_DIR, filename);
  fs.writeFileSync(filePath, JSON.stringify(data, null, 2), 'utf8');
}

function getSettings() {
  const defaults = {
    teacherId: process.env.TEACHER_ID || null,
    teacherName: process.env.TEACHER_NAME || 'សក់ វាសនា',
    className: process.env.CLASS_NAME || 'វិទ្យាល័យ នេត យ៉ង់ - ថ្នាក់ទី 10D',
    sendReportToPrivateOnly: true // If true, always sends Excel report to Teacher PM, even if requested in group
  };
  return readJson('settings.json', defaults);
}

function saveSettings(settings) {
  writeJson('settings.json', settings);
}

function getTeacherId() {
  const s = getSettings();
  return s.teacherId ? String(s.teacherId) : null;
}

function setTeacherId(id, name) {
  const s = getSettings();
  s.teacherId = String(id);
  if (name) s.teacherName = name;
  saveSettings(s);
  return s;
}

module.exports = {
  BOT_TOKEN: process.env.BOT_TOKEN,
  TIMEZONE: process.env.TIMEZONE || 'Asia/Phnom_Penh',
  DATA_DIR,
  RECORDS_DIR,
  TEMP_DIR,
  readJson,
  writeJson,
  getSettings,
  saveSettings,
  getTeacherId,
  setTeacherId,
  get CLASS_NAME() { return getSettings().className; },
  get TEACHER_NAME() { return getSettings().teacherName; }
};
