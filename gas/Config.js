/**
 * ค่าคงที่ โครงสร้างตาราง และค่าตั้งค่าของระบบ
 *
 * Script Properties (ตั้งผ่านเมนู "SearchDocSystem → 2. ตั้งค่า" ใน Google Sheet — ห้ามเขียนลงโค้ด):
 *   SOURCE             แหล่งเอกสาร (ดู Source.js):
 *                        ลิงก์โฟลเดอร์ Google Drive หรือ ID ของโฟลเดอร์ — ไม่ต้องใช้ค่าอื่นเพิ่ม
 *                        ลิงก์แชร์ OneDrive, URL จาก browser (id=/personal/...) หรือ อีเมลเจ้าของ:/path
 *                        — ต้องใช้ TENANT_ID, CLIENT_ID, CLIENT_SECRET ด้วย (Graph.js)
 *   TENANT_ID, CLIENT_ID, CLIENT_SECRET   App ใน Microsoft Entra ID (เฉพาะ OneDrive)
 *   SITE_PASSWORD      (ไม่บังคับ) รหัสเข้าหน้าเว็บค้นหา — เว้นว่าง = ใครมีลิงก์เว็บก็ค้นได้
 *   SPREADSHEET_ID, DRIVE_ROOT_ID, TEXT_FOLDER_ID, INDEX_FOLDER_ID, TMP_FOLDER_ID,
 *   SOURCE_DRIVE_ID, SOURCE_ITEM_ID, INDEX_MANIFEST, LAST_SYNC  — ระบบตั้งให้เอง
 */

const APP_NAME = 'SearchDocSystem';
const MAX_RUNTIME_MS = 4 * 60 * 1000; // GAS ตัดที่ 6 นาที เผื่อเวลาสร้าง index และบันทึกข้อมูล
const MAX_ATTEMPTS = 3;

const SCHEMA = {
  Files: ['id', 'name', 'ext', 'path', 'server_url', 'web_url', 'size', 'modified_at', 'modified_by', 'change_key',
    'status', 'indexed_key', 'text_file_id', 'text_chars', 'attempts', 'last_error', 'first_seen_at', 'indexed_at'],
  Settings: ['key', 'value', 'description'],
  SearchLog: ['time', 'query', 'results'],
  Logs: ['time', 'level', 'message'],
};

const DEFAULT_SETTINGS = [
  ['sync_interval_minutes', '30', 'รอบการ sync อัตโนมัติ (5, 10, 15 หรือ 30 นาที) — รันเมนู "ติดตั้ง / อัปเดตระบบ" ใหม่หลังแก้'],
  ['file_extensions', 'pdf,docx,doc,xlsx,xls,csv,pptx,ppt,txt,gdoc,gsheet,gslides',
    'นามสกุลที่แตกเนื้อหามาค้นหา (gdoc/gsheet/gslides = ไฟล์ Google) ไฟล์อื่นค้นได้จากชื่ออย่างเดียว'],
  ['ocr_language', 'th', 'ภาษาของ OCR สำหรับ PDF'],
  ['max_file_mb', '40', 'ไฟล์ใหญ่กว่านี้ (MB) ไม่แตกเนื้อหา ค้นได้จากชื่อเท่านั้น'],
  ['max_text_chars', '20000', 'จำนวนตัวอักษรของเนื้อหาต่อไฟล์ที่ใส่ใน search index'],
];

/** สถานะไฟล์ในตาราง Files */
const STATUS = {
  QUEUED: 'queued',    // รอแตกเนื้อหา
  INDEXED: 'indexed',  // แตกเนื้อหาแล้ว
  NAME_ONLY: 'name_only', // ค้นได้จากชื่อเท่านั้น (ชนิดไฟล์ไม่รองรับ / ใหญ่เกิน)
  ERROR: 'error',      // แตกเนื้อหาไม่สำเร็จครบ MAX_ATTEMPTS ครั้ง
  DELETED: 'deleted',  // ไม่พบใน OneDrive แล้ว
};

function props_() {
  return PropertiesService.getScriptProperties();
}

function getSpreadsheet_() {
  const id = props_().getProperty('SPREADSHEET_ID');
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ไม่พบ Google Sheet: เปิด Apps Script จากเมนู ส่วนขยาย → Apps Script ของ Sheet แล้วรัน "ติดตั้ง / อัปเดตระบบ"');
  return ss;
}

let settingsCache_ = null;

function getSettings_() {
  if (settingsCache_) return settingsCache_;
  const raw = {};
  DEFAULT_SETTINGS.forEach(([k, v]) => (raw[k] = v));
  const sheet = getSpreadsheet_().getSheetByName('Settings');
  if (sheet && sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues()
      .forEach(([k, v]) => { if (k && String(v).trim()) raw[String(k).trim()] = String(v).trim(); });
  }
  const interval = Number(raw.sync_interval_minutes);
  settingsCache_ = {
    syncIntervalMinutes: [5, 10, 15, 30].includes(interval) ? interval : 30,
    fileExtensions: (raw.file_extensions === 'pdf,docx,doc,xlsx,xls,csv,pptx,ppt,txt' // ค่าเริ่มต้นรุ่นแรก → เพิ่มไฟล์ Google
      ? DEFAULT_SETTINGS[1][1] : raw.file_extensions).split(',').map((s) => s.trim().toLowerCase().replace(/^\./, '')).filter(Boolean),
    ocrLanguage: raw.ocr_language || 'th',
    maxFileBytes: (Number(raw.max_file_mb) || 40) * 1024 * 1024,
    maxTextChars: Number(raw.max_text_chars) || 20000,
  };
  return settingsCache_;
}

function nowIso_() {
  return new Date().toISOString();
}

/**
 * ฟังก์ชันที่ชื่อไม่ลงท้ายด้วย _ ถูกเรียกจากหน้าเว็บผ่าน google.script.run ได้
 * ฟังก์ชันสำหรับผู้ดูแล (เมนูใน Sheet / editor) จึงต้องเรียก requireOwner_() ก่อนเสมอ
 */
function requireOwner_() {
  const active = Session.getActiveUser().getEmail();
  if (!active || active !== Session.getEffectiveUser().getEmail()) {
    throw new Error('ฟังก์ชันนี้ใช้ได้จากเมนูใน Google Sheet หรือ Apps Script editor เท่านั้น');
  }
}

/**
 * ทำให้ข้อความค้นหาง่าย: Unicode NFC, ตัวเล็ก, เลขไทย → อารบิก, ตัด zero-width, ยุบช่องว่าง
 * หน้าเว็บใช้ฟังก์ชันเดียวกันนี้ (ui/Index.html) — แก้ต้องแก้ทั้งสองที่
 */
function normalizeText_(s) {
  return String(s || '')
    .normalize('NFC')
    .toLowerCase()
    .replace(/[​-‍﻿]/g, '')
    .replace(/[๐-๙]/g, (d) => String(d.charCodeAt(0) - 0x0e50))
    .replace(/\s+/g, ' ')
    .trim();
}
