/**
 * SearchDocSystem — ไฟล์นี้สร้างจาก gas/*.js ด้วย `npm run build` ห้ามแก้ไขโดยตรง
 */

// ==================== Config.js ====================

/**
 * ค่าคงที่ โครงสร้างตาราง และค่าตั้งค่าของระบบ
 *
 * Script Properties (ตั้งผ่านเมนู "SearchDocSystem → 2. ตั้งค่า" ใน Google Sheet — ห้ามเขียนลงโค้ด):
 *   TENANT_ID, CLIENT_ID, CLIENT_SECRET   App ใน Microsoft Entra ID (Graph API)
 *   SOURCE             โฟลเดอร์ OneDrive: ลิงก์แชร์ หรือ อีเมลเจ้าของ:/path เช่น user@daihen.co.th:/Documents/QA
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
  ['file_extensions', 'pdf,docx,doc,xlsx,xls,csv,pptx,ppt,txt', 'นามสกุลที่แตกเนื้อหามาค้นหา (ไฟล์อื่นค้นได้จากชื่ออย่างเดียว)'],
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
    fileExtensions: raw.file_extensions.split(',').map((s) => s.trim().toLowerCase().replace(/^\./, '')).filter(Boolean),
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

// ==================== Repo.js ====================

/**
 * ห่อการอ่าน/เขียน Google Sheets: โหลดทั้งตารางครั้งเดียว แก้ใน memory แล้วเขียนกลับเป็นก้อน
 * ทุก cell เก็บเป็นข้อความ (number format '@') เพื่อไม่ให้ Sheets แปลงวันที่/ตัวเลขเอง
 */

class Table {
  constructor(ss, name) {
    this.name = name;
    this.sheet = ss.getSheetByName(name);
    if (!this.sheet) throw new Error(`ไม่พบ sheet "${name}" — รันเมนู "ติดตั้ง / อัปเดตระบบ" ก่อน`);
    this.headers = SCHEMA[name];
    const lastRow = this.sheet.getLastRow();
    const values = lastRow > 1 ? this.sheet.getRange(2, 1, lastRow - 1, this.headers.length).getValues() : [];
    this.rows = values
      .filter((r) => r.some((c) => c !== ''))
      .map((r) => {
        const o = {};
        this.headers.forEach((h, i) => (o[h] = cellToString_(r[i])));
        return o;
      });
    this.dirty = false;
    this.index_ = null;
  }

  get(id) {
    if (!this.index_) this.index_ = new Map(this.rows.map((r) => [r.id, r]));
    return this.index_.get(id) || null;
  }

  filter(fn) {
    return this.rows.filter(fn);
  }

  insert(obj) {
    const row = {};
    this.headers.forEach((h) => (row[h] = obj[h] === undefined || obj[h] === null ? '' : String(obj[h])));
    this.rows.push(row);
    if (this.index_ && row.id) this.index_.set(row.id, row);
    this.dirty = true;
    return row;
  }

  update(row, patch) {
    Object.keys(patch).forEach((k) => {
      if (!this.headers.includes(k)) throw new Error(`${this.name}: ไม่มีคอลัมน์ ${k}`);
      const v = patch[k] === undefined || patch[k] === null ? '' : String(patch[k]);
      if (row[k] !== v) {
        row[k] = v;
        this.dirty = true;
      }
    });
    return row;
  }

  save() {
    if (!this.dirty) return;
    const n = this.headers.length;
    const data = this.rows.map((r) => this.headers.map((h) => r[h]));
    const needed = data.length + 1;
    if (needed > this.sheet.getMaxRows()) this.sheet.insertRowsAfter(this.sheet.getMaxRows(), needed - this.sheet.getMaxRows());
    if (data.length) {
      const range = this.sheet.getRange(2, 1, data.length, n);
      range.setNumberFormat('@');
      range.setValues(data);
    }
    const lastRow = this.sheet.getLastRow();
    if (lastRow > needed) this.sheet.getRange(needed + 1, 1, lastRow - needed, n).clearContent();
    this.dirty = false;
  }
}

function cellToString_(v) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

// ---------- Logs ----------

const logBuffer_ = [];

function log_(level, message) {
  console.log(`[${level}] ${message}`);
  logBuffer_.push([nowIso_(), level, String(message).slice(0, 2000)]);
}

function flushLogs_() {
  if (!logBuffer_.length) return;
  const sheet = getSpreadsheet_().getSheetByName('Logs');
  if (!sheet) return;
  appendRows_(sheet, logBuffer_.splice(0), 2000);
}

/** เพิ่มแถวท้าย sheet แล้วตัดแถวเก่าให้เหลือไม่เกิน keep แถว */
function appendRows_(sheet, rows, keep) {
  const range = sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length);
  range.setNumberFormat('@');
  range.setValues(rows);
  const excess = sheet.getLastRow() - 1 - keep;
  if (excess > 0) sheet.deleteRows(2, excess);
}

// ==================== Graph.js ====================

/**
 * เชื่อมต่อ OneDrive องค์กรผ่าน Microsoft Graph API ด้วย App ใน Microsoft Entra ID (client credentials)
 * ใช้ Script Properties: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SOURCE
 */

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const ITEM_SELECT = 'id,name,cTag,eTag,file,folder,fileSystemInfo,lastModifiedBy,webUrl,size,parentReference';

function graphToken_(forceRefresh) {
  const cache = CacheService.getScriptCache();
  if (!forceRefresh) {
    const cached = cache.get('graph_token');
    if (cached) return cached;
  }
  const p = props_();
  const tenant = p.getProperty('TENANT_ID');
  const clientId = p.getProperty('CLIENT_ID');
  const secret = p.getProperty('CLIENT_SECRET');
  if (!tenant || !clientId || !secret) {
    throw new Error('ยังไม่ได้ตั้งค่า Tenant ID / Client ID / Client secret (เมนู 2. ตั้งค่า)');
  }
  const res = UrlFetchApp.fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenant)}/oauth2/v2.0/token`, {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: secret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    },
    muteHttpExceptions: true,
  });
  let body = {};
  try { body = JSON.parse(res.getContentText()); } catch (e) { /* ข้อความไม่ใช่ JSON */ }
  if (res.getResponseCode() !== 200) {
    throw new Error(`ขอ token ไม่สำเร็จ: ${body.error || res.getResponseCode()} — ${String(body.error_description || '').split('\r')[0]}`);
  }
  cache.put('graph_token', body.access_token, Math.max(60, Math.min(body.expires_in - 300, 21600)));
  return body.access_token;
}

/** เรียก Graph พร้อม retry เมื่อโดน throttle (429) หรือ server error */
function graphFetch_(pathOrUrl, options) {
  const url = /^https?:\/\//.test(pathOrUrl) ? pathOrUrl : GRAPH_BASE + pathOrUrl;
  let refreshed = false;
  for (let attempt = 0; attempt < 5; attempt++) {
    const res = UrlFetchApp.fetch(url, Object.assign({
      method: 'get',
      headers: { Authorization: 'Bearer ' + graphToken_(refreshed) },
      muteHttpExceptions: true,
    }, options || {}));
    const code = res.getResponseCode();
    if (code === 401 && !refreshed) {
      refreshed = true;
      continue;
    }
    if (code === 429 || code >= 500) {
      const retryAfter = Number(res.getHeaders()['Retry-After']) || Math.pow(2, attempt);
      Utilities.sleep(Math.min(retryAfter, 30) * 1000);
      continue;
    }
    if (code >= 400) {
      let msg = res.getContentText();
      try { msg = JSON.parse(msg).error.message; } catch (e) { /* ใช้ข้อความดิบ */ }
      throw new Error(`Graph ${code}: ${String(msg).slice(0, 300)}`);
    }
    return res;
  }
  throw new Error(`Graph: ลองใหม่หลายครั้งแล้วไม่สำเร็จ (${url.replace(/\?.*$/, '')})`);
}

function graphJson_(pathOrUrl) {
  return JSON.parse(graphFetch_(pathOrUrl).getContentText());
}

/**
 * แปลงค่า SOURCE เป็นโฟลเดอร์ใน OneDrive รองรับ 2 แบบ
 *   1) ลิงก์แชร์ OneDrive/SharePoint
 *   2) "อีเมลเจ้าของ:/path/โฟลเดอร์" เช่น user@daihen.co.th:/Documents/QA (ใช้เมื่อเปิดผ่านลิงก์ไม่ได้)
 */
function resolveSource_(input) {
  input = String(input || '').trim();
  if (!input) throw new Error('ยังไม่ได้ตั้งค่าโฟลเดอร์ OneDrive (เมนู 2. ตั้งค่า)');
  let item;
  if (/^https?:\/\//i.test(input)) {
    const encoded = 'u!' + Utilities.base64EncodeWebSafe(input, Utilities.Charset.UTF_8).replace(/=+$/, '');
    item = graphJson_(`/shares/${encoded}/driveItem?$select=${ITEM_SELECT}`);
  } else {
    const m = input.match(/^([^\s:/]+@[^\s:/]+):(.*)$/);
    if (!m) throw new Error('รูปแบบโฟลเดอร์ไม่ถูกต้อง: ใส่ลิงก์ OneDrive หรือ อีเมล:/path/โฟลเดอร์');
    const path = m[2].replace(/^\/+|\/+$/g, '');
    const upn = encodeURIComponent(m[1]);
    item = graphJson_(path
      ? `/users/${upn}/drive/root:/${path.split('/').map(encodeURIComponent).join('/')}?$select=${ITEM_SELECT}`
      : `/users/${upn}/drive/root?$select=${ITEM_SELECT}`);
  }
  if (!item.folder) throw new Error(`"${item.name}" ไม่ใช่โฟลเดอร์`);
  return { driveId: item.parentReference && item.parentReference.driveId, itemId: item.id, name: item.name };
}

/** โฟลเดอร์ต้นทางที่ resolve แล้ว (เก็บไว้ใน Script Properties ไม่ต้อง resolve ทุกรอบ) */
function getSource_() {
  const p = props_();
  const input = (p.getProperty('SOURCE') || '').trim();
  if (p.getProperty('SOURCE_RESOLVED_FROM') === input && p.getProperty('SOURCE_ITEM_ID')) {
    return { driveId: p.getProperty('SOURCE_DRIVE_ID'), itemId: p.getProperty('SOURCE_ITEM_ID'), name: p.getProperty('SOURCE_NAME') };
  }
  const s = resolveSource_(input);
  p.setProperty('SOURCE_DRIVE_ID', s.driveId);
  p.setProperty('SOURCE_ITEM_ID', s.itemId);
  p.setProperty('SOURCE_NAME', s.name);
  p.setProperty('SOURCE_RESOLVED_FROM', input);
  return s;
}

/** ไฟล์ทั้งหมดใต้โฟลเดอร์ (ไล่โฟลเดอร์ย่อย) → [{ item, path }] */
function listAllFiles_(source) {
  const out = [];
  const queue = [{ id: source.itemId, path: '' }];
  let folders = 0;
  while (queue.length) {
    const { id, path } = queue.shift();
    if (++folders > 5000) throw new Error('โฟลเดอร์ย่อยมากเกิน 5,000 โฟลเดอร์');
    let url = `/drives/${source.driveId}/items/${id}/children?$select=${ITEM_SELECT}&$top=200`;
    while (url) {
      const page = graphJson_(url);
      page.value.forEach((child) => {
        if (child.folder) queue.push({ id: child.id, path: path ? `${path}/${child.name}` : child.name });
        else if (child.file) out.push({ item: child, path });
      });
      url = page['@odata.nextLink'] || null;
    }
  }
  return out;
}

/** ดาวน์โหลดไฟล์: Graph ตอบ 302 ไปยัง URL ที่ยืนยันตัวตนแล้ว จึงตามไปโดยไม่ส่ง token ต่อ */
function downloadItem_(driveId, itemId) {
  const res = graphFetch_(`/drives/${driveId}/items/${itemId}/content`, { followRedirects: false });
  if (res.getResponseCode() === 302) {
    const h = res.getHeaders();
    return UrlFetchApp.fetch(h.Location || h.location).getBlob();
  }
  return res.getBlob();
}

// ==================== Extract.js ====================

/**
 * แตกข้อความจากไฟล์ด้วยการแปลงเป็นไฟล์ Google ชั่วคราว (ต้องเปิด Advanced Service: Drive API v3)
 *   PDF          → Google Docs + OCR (ได้ทั้งไฟล์พิมพ์และไฟล์สแกน)
 *   Word         → Google Docs
 *   Excel / CSV  → Google Sheets
 *   PowerPoint   → Google Slides
 *   txt / md     → อ่านตรง
 */

const CONVERT_TARGET = {
  pdf: 'application/vnd.google-apps.document',
  doc: 'application/vnd.google-apps.document',
  docx: 'application/vnd.google-apps.document',
  xls: 'application/vnd.google-apps.spreadsheet',
  xlsx: 'application/vnd.google-apps.spreadsheet',
  csv: 'application/vnd.google-apps.spreadsheet',
  ppt: 'application/vnd.google-apps.presentation',
  pptx: 'application/vnd.google-apps.presentation',
};

function extractText_(blob, ext) {
  if (ext === 'txt' || ext === 'md') return blob.getDataAsString('UTF-8');
  const target = CONVERT_TARGET[ext];
  if (!target) return '';
  const options = ext === 'pdf' ? { ocrLanguage: getSettings_().ocrLanguage } : {};
  const tmp = Drive.Files.create(
    { name: `_tmp_extract.${ext}`, mimeType: target, parents: [props_().getProperty('TMP_FOLDER_ID')] },
    blob,
    options
  );
  try {
    if (target === CONVERT_TARGET.docx) return DocumentApp.openById(tmp.id).getBody().getText();
    if (target === CONVERT_TARGET.xlsx) {
      return SpreadsheetApp.openById(tmp.id).getSheets()
        .map((s) => s.getDataRange().getDisplayValues().map((r) => r.filter(String).join(' ')).filter(Boolean).join('\n'))
        .join('\n');
    }
    return SlidesApp.openById(tmp.id).getSlides()
      .map((slide) => slide.getShapes().map((sh) => {
        try { return sh.getText().asString(); } catch (e) { return ''; }
      }).join('\n'))
      .join('\n');
  } finally {
    try { Drive.Files.remove(tmp.id); } catch (e) { log_('WARN', `ลบไฟล์ชั่วคราว ${tmp.id} ไม่ได้: ${e.message}`); }
  }
}

/** บันทึกข้อความลง Drive /text — ใช้ไฟล์เดิมถ้ามี */
function saveText_(fileRow, text) {
  if (fileRow.text_file_id) {
    try {
      const f = DriveApp.getFileById(fileRow.text_file_id);
      if (!f.isTrashed()) {
        f.setContent(text);
        return f.getId();
      }
    } catch (e) { /* ไฟล์ถูกลบไปแล้ว สร้างใหม่ */ }
  }
  const folder = DriveApp.getFolderById(props_().getProperty('TEXT_FOLDER_ID'));
  return folder.createFile(`${fileRow.id}.txt`, text, MimeType.PLAIN_TEXT).getId();
}

function readText_(textFileId) {
  if (!textFileId) return '';
  try {
    return DriveApp.getFileById(textFileId).getBlob().getDataAsString('UTF-8');
  } catch (e) {
    return '';
  }
}

// ==================== Sync.js ====================

/**
 * Sync: เทียบรายการไฟล์ใน OneDrive กับ sheet Files → แตกข้อความไฟล์ใหม่/ที่เปลี่ยน → สร้าง search index
 *   runScheduled(e)  เรียกโดย time trigger
 *   syncNow()        เรียกจากเมนูใน Google Sheet
 * งานที่ทำไม่ทันในรอบเดียว (GAS รันได้ครั้งละ 6 นาที) จะทำต่อในรอบถัดไป
 */

function runScheduled(e) {
  const isTrigger = e && e.triggerUid &&
    ScriptApp.getProjectTriggers().some((t) => t.getUniqueId() === e.triggerUid);
  if (!isTrigger) requireOwner_();
  return runSync_();
}

function syncNow() {
  requireOwner_();
  return runSync_();
}

function runSync_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    log_('INFO', 'ข้ามรอบนี้: มีงาน sync อื่นกำลังทำงานอยู่');
    flushLogs_();
    return { skipped: true };
  }
  const started = Date.now();
  const summary = { total: 0, added: 0, changed: 0, moved: 0, deleted: 0, extracted: 0, failed: 0, remaining: 0 };
  try {
    const files = new Table(getSpreadsheet_(), 'Files');
    const diff = applyListing_(files, listAllFiles_(getSource_()));
    Object.assign(summary, diff);
    files.save();

    const processed = processQueue_(files, started);
    summary.extracted = processed.extracted;
    summary.failed = processed.failed;
    files.save();

    summary.remaining = files.filter((f) => f.status === STATUS.QUEUED).length;
    summary.total = files.filter((f) => f.status !== STATUS.DELETED).length;
    if (diff.added || diff.changed || diff.moved || diff.deleted || processed.extracted || processed.failed ||
        !props_().getProperty('INDEX_MANIFEST')) {
      rebuildSearchIndex_(files);
    }
    props_().setProperty('LAST_SYNC', JSON.stringify({ at: nowIso_(), ok: true, total: summary.total, remaining: summary.remaining }));
    log_('INFO', `sync เสร็จใน ${Math.round((Date.now() - started) / 1000)} วินาที: ${JSON.stringify(summary)}`);
    return summary;
  } catch (err) {
    props_().setProperty('LAST_SYNC', JSON.stringify({ at: nowIso_(), ok: false, error: err.message }));
    log_('ERROR', `sync ล้มเหลว: ${err.message}`);
    throw err;
  } finally {
    flushLogs_();
    lock.releaseLock();
  }
}

/** เทียบรายการจาก OneDrive กับตาราง Files (แก้ใน memory) */
function applyListing_(files, listing) {
  const settings = getSettings_();
  const result = { added: 0, changed: 0, moved: 0, deleted: 0 };
  const seen = new Set();

  listing.forEach(({ item, path }) => {
    seen.add(item.id);
    const ext = (item.name.match(/\.([A-Za-z0-9]{1,6})$/) || ['', ''])[1].toLowerCase();
    const changeKey = item.cTag || item.eTag || `${item.fileSystemInfo && item.fileSystemInfo.lastModifiedDateTime}|${item.size}`;
    const extractable = settings.fileExtensions.includes(ext) && Number(item.size) <= settings.maxFileBytes;
    const meta = {
      name: item.name,
      ext,
      path,
      web_url: item.webUrl || '',
      size: item.size || 0,
      modified_at: (item.fileSystemInfo && item.fileSystemInfo.lastModifiedDateTime) || '',
      modified_by: (item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName) || '',
      change_key: changeKey,
    };

    const row = files.get(item.id);
    if (!row) {
      files.insert(Object.assign({ id: item.id, first_seen_at: nowIso_(), attempts: 0 }, meta,
        { status: extractable ? STATUS.QUEUED : STATUS.NAME_ONLY }));
      result.added++;
      return;
    }
    if (row.name !== meta.name || row.path !== meta.path || row.web_url !== meta.web_url) result.moved++;
    const contentChanged = row.indexed_key !== changeKey;
    const wasDeleted = row.status === STATUS.DELETED;
    files.update(row, meta);
    if (!extractable) {
      if (row.status !== STATUS.NAME_ONLY) files.update(row, { status: STATUS.NAME_ONLY });
    } else if (contentChanged && row.status !== STATUS.QUEUED) {
      // เนื้อหาเปลี่ยน / ไฟล์กลับมา / เพิ่งรองรับชนิดไฟล์นี้ → แตกข้อความใหม่
      // (ไฟล์ที่ล้มเหลวครบแล้วมี indexed_key = key ตอนล้มเหลว จึงไม่วนลองใหม่จนกว่าไฟล์จะถูกแก้)
      files.update(row, { status: STATUS.QUEUED, attempts: 0, last_error: '' });
      result.changed++;
    } else if (!contentChanged && wasDeleted) {
      files.update(row, { status: row.text_file_id ? STATUS.INDEXED : STATUS.QUEUED });
    }
  });

  files.rows.forEach((row) => {
    if (!seen.has(row.id) && row.status !== STATUS.DELETED) {
      files.update(row, { status: STATUS.DELETED });
      result.deleted++;
    }
  });
  return result;
}

/** แตกข้อความไฟล์ในคิวจนกว่าจะใกล้หมดเวลา */
function processQueue_(files, started) {
  const source = getSource_();
  const out = { extracted: 0, failed: 0 };
  const queue = files.filter((f) => f.status === STATUS.QUEUED)
    .sort((a, b) => b.modified_at.localeCompare(a.modified_at)); // ไฟล์ใหม่ก่อน
  for (const row of queue) {
    if (Date.now() - started > MAX_RUNTIME_MS) break;
    const key = row.change_key;
    try {
      const blob = downloadItem_(source.driveId, row.id);
      const text = extractText_(blob, row.ext).replace(/\r\n?/g, '\n').trim();
      files.update(row, {
        text_file_id: saveText_(row, text),
        text_chars: text.length,
        status: STATUS.INDEXED,
        indexed_key: key,
        indexed_at: nowIso_(),
        attempts: 0,
        last_error: '',
      });
      out.extracted++;
    } catch (err) {
      const attempts = Number(row.attempts || 0) + 1;
      files.update(row, {
        attempts,
        last_error: String(err.message).slice(0, 500),
        status: attempts >= MAX_ATTEMPTS ? STATUS.ERROR : STATUS.QUEUED,
        // ถ้าล้มเหลวครบแล้ว จำ key ไว้ จะไม่ลองใหม่จนกว่าไฟล์จะถูกแก้
        indexed_key: attempts >= MAX_ATTEMPTS ? key : row.indexed_key,
      });
      log_('WARN', `แตกข้อความ "${row.name}" ไม่สำเร็จ (ครั้งที่ ${attempts}): ${err.message}`);
      out.failed++;
    }
  }
  return out;
}

/** เมนู: ให้ไฟล์ที่ล้มเหลวลองใหม่ */
function retryFailed() {
  requireOwner_();
  const files = new Table(getSpreadsheet_(), 'Files');
  let n = 0;
  files.filter((f) => f.status === STATUS.ERROR).forEach((f) => {
    files.update(f, { status: STATUS.QUEUED, attempts: 0, indexed_key: '' });
    n++;
  });
  files.save();
  return n;
}

// ==================== SearchIndex.js ====================

/**
 * Search index: รายการไฟล์ + ข้อความ (normalize แล้ว) แบ่งเป็นชุด บีบอัด gzip เก็บใน Drive /index
 * หน้าเว็บโหลดไปครั้งเดียว (cache ใน IndexedDB ตาม version) แล้วค้นหาใน browser
 *
 * ข้อมูลแต่ละไฟล์: { i: id, n: ชื่อ, e: นามสกุล, p: โฟลเดอร์, s: ขนาด, m: วันที่แก้ไข, b: ผู้แก้ไข, u: ลิงก์ OneDrive,
 *                   k: key ของข้อความ, t: ข้อความ }
 */

const INDEX_CHUNK_CHARS = 3000000; // ขนาด JSON ต่อชุดก่อนบีบอัด (~1 MB หลัง gzip)

function rebuildSearchIndex_(files) {
  const p = props_();
  const maxChars = getSettings_().maxTextChars;
  const oldManifest = getIndexManifest_();
  const oldTexts = loadIndexTexts_(oldManifest);
  let reads = 0;

  const docs = files.filter((f) => f.status !== STATUS.DELETED)
    .sort((a, b) => b.modified_at.localeCompare(a.modified_at))
    .map((f) => {
      let text = '';
      if (f.status === STATUS.INDEXED && f.text_file_id) {
        const old = oldTexts[f.id];
        if (old && old.k === f.indexed_key) text = old.t;
        else {
          text = normalizeText_(readText_(f.text_file_id)).slice(0, maxChars);
          reads++;
        }
      }
      return {
        i: f.id, n: f.name, e: f.ext, p: f.path, s: Number(f.size) || 0, m: f.modified_at, b: f.modified_by,
        u: f.web_url, k: f.status === STATUS.INDEXED ? f.indexed_key : '', t: text,
      };
    });

  const chunks = [];
  let current = [];
  let size = 0;
  docs.forEach((d) => {
    const len = d.t.length + d.n.length + d.p.length + 200;
    if (current.length && size + len > INDEX_CHUNK_CHARS) {
      chunks.push(current);
      current = [];
      size = 0;
    }
    current.push(d);
    size += len;
  });
  if (current.length || !chunks.length) chunks.push(current);

  const folder = DriveApp.getFolderById(p.getProperty('INDEX_FOLDER_ID'));
  const version = String(Date.now());
  const ids = chunks.map((c, n) => {
    const gz = Utilities.gzip(Utilities.newBlob(JSON.stringify(c), 'application/json'), `index-${version}-${n}.json.gz`);
    return folder.createFile(gz).getId();
  });
  const manifest = { version, builtAt: nowIso_(), count: docs.length, indexed: docs.filter((d) => d.k).length, chunks: ids };
  p.setProperty('INDEX_MANIFEST', JSON.stringify(manifest));
  (oldManifest.chunks || []).forEach((id) => {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { /* ลบไปแล้ว */ }
  });
  log_('INFO', `สร้าง search index: ${docs.length} ไฟล์, ${ids.length} ชุด (อ่านข้อความใหม่ ${reads} ไฟล์)`);
  return manifest;
}

function getIndexManifest_() {
  try {
    return JSON.parse(props_().getProperty('INDEX_MANIFEST') || '{}');
  } catch (e) {
    return {};
  }
}

function readIndexChunk_(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob().setContentType('application/x-gzip');
  return JSON.parse(Utilities.ungzip(blob).getDataAsString('UTF-8'));
}

/** ข้อความเดิมจาก index ชุดก่อน → { id: { k, t } } ไม่ต้องอ่านไฟล์ที่ไม่เปลี่ยนจาก Drive ซ้ำ */
function loadIndexTexts_(manifest) {
  const out = {};
  (manifest.chunks || []).forEach((id) => {
    try {
      readIndexChunk_(id).forEach((d) => { if (d.k) out[d.i] = { k: d.k, t: d.t }; });
    } catch (err) {
      log_('WARN', `อ่าน index เดิมไม่ได้ จะอ่านข้อความใหม่: ${err.message}`);
    }
  });
  return out;
}

// ==================== Api.js ====================

/**
 * ฟังก์ชันที่หน้าเว็บเรียกผ่าน google.script.run
 * ถ้าตั้ง SITE_PASSWORD ไว้ ทุกฟังก์ชัน (ยกเว้น getAppInfo / login) ต้องส่ง token ที่ได้จาก login มาด้วย
 */

const SESSION_TTL = 6 * 60 * 60; // วินาที

function getAppInfo() {
  const manifest = getIndexManifest_();
  let lastSync = {};
  try { lastSync = JSON.parse(props_().getProperty('LAST_SYNC') || '{}'); } catch (e) { /* ยังไม่เคย sync */ }
  return {
    name: APP_NAME,
    needCode: !!props_().getProperty('SITE_PASSWORD'),
    indexVersion: manifest.version || '',
    count: manifest.count || 0,
    lastSyncAt: lastSync.at || '',
    lastSyncOk: lastSync.ok !== false,
    remaining: lastSync.remaining || 0,
  };
}

function login(code) {
  const expected = props_().getProperty('SITE_PASSWORD');
  if (!expected) return '';
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('login_fails') || 0);
  if (fails >= 20) throw new Error('ใส่รหัสผิดหลายครั้งเกินไป ลองใหม่ในอีก 10 นาที');
  if (String(code || '') !== expected) {
    cache.put('login_fails', String(fails + 1), 600);
    throw new Error('รหัสไม่ถูกต้อง');
  }
  const token = Utilities.getUuid();
  cache.put('session_' + token, '1', SESSION_TTL);
  return token;
}

function checkAccess_(token) {
  if (!props_().getProperty('SITE_PASSWORD')) return;
  if (!token || !CacheService.getScriptCache().get('session_' + token)) {
    throw new Error('SESSION_EXPIRED');
  }
}

function getIndexMeta(token) {
  checkAccess_(token);
  const m = getIndexManifest_();
  return { version: m.version || '', builtAt: m.builtAt || '', count: m.count || 0, chunks: (m.chunks || []).length };
}

/** ชุดที่ n ของ index เป็น base64 ของ gzip (หน้าเว็บแตกเองด้วย DecompressionStream) หรือ JSON ถ้า plain = true */
function getIndexChunk(token, version, n, plain) {
  checkAccess_(token);
  const m = getIndexManifest_();
  if (m.version !== version) throw new Error('INDEX_CHANGED');
  const id = (m.chunks || [])[n];
  if (!id) throw new Error('ไม่พบ index ชุดที่ ' + n);
  const blob = DriveApp.getFileById(id).getBlob();
  if (plain) return Utilities.ungzip(blob.setContentType('application/x-gzip')).getDataAsString('UTF-8');
  return Utilities.base64Encode(blob.getBytes());
}

/** ข้อความเต็มของไฟล์ (สำหรับแผงพรีวิว) */
function getText(token, fileId) {
  checkAccess_(token);
  const files = new Table(getSpreadsheet_(), 'Files');
  const f = files.get(String(fileId));
  if (!f || f.status === STATUS.DELETED) throw new Error('ไม่พบไฟล์');
  return readText_(f.text_file_id).slice(0, 200000);
}

/** สถิติคำค้น — ไม่เก็บตัวตนผู้ค้น */
function logSearch(token, query, results) {
  checkAccess_(token);
  const q = String(query || '').trim().slice(0, 200);
  if (!q) return;
  const sheet = getSpreadsheet_().getSheetByName('SearchLog');
  if (sheet) appendRows_(sheet, [[nowIso_(), q, String(Number(results) || 0)]], 5000);
}

// ==================== WebApp.js ====================

/**
 * หน้าเว็บ (Deploy → Web app) — ชื่อไฟล์ HTML ใน Apps Script ต้องตรงกับชื่อที่ส่งให้ createHtmlOutputFromFile
 * (ตัวพิมพ์เล็ก/ใหญ่มีผล: ไฟล์ "index.html" → 'index')
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== Menu.js ====================

/**
 * เมนู "🔎 SearchDocSystem" ใน Google Sheet — ติดตั้ง ตั้งค่า sync และทดสอบ โดยไม่ต้องเข้า Apps Script editor
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔎 SearchDocSystem')
    .addItem('1. ติดตั้ง / อัปเดตระบบ', 'menuInstall')
    .addItem('2. ตั้งค่า (OneDrive, รหัสเข้าเว็บ)', 'menuSettings')
    .addItem('3. เปิดหน้าเว็บ', 'menuOpenWebApp')
    .addSeparator()
    .addItem('Sync ตอนนี้', 'menuSync')
    .addItem('ลองแตกข้อความไฟล์ที่ล้มเหลวใหม่', 'menuRetryFailed')
    .addSubMenu(ui.createMenu('ทดสอบ')
      .addItem('เชื่อมต่อ OneDrive', 'menuTestConnection')
      .addItem('แตกข้อความ (OCR) ไฟล์แรก', 'menuTestExtract'))
    .addToUi();
}

function runMenu_(title, fn) {
  const ui = SpreadsheetApp.getUi();
  try {
    const message = fn();
    if (message) ui.alert(title, message, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert(`${title}: ไม่สำเร็จ`, err.message, ui.ButtonSet.OK);
  }
}

function menuInstall() { runMenu_('ติดตั้งระบบ', () => setup()); }
function menuTestConnection() { runMenu_('ทดสอบเชื่อมต่อ OneDrive', () => testConnection()); }
function menuTestExtract() { runMenu_('ทดสอบแตกข้อความ', () => testExtract()); }

function menuSync() {
  runMenu_('Sync', () => {
    const r = syncNow();
    if (r.skipped) return 'มีงาน sync อื่นกำลังทำงานอยู่ ลองใหม่ในอีกสักครู่';
    return `ไฟล์ทั้งหมด ${r.total}\nใหม่ ${r.added} · เปลี่ยนแปลง ${r.changed} · ย้าย/เปลี่ยนชื่อ ${r.moved} · ถูกลบ ${r.deleted}\n` +
      `แตกข้อความแล้ว ${r.extracted} · ล้มเหลว ${r.failed}` +
      (r.remaining ? `\nเหลืออีก ${r.remaining} ไฟล์ — ระบบจะทำต่อในรอบถัดไป` : '');
  });
}

function menuRetryFailed() {
  runMenu_('ลองใหม่', () => `ใส่ ${retryFailed()} ไฟล์กลับเข้าคิวแล้ว — จะแตกข้อความในรอบ sync ถัดไป`);
}

function menuOpenWebApp() {
  runMenu_('หน้าเว็บ', () => {
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      return 'ยังไม่ได้ Deploy หน้าเว็บ\n\nที่ Apps Script: การทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป\n' +
        'ดำเนินการในฐานะ: ฉัน → ทำให้ใช้งานได้';
    }
    const html = HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif;font-size:14px;line-height:1.6">
        <p>ลิงก์หน้าเว็บ (ส่งให้ผู้ใช้ได้เลย):</p>
        <p><a href="${url}" target="_blank" style="color:#cc785c;word-break:break-all">${url}</a></p>
      </div>`).setWidth(460).setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(html, 'เปิดหน้าเว็บ');
    return '';
  });
}

// ---------- หน้าต่างตั้งค่า ----------

const CONFIG_FIELDS = [
  { key: 'TENANT_ID', label: 'Directory (tenant) ID — จาก IT', secret: false },
  { key: 'CLIENT_ID', label: 'Application (client) ID — จาก IT', secret: false },
  { key: 'CLIENT_SECRET', label: 'Client secret (Value) — จาก IT', secret: true },
  { key: 'SOURCE', label: 'โฟลเดอร์ OneDrive: ลิงก์แชร์ หรือ อีเมล:/path', secret: false },
  { key: 'SITE_PASSWORD', label: 'รหัสเข้าหน้าเว็บค้นหา (ไม่บังคับ)', secret: true },
];

function menuSettings() {
  runMenu_('ตั้งค่า', () => {
    requireOwner_();
    const p = props_();
    const esc = (s) => String(s || '').replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
    const fields = CONFIG_FIELDS.map((f) => {
      const current = p.getProperty(f.key);
      return `<label>${esc(f.label)}
        <input name="${f.key}" type="${f.secret ? 'password' : 'text'}" autocomplete="off"
          value="${f.secret ? '' : esc(current)}" placeholder="${f.secret && current ? 'ตั้งไว้แล้ว — เว้นว่างเพื่อใช้ค่าเดิม' : ''}">
      </label>`;
    }).join('');
    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: sans-serif; font-size: 14px; color: #141413; margin: 0; padding: 4px 2px; }
        label { display: block; margin-bottom: 12px; color: #6c6a64; }
        input { display: block; width: 100%; box-sizing: border-box; height: 36px; margin-top: 4px; padding: 0 10px;
          border: 1px solid #e6dfd8; border-radius: 8px; font-size: 14px; color: #141413; }
        p { color: #6c6a64; font-size: 13px; margin: 0 0 14px; }
        button { height: 36px; padding: 0 18px; border: 0; border-radius: 8px; background: #cc785c; color: #fff; font-size: 14px; cursor: pointer; }
        #msg { margin-left: 10px; font-size: 13px; }
      </style>
      <p>ใส่ค่าจาก IT ทีหลังได้ · ช่องรหัสที่เว้นว่างจะใช้ค่าเดิม · พิมพ์ - เพื่อลบรหัสเข้าเว็บ</p>
      <form id="f">${fields}<button type="submit">บันทึก</button><span id="msg"></span></form>
      <script>
        document.getElementById('f').onsubmit = (e) => {
          e.preventDefault();
          const data = Object.fromEntries(new FormData(e.target));
          document.getElementById('msg').textContent = 'กำลังบันทึก…';
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .withFailureHandler((err) => { document.getElementById('msg').textContent = err.message; })
            .saveSettingsFromDialog(data);
        };
      </script>`).setWidth(440).setHeight(470);
    SpreadsheetApp.getUi().showModalDialog(html, 'ตั้งค่า SearchDocSystem');
    return '';
  });
}

/** เรียกจากหน้าต่างตั้งค่า — ช่องที่เว้นว่างจะไม่เปลี่ยนค่าเดิม, "-" = ลบค่า */
function saveSettingsFromDialog(values) {
  requireOwner_();
  const p = props_();
  CONFIG_FIELDS.forEach(({ key }) => {
    const v = String((values && values[key]) || '').trim();
    if (v === '-') p.deleteProperty(key);
    else if (v) p.setProperty(key, v);
  });
  CacheService.getScriptCache().remove('graph_token');
  return true;
}

// ==================== Setup.js ====================

/**
 * ติดตั้งระบบและฟังก์ชันทดสอบ — เรียกจากเมนูใน Google Sheet (Menu.js) หรือจาก Apps Script editor
 *   setup()            สร้าง sheet, โฟลเดอร์ใน Drive และ trigger — รันซ้ำได้ ไม่ลบข้อมูลเดิม
 *   testConnection()   ทดสอบ Graph API + อ่านรายการไฟล์
 *   testExtract()      ทดสอบดาวน์โหลด + แตกข้อความไฟล์แรกที่รองรับ
 */

function setup() {
  requireOwner_();
  const ss = getSpreadsheet_();
  const p = props_();
  if (!p.getProperty('SPREADSHEET_ID')) p.setProperty('SPREADSHEET_ID', ss.getId());

  Object.entries(SCHEMA).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const existing = sheet.getLastColumn() ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String) : [];
    if (existing.filter(Boolean).length && existing.slice(0, headers.length).join() !== headers.join()) {
      throw new Error(`sheet "${name}" มีหัวคอลัมน์ไม่ตรงกับระบบ — เปลี่ยนชื่อ sheet เดิมก่อนแล้วรันใหม่`);
    }
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#efe9de');
    sheet.getRange(1, 1, sheet.getMaxRows(), headers.length).setNumberFormat('@');
    sheet.setFrozenRows(1);
  });

  const settings = ss.getSheetByName('Settings');
  const keys = settings.getLastRow() > 1 ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map((r) => String(r[0])) : [];
  const missing = DEFAULT_SETTINGS.filter(([k]) => !keys.includes(k));
  if (missing.length) settings.getRange(settings.getLastRow() + 1, 1, missing.length, 3).setValues(missing);

  // ลบ sheet ว่างเริ่มต้น (Sheet1 / ชีต1)
  ss.getSheets().forEach((s) => {
    if (!SCHEMA[s.getName()] && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });

  const root = ensureFolder_(p.getProperty('DRIVE_ROOT_ID'), null, APP_NAME);
  p.setProperty('DRIVE_ROOT_ID', root.getId());
  [['TEXT_FOLDER_ID', 'text'], ['INDEX_FOLDER_ID', 'index'], ['TMP_FOLDER_ID', 'tmp']].forEach(([key, name]) => {
    p.setProperty(key, ensureFolder_(p.getProperty(key), root, name).getId());
  });

  ScriptApp.getProjectTriggers()
    .filter((t) => t.getHandlerFunction() === 'runScheduled')
    .forEach((t) => ScriptApp.deleteTrigger(t));
  settingsCache_ = null;
  const minutes = getSettings_().syncIntervalMinutes;
  ScriptApp.newTrigger('runScheduled').timeBased().everyMinutes(minutes).create();

  log_('INFO', `setup เสร็จ: sync ทุก ${minutes} นาที`);
  flushLogs_();

  const need = ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET', 'SOURCE'].filter((k) => !p.getProperty(k));
  return [
    '✅ ติดตั้งเสร็จแล้ว',
    `• sheet: ${Object.keys(SCHEMA).join(', ')}`,
    `• โฟลเดอร์ใน Google Drive: ${root.getUrl()}`,
    `• Sync อัตโนมัติทุก ${minutes} นาที`,
    need.length ? `\nขั้นต่อไป: เมนู "2. ตั้งค่า" — ยังไม่ได้ใส่ ${need.join(', ')}` : '\nขั้นต่อไป: เมนู "ทดสอบ → เชื่อมต่อ OneDrive"',
  ].join('\n');
}

function ensureFolder_(id, parent, name) {
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (e) { /* ถูกลบไปแล้ว สร้างใหม่ */ }
  }
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

function testConnection() {
  requireOwner_();
  graphToken_(true);
  const lines = ['✅ ขอ token จาก Microsoft Entra ID สำเร็จ'];
  const source = getSource_();
  lines.push(`✅ เปิดโฟลเดอร์ได้: "${source.name}"`);
  const items = listAllFiles_(source);
  const exts = getSettings_().fileExtensions;
  const supported = items.filter(({ item }) => exts.includes((item.name.split('.').pop() || '').toLowerCase())).length;
  lines.push(`พบ ${items.length} ไฟล์ (แตกเนื้อหาได้ ${supported} ไฟล์)${items.length > 15 ? ' — แสดง 15 ไฟล์แรก' : ''}:`, '');
  items.slice(0, 15).forEach(({ item, path }) => lines.push(`• ${path ? path + '/' : ''}${item.name}`));
  return lines.join('\n');
}

function testExtract() {
  requireOwner_();
  const source = getSource_();
  const exts = getSettings_().fileExtensions;
  const found = listAllFiles_(source).find(({ item }) => exts.includes((item.name.split('.').pop() || '').toLowerCase()));
  if (!found) throw new Error('ไม่พบไฟล์ชนิดที่รองรับในโฟลเดอร์');
  const { item } = found;
  const started = Date.now();
  const text = extractText_(downloadItem_(source.driveId, item.id), item.name.split('.').pop().toLowerCase());
  const chars = Array.from(text.replace(/\s/g, ''));
  const thai = chars.filter((c) => /[฀-๿]/.test(c)).length;
  return [
    `ไฟล์: ${item.name} (${(item.size / 1024 / 1024).toFixed(2)} MB)`,
    `ใช้เวลา ${((Date.now() - started) / 1000).toFixed(1)} วินาที · ได้ ${text.length.toLocaleString()} ตัวอักษร ` +
      `(ภาษาไทย ${chars.length ? Math.round(thai / chars.length * 100) : 0}%)`,
    '',
    '— ข้อความช่วงต้น —',
    text.slice(0, 600) || '(ไม่มีข้อความ)',
  ].join('\n');
}
