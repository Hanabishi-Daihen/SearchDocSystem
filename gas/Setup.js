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
