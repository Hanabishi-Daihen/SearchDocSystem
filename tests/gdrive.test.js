const test = require('node:test');
const assert = require('node:assert');
const zlib = require('node:zlib');
const { createGas } = require('./gas-harness');

const FOLDER_URL = 'https://drive.google.com/drive/folders/GROOT_FOLDER_ID_0000001?usp=sharing';

function install() {
  const g = createGas();
  for (const k of ['TENANT_ID', 'CLIENT_ID', 'CLIENT_SECRET']) g.props.delete(k); // ไม่ต้องใช้ App จาก IT
  g.props.set('SOURCE', FOLDER_URL);
  g.run('setup()');
  const d = g.gdrive;
  d.add({ id: 'SUBFOLDER_ID', name: 'Procedure', mimeType: 'application/vnd.google-apps.folder' });
  d.add({ id: 'pdf1', name: 'ขั้นตอนการตรวจรับวัตถุดิบ Rev.2.pdf', parent: 'SUBFOLDER_ID', content: 'ตรวจรับวัตถุดิบ ๕ ล็อต' });
  d.add({ id: 'docx1', name: 'ใบขอซื้อ.docx', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', content: 'Purchase Request' });
  d.add({ id: 'gdoc1', name: 'คู่มือพนักงานใหม่', mimeType: 'application/vnd.google-apps.document', content: 'ยินดีต้อนรับ เวลาทำงาน 08:00' });
  d.add({ id: 'gsheet1', name: 'QC Checklist', mimeType: 'application/vnd.google-apps.spreadsheet', content: 'Torque,45' });
  d.add({ id: 'gslides1', name: 'Machine Setup', mimeType: 'application/vnd.google-apps.presentation', content: 'Robot welding' });
  d.add({ id: 'form1', name: 'แบบสอบถาม', mimeType: 'application/vnd.google-apps.form' });
  d.add({ id: 'short1', name: 'ทางลัด', mimeType: 'application/vnd.google-apps.shortcut' });
  d.add({ id: 'trash1', name: 'ลบแล้ว.pdf', content: 'x', trashed: true });
  // โฟลเดอร์ของระบบเองอยู่ในโฟลเดอร์ต้นทาง → ต้องไม่ถูกนำมา index
  d.add({ id: g.props.get('DRIVE_ROOT_ID'), name: 'SearchDocSystem', mimeType: 'application/vnd.google-apps.folder' });
  d.add({ id: 'own1', name: 'pdf1.txt', mimeType: 'text/plain', parent: g.props.get('DRIVE_ROOT_ID'), content: 'x' });
  return g;
}

function indexDocs(g) {
  const meta = g.run('getIndexMeta("")');
  const docs = [];
  for (let n = 0; n < meta.chunks; n++) {
    const b64 = g.run(`getIndexChunk("", ${JSON.stringify(meta.version)}, ${n})`);
    docs.push(...JSON.parse(zlib.gunzipSync(Buffer.from(b64, 'base64')).toString('utf8')));
  }
  return Object.fromEntries(docs.map((d) => [d.i, d]));
}

test('รู้จักลิงก์โฟลเดอร์ Google Drive หลายรูปแบบ', () => {
  const g = createGas();
  const cases = {
    'https://drive.google.com/drive/folders/1AbC_dEf-123456789?usp=sharing': '1AbC_dEf-123456789',
    'https://drive.google.com/drive/u/0/folders/1AbC_dEf-123456789': '1AbC_dEf-123456789',
    'https://drive.google.com/open?id=1AbC_dEf-123456789': '1AbC_dEf-123456789',
    '1AbC_dEf-123456789xyz12': '1AbC_dEf-123456789xyz12',
    'https://daihen-my.sharepoint.com/:f:/g/personal/x/abc': '',
    'user@daihen.co.th:/Documents/QA': '',
  };
  for (const [input, id] of Object.entries(cases)) {
    assert.strictEqual(g.run(`parseDriveFolderId_(${JSON.stringify(input)})`), id, input);
  }
});

test('ติดตั้งโดยไม่มีค่า App จาก IT และทดสอบการเชื่อมต่อได้', () => {
  const g = install();
  const msg = g.run('setup()');
  assert.doesNotMatch(msg, /TENANT_ID|CLIENT/);
  const t = g.run('testConnection()');
  assert.match(t, /เปิดโฟลเดอร์ Google Drive ได้: "เอกสาร QA"/);
  assert.match(t, /พบ 6 ไฟล์ \(อ่านเนื้อหาได้ 5 ไฟล์\)/);
  assert.doesNotMatch(t, /token/);
  assert.match(g.run('testExtract()'), /ขั้นตอนการตรวจรับวัตถุดิบ/, 'ทดสอบ PDF ก่อน');
  assert.strictEqual(g.run('getAppInfo()').source, 'Google Drive');
});

test('sync จาก Google Drive: PDF/Office ใช้การแปลงไฟล์, ไฟล์ Google อ่านตรง', () => {
  const g = install();
  const r = g.run('syncNow()');
  assert.strictEqual(r.total, 6, 'ไม่รวม shortcut, ถังขยะ และโฟลเดอร์ของระบบ');
  assert.strictEqual(r.extracted, 5);
  const convertedBefore = g.convertCalls.length;
  assert.strictEqual(convertedBefore, 2, 'แปลงเฉพาะ PDF และ docx');

  const docs = indexDocs(g);
  assert.match(docs.pdf1.t, /ตรวจรับวัตถุดิบ 5 ล็อต/);
  assert.strictEqual(docs.pdf1.p, 'Procedure');
  assert.strictEqual(docs.pdf1.b, 'สุภัสสา');
  assert.strictEqual(docs.pdf1.u, 'https://drive.google.com/file/d/pdf1/view');
  assert.match(docs.docx1.t, /purchase request/);
  assert.match(docs.gdoc1.t, /ยินดีต้อนรับ/);
  assert.strictEqual(docs.gdoc1.e, 'gdoc');
  assert.match(docs.gsheet1.t, /torque 45/);
  assert.match(docs.gslides1.t, /robot welding/);
  assert.strictEqual(docs.form1.t, '', 'Google Form ค้นได้จากชื่อเท่านั้น');
  assert.ok(!docs.own1 && !docs.short1 && !docs.trash1);
});

test('Google Drive: ไม่อ่านซ้ำถ้าไม่เปลี่ยน, อ่านใหม่เมื่อแก้ไข', () => {
  const g = install();
  g.run('syncNow()');
  const reads = g.gdrive.reads;
  g.run('syncNow()');
  assert.strictEqual(g.gdrive.reads, reads);

  g.gdrive.edit('pdf1', 'ตรวจรับวัตถุดิบ ฉบับแก้ไข');
  g.gdrive.edit('gdoc1', 'คู่มือฉบับใหม่ 2027');
  const r = g.run('syncNow()');
  assert.strictEqual(r.changed, 2);
  const docs = indexDocs(g);
  assert.match(docs.pdf1.t, /ฉบับแก้ไข/);
  assert.match(docs.gdoc1.t, /ฉบับใหม่ 2027/);
});

test('ไม่มีสิทธิ์เปิดโฟลเดอร์ → แจ้งให้ตรวจสิทธิ์ของบัญชี', () => {
  const g = install();
  g.props.set('SOURCE', 'https://drive.google.com/drive/folders/NO_ACCESS_FOLDER_123');
  assert.throws(() => g.run('testConnection()'), /ตรวจว่าบัญชีที่ติดตั้ง Script มีสิทธิ์/);
});

test('settings รุ่นเดิมได้นามสกุลไฟล์ Google เพิ่มอัตโนมัติ', () => {
  const g = install();
  const sheet = g.ss.getSheetByName('Settings');
  for (let r = 2; r <= sheet.getLastRow(); r++) {
    if (sheet.cell(r, 1) === 'file_extensions') sheet.set(r, 2, 'pdf,docx,doc,xlsx,xls,csv,pptx,ppt,txt');
  }
  assert.ok(g.run('getSettings_().fileExtensions').includes('gdoc'));
});
