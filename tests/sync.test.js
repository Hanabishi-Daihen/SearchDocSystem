const test = require('node:test');
const assert = require('node:assert');
const zlib = require('node:zlib');
const { createGas } = require('./gas-harness');

function install(opts) {
  const g = createGas(opts);
  g.run('setup()');
  const od = g.onedrive;
  od.add({ id: 'SUB', name: 'QA', folder: true });
  od.add({ id: 'P1', name: 'ขั้นตอนการตรวจรับวัตถุดิบ Rev.2.pdf', parent: 'SUB', content: 'ตรวจรับวัตถุดิบ ล็อต ๑๒๓', modified: '2026-10-20T00:00:00Z' });
  od.add({ id: 'W1', name: 'แบบฟอร์มขอซื้อ.docx', content: 'Purchase Request ขอซื้อ' });
  od.add({ id: 'X1', name: 'QC Checklist.xlsx', content: 'Torque,45\nWeld,OK' });
  od.add({ id: 'S1', name: 'Machine Setup.pptx', content: 'Robot welding' });
  od.add({ id: 'J1', name: 'photo.jpg', content: 'jpg' });
  od.add({ id: 'BIG', name: 'huge.pdf', content: 'x', size: 100 * 1024 * 1024 });
  return g;
}

function files(g) {
  return g.run(`new Table(getSpreadsheet_(), 'Files').rows`);
}

function byId(g) {
  return Object.fromEntries(files(g).map((f) => [f.id, f]));
}

function indexDocs(g, token) {
  const meta = g.run(`getIndexMeta(${JSON.stringify(token || '')})`);
  const docs = [];
  for (let n = 0; n < meta.chunks; n++) {
    const b64 = g.run(`getIndexChunk(${JSON.stringify(token || '')}, ${JSON.stringify(meta.version)}, ${n})`);
    docs.push(...JSON.parse(zlib.gunzipSync(Buffer.from(b64, 'base64')).toString('utf8')));
  }
  return { meta, docs: Object.fromEntries(docs.map((d) => [d.i, d])) };
}

test('setup สร้าง sheet, โฟลเดอร์ Drive และ trigger และรันซ้ำได้', () => {
  const g = install();
  assert.deepStrictEqual(g.ss.getSheets().map((s) => s.getName()).sort(), ['Files', 'Logs', 'SearchLog', 'Settings']);
  for (const k of ['DRIVE_ROOT_ID', 'TEXT_FOLDER_ID', 'INDEX_FOLDER_ID', 'TMP_FOLDER_ID']) assert.ok(g.props.get(k), k);
  g.run('setup()');
  assert.strictEqual(g.triggers.filter((t) => t.fn === 'runScheduled').length, 1);
  assert.strictEqual(g.triggers[0].m, 30);
});

test('sync ครั้งแรก: แตกข้อความทุกชนิดไฟล์ที่รองรับ และสร้าง index', () => {
  const g = install();
  const r = g.run('syncNow()');
  assert.strictEqual(r.total, 6);
  assert.strictEqual(r.added, 6);
  assert.strictEqual(r.extracted, 4);
  const f = byId(g);
  assert.strictEqual(f.P1.status, 'indexed');
  assert.strictEqual(f.P1.path, 'QA');
  assert.strictEqual(f.J1.status, 'name_only');
  assert.strictEqual(f.BIG.status, 'name_only', 'ไฟล์ใหญ่เกิน max_file_mb ไม่ต้องดาวน์โหลด');
  assert.ok(g.convertCalls.some((c) => c.opts && c.opts.ocrLanguage === 'th'), 'PDF ต้องใช้ OCR ภาษาไทย');

  const { meta, docs } = indexDocs(g);
  assert.strictEqual(meta.count, 6);
  assert.match(docs.P1.t, /ตรวจรับวัตถุดิบ ล็อต 123/, 'เลขไทยถูกแปลงเป็นเลขอารบิกใน index');
  assert.match(docs.W1.t, /purchase request/, 'ข้อความเป็นตัวเล็ก');
  assert.match(docs.X1.t, /torque 45/);
  assert.match(docs.S1.t, /robot welding/);
  assert.strictEqual(docs.J1.t, '');
  assert.strictEqual(docs.P1.n, 'ขั้นตอนการตรวจรับวัตถุดิบ Rev.2.pdf');
  assert.strictEqual(docs.P1.u, 'https://onedrive/P1');
});

test('sync รอบถัดไป: ไม่ดาวน์โหลดซ้ำถ้าไม่มีอะไรเปลี่ยน', () => {
  const g = install();
  g.run('syncNow()');
  const downloads = g.onedrive.downloads;
  const version = g.run('getIndexMeta("")').version;
  const r = g.run('syncNow()');
  assert.strictEqual(r.extracted, 0);
  assert.strictEqual(g.onedrive.downloads, downloads);
  assert.strictEqual(g.run('getIndexMeta("")').version, version, 'ไม่ต้องสร้าง index ใหม่');
});

test('ไฟล์ถูกแก้ / ย้าย / ลบ', () => {
  const g = install();
  g.run('syncNow()');
  const downloads = g.onedrive.downloads;
  g.onedrive.edit('W1', 'ฉบับแก้ไข อนุมัติ 50,000 บาท', '2026-10-25T00:00:00Z');
  g.onedrive.move('X1', 'SUB');
  g.onedrive.rename('X1', 'QC Checklist Rev 4.xlsx');
  g.onedrive.remove('S1');
  const r = g.run('syncNow()');
  assert.strictEqual(r.changed, 1);
  assert.strictEqual(r.moved, 1);
  assert.strictEqual(r.deleted, 1);
  assert.strictEqual(g.onedrive.downloads, downloads + 1, 'ย้าย/เปลี่ยนชื่อไม่ต้องดาวน์โหลดใหม่');

  const { docs } = indexDocs(g);
  assert.match(docs.W1.t, /อนุมัติ 50,000/);
  assert.strictEqual(docs.X1.p, 'QA');
  assert.strictEqual(docs.X1.n, 'QC Checklist Rev 4.xlsx');
  assert.match(docs.X1.t, /torque/, 'ข้อความเดิมยังอยู่');
  assert.ok(!docs.S1, 'ไฟล์ที่ถูกลบไม่อยู่ใน index');
  assert.strictEqual(byId(g).S1.status, 'deleted');
});

test('ไฟล์ที่แตกข้อความไม่สำเร็จ: ลองใหม่ 3 รอบแล้วหยุด จนกว่าไฟล์จะถูกแก้หรือสั่งลองใหม่', () => {
  const g = install();
  g.onedrive.add({ id: 'BAD', name: 'broken.pdf', content: 'FAIL' });
  for (let i = 0; i < 3; i++) g.run('syncNow()');
  let bad = byId(g).BAD;
  assert.strictEqual(bad.status, 'error');
  assert.strictEqual(bad.attempts, '3');
  assert.match(bad.last_error, /conversion failed/);

  const downloads = g.onedrive.downloads;
  g.run('syncNow()');
  assert.strictEqual(g.onedrive.downloads, downloads, 'ไม่ลองซ้ำเอง');
  assert.ok(indexDocs(g).docs.BAD, 'ยังค้นได้จากชื่อไฟล์');

  g.onedrive.edit('BAD', 'ไฟล์ใหม่ใช้ได้แล้ว');
  g.run('syncNow()');
  bad = byId(g).BAD;
  assert.strictEqual(bad.status, 'indexed');

  assert.strictEqual(g.run('retryFailed()'), 0);
});

test('รหัสเข้าเว็บ: ต้อง login ก่อนเรียก API', () => {
  const g = install();
  g.run('syncNow()');
  g.props.set('SITE_PASSWORD', 'secret-code');
  assert.strictEqual(g.run('getAppInfo()').needCode, true);
  assert.throws(() => g.run('getIndexMeta("")'), /SESSION_EXPIRED/);
  assert.throws(() => g.run('login("wrong")'), /รหัสไม่ถูกต้อง/);
  const token = g.run('login("secret-code")');
  assert.ok(token);
  assert.strictEqual(indexDocs(g, token).meta.count, 6);
  assert.match(g.run(`getText(${JSON.stringify(token)}, "W1")`), /Purchase Request/);
  g.run(`logSearch(${JSON.stringify(token)}, "ขอซื้อ", 1)`);
  assert.strictEqual(g.ss.getSheetByName('SearchLog').getLastRow(), 2);
});

test('ผู้ใช้เว็บ (ไม่ได้ login Google) เรียกฟังก์ชันผู้ดูแลไม่ได้', () => {
  const g = install();
  g.ctx.Session.getActiveUser = () => ({ getEmail: () => '' });
  for (const call of ['setup()', 'syncNow()', 'retryFailed()', 'testConnection()', 'runScheduled({})',
    'saveSettingsFromDialog({ SOURCE: "x" })']) {
    assert.throws(() => g.run(call), /เมนูใน Google Sheet/, call);
  }
  // trigger จริงยังทำงานได้
  const uid = g.triggers[0].getUniqueId();
  assert.strictEqual(g.run(`runScheduled({ triggerUid: ${JSON.stringify(uid)} })`).total, 6);
});

test('ตั้งค่า: ช่องว่างใช้ค่าเดิม และ "-" ลบค่า', () => {
  const g = install();
  g.props.set('SITE_PASSWORD', 'a');
  g.run('saveSettingsFromDialog({ SOURCE: "user@daihen.co.th:/Documents/QA", CLIENT_SECRET: "", SITE_PASSWORD: "-" })');
  assert.strictEqual(g.props.get('SOURCE'), 'user@daihen.co.th:/Documents/QA');
  assert.strictEqual(g.props.get('CLIENT_SECRET'), 's');
  assert.ok(!g.props.has('SITE_PASSWORD'));
  // เปลี่ยน SOURCE แล้วต้อง resolve ใหม่
  g.run('syncNow()');
  assert.strictEqual(g.props.get('SOURCE_RESOLVED_FROM'), 'user@daihen.co.th:/Documents/QA');
});

test('OneDrive: ใช้ URL จากแถบที่อยู่ของ browser (shared?listurl=…&id=…) เป็นแหล่งเอกสารได้', () => {
  const g = install();
  g.props.set('SOURCE', 'https://daihenelectric-my.sharepoint.com/shared?listurl=https%3A%2F%2Fdaihenelectric%2Dmy%2Esharepoint%2Ecom%2Fpersonal%2Fsupansa%5Fdaihen%5Fco%5Fth%2FDocuments&id=%2Fpersonal%2Fsupansa%5Fdaihen%5Fco%5Fth%2FDocuments%2FOrder%20Spec%20Control&shareLink=1&ga=1');
  assert.strictEqual(g.run('getSourceKind_()'), 'onedrive');
  assert.match(g.run('testConnection()'), /เปิดโฟลเดอร์ OneDrive ได้/);
  const reqs = g.onedrive.requests.join('\n');
  assert.match(reqs, /\/sites\/daihenelectric-my\.sharepoint\.com:\/personal\/supansa_daihen_co_th\?/);
  assert.match(reqs, /\/drives\/DRIVE\/root:\/Order%20Spec%20Control\?/, 'เลือก library Documents ไม่ใช่ library อื่น');
  assert.doesNotMatch(reqs, /\/shares\//, 'URL จาก browser ไม่ใช่ลิงก์แชร์');
  assert.strictEqual(g.run('syncNow()').total, 6);
});

test('OneDrive: URL จาก browser ที่ไม่มี listurl และรูปแบบ อีเมล:/Documents/… ก็ใช้ได้', () => {
  const g = install();
  g.props.set('SOURCE', 'https://daihenelectric-my.sharepoint.com/my?id=%2Fpersonal%2Fsupansa%5Fdaihen%5Fco%5Fth%2FDocuments%2FOrder%20Spec%20Control');
  assert.match(g.run('testConnection()'), /เปิดโฟลเดอร์ OneDrive ได้/);

  const h = install();
  h.props.set('SOURCE', 'supansa@daihen.co.th:/Documents/Order Spec Control');
  assert.match(h.run('testConnection()'), /เปิดโฟลเดอร์ OneDrive ได้/);
  assert.match(h.onedrive.requests.join('\n'), /root:\/Order%20Spec%20Control\?/, 'ตัด Documents/ ออกแล้วลองใหม่');
});
