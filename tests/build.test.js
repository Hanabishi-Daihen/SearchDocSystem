const test = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { buildAll, DIST } = require('../scripts/build');
const { createGas } = require('./gas-harness');

test('dist/ ตรงกับซอร์สล่าสุด (รัน npm run build หลังแก้ gas/)', () => {
  for (const [name, content] of Object.entries(buildAll())) {
    assert.strictEqual(fs.readFileSync(path.join(DIST, name), 'utf8'), content, `dist/${name} ไม่ตรง`);
  }
});

test('dist/Code.gs รันได้ทั้งระบบ', () => {
  const g = createGas({ source: 'dist' });
  g.run('setup()');
  g.onedrive.add({ id: 'A', name: 'a.pdf', content: 'สวัสดี' });
  assert.strictEqual(g.run('syncNow()').extracted, 1);
  assert.strictEqual(g.run('doGet()').file, 'index', 'doGet ต้องใช้ชื่อไฟล์ index ให้ตรงกับ index.html');
});

test('ไม่มี client secret ฝังในโค้ด', () => {
  const all = Object.values(buildAll()).join('\n');
  assert.doesNotMatch(all, /client_secret['"]?\s*[:=]\s*['"][A-Za-z0-9~._-]{20,}/i);
});

test('หน้าเว็บเรียกเฉพาะฟังก์ชันที่มีใน Api.js', () => {
  const html = buildAll()['index.html'];
  const called = [...html.matchAll(/call\('(\w+)'/g)].map((m) => m[1]);
  const api = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Api.js'), 'utf8');
  assert.ok(called.length >= 5);
  for (const fn of new Set(called)) assert.match(api, new RegExp(`function ${fn}\\(`), fn);
});
