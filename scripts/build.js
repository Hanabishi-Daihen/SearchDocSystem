#!/usr/bin/env node
/**
 * รวมซอร์สใน gas/ เป็นไฟล์พร้อมวางใน Apps Script ที่ dist/
 *   dist/Code.gs          โค้ดฝั่ง server ทุกไฟล์ต่อกัน
 *   dist/index.html       หน้าเว็บ
 *   dist/appsscript.json
 *
 * ใช้: npm run build             (เขียนไฟล์)
 *      require('./build').buildAll()  (คืนเนื้อหาไฟล์ — ใช้ในเทสต์)
 */
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..');
const GAS = path.join(ROOT, 'gas');
const DIST = path.join(ROOT, 'dist');

// ลำดับไฟล์ใน Code.gs (ใน GAS ทุกไฟล์อยู่ global scope เดียวกัน ลำดับมีผลแค่การอ่าน)
const SERVER_FILES = [
  'Config.js', 'Repo.js', 'Graph.js', 'Extract.js', 'Sync.js', 'SearchIndex.js',
  'Api.js', 'WebApp.js', 'Menu.js', 'Setup.js',
];

function buildCode() {
  const missing = fs.readdirSync(GAS).filter((f) => f.endsWith('.js') && !SERVER_FILES.includes(f));
  if (missing.length) throw new Error(`เพิ่มไฟล์ใน SERVER_FILES ของ scripts/build.js: ${missing.join(', ')}`);
  const header = '/**\n * SearchDocSystem — ไฟล์นี้สร้างจาก gas/*.js ด้วย `npm run build` ห้ามแก้ไขโดยตรง\n */\n';
  return header + SERVER_FILES
    .map((f) => `\n// ==================== ${f} ====================\n\n` + fs.readFileSync(path.join(GAS, f), 'utf8').trim() + '\n')
    .join('');
}

function buildAll() {
  return {
    'Code.gs': buildCode(),
    'index.html': fs.readFileSync(path.join(GAS, 'ui', 'index.html'), 'utf8'),
    'appsscript.json': fs.readFileSync(path.join(GAS, 'appsscript.json'), 'utf8'),
  };
}

if (require.main === module) {
  fs.mkdirSync(DIST, { recursive: true });
  for (const [name, content] of Object.entries(buildAll())) {
    fs.writeFileSync(path.join(DIST, name), content);
    console.log(`dist/${name}  ${(Buffer.byteLength(content) / 1024).toFixed(0)} KB`);
  }
}

module.exports = { buildAll, DIST };
