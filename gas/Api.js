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
    source: sourceLabel_(getSourceKind_()),
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
