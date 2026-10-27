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
