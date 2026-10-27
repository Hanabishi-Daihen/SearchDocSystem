/**
 * จำลองบริการของ Google Apps Script (Sheets, Drive, UrlFetch ฯลฯ) และ Microsoft Graph ใน memory
 * เพื่อรันโค้ดใน gas/ ด้วย Node ได้
 */
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const zlib = require('node:zlib');
const crypto = require('node:crypto');

const GAS_DIR = path.join(__dirname, '..', 'gas');
// SearchDocSystem: ปรับจาก tests/gas-harness.js ของ Document_Schematic

class FakeBlob {
  constructor(bytes, contentType = 'application/octet-stream', name = '') {
    this.bytes = Buffer.from(bytes);
    this.contentType = contentType;
    this.name = name;
  }
  getBytes() { return Buffer.from(this.bytes); }
  getDataAsString() { return this.bytes.toString('utf8'); }
  getContentType() { return this.contentType; }
  setContentType(t) { this.contentType = t; return this; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  copyBlob() { return new FakeBlob(this.bytes, this.contentType, this.name); }
}

class FakeRange {
  constructor(sheet, row, col, nr, nc) { Object.assign(this, { sheet, row, col, nr, nc }); }
  getValues() {
    const out = [];
    for (let r = 0; r < this.nr; r++) {
      const line = [];
      for (let c = 0; c < this.nc; c++) line.push(this.sheet.cell(this.row + r, this.col + c));
      out.push(line);
    }
    return out;
  }
  setValues(values) {
    if (values.length !== this.nr || values.some((r) => r.length !== this.nc)) throw new Error('setValues: ขนาดไม่ตรงกับ range');
    if (this.row + this.nr - 1 > this.sheet.maxRows) throw new Error('range อยู่นอกขนาดของ sheet');
    values.forEach((r, i) => r.forEach((v, j) => this.sheet.set(this.row + i, this.col + j, v)));
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.nr; r++) for (let c = 0; c < this.nc; c++) this.sheet.set(this.row + r, this.col + c, '');
    return this;
  }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
}

class FakeSheet {
  constructor(name) { this.name = name; this.data = []; this.maxRows = 1000; }
  getName() { return this.name; }
  cell(r, c) { return (this.data[r - 1] && this.data[r - 1][c - 1] !== undefined) ? this.data[r - 1][c - 1] : ''; }
  set(r, c, v) {
    while (this.data.length < r) this.data.push([]);
    this.data[r - 1][c - 1] = v;
  }
  getLastRow() {
    for (let r = this.data.length; r > 0; r--) if ((this.data[r - 1] || []).some((v) => v !== '' && v !== undefined)) return r;
    return 0;
  }
  getLastColumn() { return this.data.reduce((m, r) => Math.max(m, r.length), 0); }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(_, n) { this.maxRows += n; }
  getRange(row, col, nr = 1, nc = 1) { return new FakeRange(this, row, col, nr, nc); }
  setFrozenRows() {}
  setColumnWidth() {}
  deleteRows(start, n) { this.data.splice(start - 1, n); }
}

class FakeSpreadsheet {
  constructor() { this.sheets = [new FakeSheet('Sheet1')]; }
  getSheetByName(n) { return this.sheets.find((s) => s.name === n) || null; }
  insertSheet(n) { const s = new FakeSheet(n); this.sheets.push(s); return s; }
  getSheets() { return this.sheets.slice(); }
  deleteSheet(s) { this.sheets = this.sheets.filter((x) => x !== s); }
  getUrl() { return 'https://sheets/fake'; }
  getId() { return 'SS'; }
}

/** OneDrive จำลอง: โฟลเดอร์ root เดียว + โฟลเดอร์ย่อยได้ */
class FakeOneDrive {
  constructor() { this.items = new Map(); this.add({ id: 'ROOT', name: 'Docs', folder: true, parent: null }); this.downloads = 0; }
  add({ id, name, folder = false, parent = 'ROOT', content = '', modified = '2026-10-01T00:00:00Z', size }) {
    this.items.set(id, { id, name, folder, parent, content, modified, ctag: 'c1', ver: 1, size });
  }
  edit(id, content, modified) {
    const it = this.items.get(id);
    it.content = content;
    it.ver++;
    it.ctag = 'c' + it.ver;
    if (modified) it.modified = modified;
  }
  rename(id, name) { this.items.get(id).name = name; }
  move(id, parent) { this.items.get(id).parent = parent; }
  remove(id) { this.items.delete(id); }
  toJson(it) {
    const j = {
      id: it.id, name: it.name, webUrl: `https://onedrive/${it.id}`, size: it.size !== undefined ? it.size : Buffer.byteLength(it.content),
      parentReference: { driveId: 'DRIVE' },
      fileSystemInfo: { lastModifiedDateTime: it.modified },
      lastModifiedBy: { user: { displayName: 'สมชาย' } },
    };
    if (it.folder) j.folder = { childCount: 0 };
    else {
      j.cTag = it.ctag;
      j.file = { mimeType: 'application/pdf', hashes: { quickXorHash: crypto.createHash('sha1').update(it.content).digest('base64') } };
    }
    return j;
  }
  handle(url) {
    let m;
    if (url.includes('login.microsoftonline.com')) return { code: 200, body: { access_token: 'tok', expires_in: 3600 } };
    if (url.includes('/shares/') || url.includes('/users/')) return { code: 200, body: this.toJson(this.items.get('ROOT')) };
    if ((m = url.match(/\/items\/([^/?]+)\/children/))) {
      const value = [...this.items.values()].filter((i) => i.parent === m[1]).map((i) => this.toJson(i));
      return { code: 200, body: { value } };
    }
    if ((m = url.match(/\/items\/([^/?]+)\/content/))) {
      if (!this.items.has(m[1])) return { code: 404, body: { error: { message: 'not found' } } };
      return { code: 302, headers: { Location: `https://download/${m[1]}` } };
    }
    if ((m = url.match(/^https:\/\/download\/(.+)$/))) {
      this.downloads++;
      return { code: 200, raw: this.items.get(m[1]).content };
    }
    if ((m = url.match(/\/items\/([^/?]+)/))) {
      const it = this.items.get(m[1]);
      return it ? { code: 200, body: this.toJson(it) } : { code: 404, body: { error: { message: 'itemNotFound' } } };
    }
    throw new Error('FakeOneDrive: ไม่รู้จัก URL ' + url);
  }
}

/** Google Drive จำลอง (แหล่งเอกสาร): โฟลเดอร์ + ไฟล์ทั่วไป + ไฟล์ Google */
class FakeGDrive {
  constructor() {
    this.items = new Map();
    this.add({ id: 'GROOT_FOLDER_ID_0000001', name: 'เอกสาร QA', mimeType: 'application/vnd.google-apps.folder', parent: null });
    this.reads = 0;
  }
  add({ id, name, mimeType = 'application/pdf', parent = 'GROOT_FOLDER_ID_0000001', content = '', modified = '2026-10-01T00:00:00Z', trashed = false }) {
    this.items.set(id, { id, name, mimeType, parent, content, modified, trashed });
  }
  edit(id, content, modified) { Object.assign(this.items.get(id), { content, modified: modified || '2026-10-26T00:00:00Z' }); }
  isNative(it) { return /google-apps\.(document|spreadsheet|presentation)$/.test(it.mimeType); }
  toJson(it) {
    const j = { id: it.id, name: it.name, mimeType: it.mimeType, modifiedTime: it.modified,
      webViewLink: `https://drive.google.com/file/d/${it.id}/view`, lastModifyingUser: { displayName: 'สุภัสสา' } };
    if (!it.mimeType.includes('google-apps')) {
      j.size = String(Buffer.byteLength(it.content));
      j.md5Checksum = crypto.createHash('md5').update(it.content).digest('hex');
    }
    return j;
  }
  list({ q, pageToken }) {
    const parent = q.match(/'([^']+)' in parents/)[1];
    const all = [...this.items.values()].filter((i) => i.parent === parent && !i.trashed);
    const start = Number(pageToken || 0);
    const page = all.slice(start, start + 2); // หน้าละ 2 รายการ เพื่อทดสอบ nextPageToken
    return { files: page.map((i) => this.toJson(i)), nextPageToken: start + 2 < all.length ? String(start + 2) : undefined };
  }
}

/** source: 'src' = โหลด gas/*.js ทีละไฟล์, 'dist' = โหลด dist/Code.gs ที่ build แล้ว */
function createGas({ source = 'src' } = {}) {
  const ss = new FakeSpreadsheet();
  const props = new Map([['SPREADSHEET_ID', 'SS'], ['TENANT_ID', 't'], ['CLIENT_ID', 'c'], ['CLIENT_SECRET', 's'],
    ['SOURCE', 'https://daihen-my.sharepoint.com/:f:/g/personal/x/abc']]);
  const cache = new Map();
  const driveFiles = new Map();
  const converted = new Map(); // ไฟล์ Google ชั่วคราวที่แปลงแล้ว → ข้อความ
  const convertCalls = [];
  const fakeConvertedSheet = (text) => ({ getSheets: () => [{ getDataRange: () => ({ getDisplayValues: () =>
    text.split('\n').map((l) => l.split(',')) }) }] });
  const triggers = [];
  const onedrive = new FakeOneDrive();
  const gdrive = new FakeGDrive();
  const nativeText = (id) => { const it = gdrive.items.get(id); return it && gdrive.isNative(it) ? it.content : null; };
  let idSeq = 0;

  const makeFile = (blob, name) => {
    const id = 'F' + (++idSeq);
    const f = {
      id, blob: blob.copyBlob().setName(name || blob.getName()), trashed: false,
      getId: () => id, getUrl: () => 'https://drive/' + id, getName: () => f.blob.getName(),
      getBlob: () => f.blob.copyBlob(), getSize: () => f.blob.bytes.length,
      setTrashed: (t) => { f.trashed = t; }, isTrashed: () => f.trashed,
      setContent: (c) => { f.blob = new FakeBlob(Buffer.from(c), 'text/plain', f.blob.getName()); return f; },
      sharing: 'PRIVATE', getSharingAccess: () => f.sharing, setSharing: (a) => { f.sharing = a; },
    };
    if (currentFolder) currentFolder.files.push(f);
    driveFiles.set(id, f);
    return f;
  };
  let currentFolder = null;
  const makeFolder = (name) => {
    const id = 'D' + (++idSeq);
    const inFolder = (fn) => { currentFolder = folder; try { return fn(); } finally { currentFolder = null; } };
    const folder = {
      getId: () => id, getUrl: () => 'https://drive/' + id, isTrashed: () => false, name, files: [], sharing: 'PRIVATE',
      createFolder: (n) => makeFolder(n),
      createFile: (a, content, mime) => inFolder(() => (typeof a === 'string' ? makeFile(new FakeBlob(Buffer.from(content), mime, a), a) : makeFile(a))),
      setSharing: (a) => { folder.sharing = a; },
      getFiles: () => { const list = folder.files.slice(); return { hasNext: () => list.length > 0, next: () => list.shift() }; },
    };
    driveFiles.set(id, folder);
    return folder;
  };

  const ctx = {
    console: { log() {}, warn() {}, error: console.error },
    SpreadsheetApp: {
      openById: (id) => (converted.has(id) ? fakeConvertedSheet(converted.get(id))
        : nativeText(id) !== null ? fakeConvertedSheet(nativeText(id)) : ss),
      getActiveSpreadsheet: () => ss,
    },
    SlidesApp: { openById: (id) => ({ getSlides: () => [{ getShapes: () => [
      { getText: () => ({ asString: () => (converted.has(id) ? converted.get(id) : nativeText(id)) }) }, { getText: () => { throw new Error('no text'); } }] }] }) },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: (k) => (props.has(k) ? props.get(k) : null),
      setProperty: (k, v) => props.set(k, String(v)),
      deleteProperty: (k) => props.delete(k),
    }) },
    CacheService: { getScriptCache: () => ({
      get: (k) => (cache.has(k) ? cache.get(k) : null), put: (k, v) => cache.set(k, v), remove: (k) => cache.delete(k),
    }) },
    LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) },
    Session: { getActiveUser: () => ({ getEmail: () => 'me@x' }), getEffectiveUser: () => ({ getEmail: () => 'me@x' }) },
    ScriptApp: {
      getService: () => ({ getUrl: () => '' }),
      getProjectTriggers: () => triggers.slice(),
      deleteTrigger: (t) => triggers.splice(triggers.indexOf(t), 1),
      newTrigger: (fn) => ({ timeBased: () => ({ everyMinutes: (m) => ({ create: () => {
        const t = { fn, m, getHandlerFunction: () => fn, getUniqueId: () => 'trig' + triggers.length };
        triggers.push(t);
        return t;
      } }) }) }),
    },
    Utilities: {
      getUuid: () => crypto.randomUUID(),
      sleep() {},
      Charset: { UTF_8: 'utf8' },
      base64Encode: (b) => Buffer.from(b).toString('base64'),
      base64EncodeWebSafe: (s) => Buffer.from(s, 'utf8').toString('base64').replace(/\+/g, '-').replace(/\//g, '_'),
      newBlob: (data, type, name) => new FakeBlob(Buffer.from(data), type, name),
      gzip: (blob, name) => new FakeBlob(zlib.gzipSync(blob.getBytes()), 'application/x-gzip', name),
      ungzip: (blob) => new FakeBlob(zlib.gunzipSync(blob.getBytes())),
    },
    UrlFetchApp: {
      fetch: (url, opts = {}) => {
        const r = onedrive.handle(url);
        if (r.code === 302 && opts.followRedirects !== false) return ctx.UrlFetchApp.fetch(r.headers.Location);
        const body = r.raw !== undefined ? r.raw : JSON.stringify(r.body || {});
        return {
          getResponseCode: () => r.code,
          getContentText: () => body,
          getHeaders: () => r.headers || {},
          getAllHeaders: () => r.headers || {},
          getBlob: () => new FakeBlob(Buffer.from(body), 'application/pdf'),
        };
      },
    },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', PRIVATE: 'PRIVATE' },
      Permission: { VIEW: 'VIEW' },
      createFolder: (n) => makeFolder(n),
      getFolderById: (id) => driveFiles.get(id),
      getFileById: (id) => {
        const f = driveFiles.get(id);
        if (f) return f;
        const it = gdrive.items.get(id);
        if (it && !gdrive.isNative(it)) {
          return { getBlob: () => { gdrive.reads++; return new FakeBlob(Buffer.from(it.content), it.mimeType, it.name); } };
        }
        throw new Error('no file ' + id);
      },
    },
    Drive: { Files: {
      create: (meta, blob, opts) => {
        const content = blob.getDataAsString();
        if (content.startsWith('FAIL')) throw new Error('conversion failed');
        const id = 'G' + (++idSeq);
        converted.set(id, (opts && opts.ocrLanguage ? `OCR(${opts.ocrLanguage}) ` : '') + content);
        convertCalls.push({ mimeType: meta.mimeType, opts });
        return { id };
      },
      remove: (id) => { converted.delete(id); },
      list: (opts) => gdrive.list(opts),
      get: (id) => {
        const it = gdrive.items.get(id);
        if (!it) throw new Error('File not found: ' + id);
        return gdrive.toJson(it);
      },
    } },
    DocumentApp: { openById: (id) => ({ getBody: () => ({ getText: () => (converted.has(id) ? converted.get(id) : nativeText(id)) }) }) },
    MimeType: { GOOGLE_DOCS: 'application/vnd.google-apps.document', PDF: 'application/pdf', PLAIN_TEXT: 'text/plain' },
    HtmlService: { createHtmlOutputFromFile: (n) => {
      const o = { file: n, setTitle: () => o, addMetaTag: () => o };
      return o;
    } },
  };
  vm.createContext(ctx);
  if (source === 'dist') {
    const file = path.join(__dirname, '..', 'dist', 'Code.gs');  // ต้อง npm run build ก่อน
    vm.runInContext(fs.readFileSync(file, 'utf8'), ctx, { filename: 'Code.gs' });
  } else {
    fs.readdirSync(GAS_DIR).filter((f) => f.endsWith('.js')).sort().forEach((f) => {
      vm.runInContext(fs.readFileSync(path.join(GAS_DIR, f), 'utf8'), ctx, { filename: f });
    });
  }
  // ล้าง cache ค่าตั้งค่าระหว่างการเรียกแต่ละครั้ง (GAS จริงเริ่มใหม่ทุก execution)
  const run = (code) => vm.runInContext(`settingsCache_ = null; ${code}`, ctx);
  return { ctx, ss, props, onedrive, gdrive, triggers, driveFiles, run, convertCalls, cache };
}

module.exports = { createGas, FakeBlob };
