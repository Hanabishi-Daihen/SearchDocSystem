/**
 * แหล่งเอกสาร — ระบบดูจากค่า SOURCE เองว่าเป็นแบบไหน
 *   Google Drive : ลิงก์โฟลเดอร์ https://drive.google.com/drive/folders/<ID> หรือ ID ของโฟลเดอร์
 *                  อ่านด้วยสิทธิ์ของบัญชีที่ติดตั้ง Script ไม่ต้องใช้ App จาก IT
 *   OneDrive     : ลิงก์แชร์ OneDrive, URL จาก browser หรือ อีเมล:/path — ต้องมี App ใน Entra ID (Graph.js)
 *
 * ทุกแบบคืนรายการไฟล์ในรูปแบบเดียวกัน:
 *   { id, name, ext, path, size, modified, modifiedBy, webUrl, changeKey }
 */

const GOOGLE_NATIVE_EXT = {
  'application/vnd.google-apps.document': 'gdoc',
  'application/vnd.google-apps.spreadsheet': 'gsheet',
  'application/vnd.google-apps.presentation': 'gslides',
};
const GDRIVE_FIELDS = 'nextPageToken, files(id, name, mimeType, size, modifiedTime, md5Checksum, webViewLink, lastModifyingUser(displayName))';

function getSourceKind_() {
  const input = (props_().getProperty('SOURCE') || '').trim();
  return parseDriveFolderId_(input) ? 'gdrive' : 'onedrive';
}

/** ID ของโฟลเดอร์ Google Drive จากลิงก์หรือ ID ตรง ๆ — ไม่ใช่ Google Drive คืน '' */
function parseDriveFolderId_(input) {
  const s = String(input || '').trim();
  let m = s.match(/^https:\/\/drive\.google\.com\/.*(?:\/folders\/|[?&]id=)([-\w]{10,})/);
  if (m) return m[1];
  m = s.match(/^[-\w]{20,}$/);
  return m ? m[0] : '';
}

function sourceLabel_(kind) {
  return kind === 'gdrive' ? 'Google Drive' : 'OneDrive';
}

/** ข้อมูลโฟลเดอร์ต้นทาง { kind, name } — ตรวจว่าเปิดได้ */
function describeSource_() {
  const kind = getSourceKind_();
  if (kind === 'onedrive') return { kind, name: getSource_().name };
  const id = parseDriveFolderId_(props_().getProperty('SOURCE'));
  let folder;
  try {
    folder = Drive.Files.get(id, { fields: 'id, name, mimeType', supportsAllDrives: true });
  } catch (err) {
    throw new Error(`เปิดโฟลเดอร์ Google Drive ไม่ได้ — ตรวจว่าบัญชีที่ติดตั้ง Script มีสิทธิ์ดูโฟลเดอร์นี้ (${err.message})`);
  }
  if (folder.mimeType !== 'application/vnd.google-apps.folder') throw new Error(`"${folder.name}" ไม่ใช่โฟลเดอร์`);
  return { kind, name: folder.name };
}

function listSource_() {
  return getSourceKind_() === 'gdrive' ? listDriveFolder_() : listOneDrive_();
}

/** แตกข้อความของไฟล์ในแหล่งเอกสาร */
function extractFromSource_(row) {
  if (getSourceKind_() === 'gdrive') {
    if (['gdoc', 'gsheet', 'gslides'].includes(row.ext)) return readGoogleFile_(row.id, row.ext);
    return extractText_(DriveApp.getFileById(row.id).getBlob(), row.ext);
  }
  const source = getSource_();
  return extractText_(downloadItem_(source.driveId, row.id), row.ext);
}

// ---------- Google Drive ----------

function listDriveFolder_() {
  const rootId = parseDriveFolderId_(props_().getProperty('SOURCE'));
  const ownFolder = props_().getProperty('DRIVE_ROOT_ID'); // โฟลเดอร์ของระบบเอง ห้ามนำมา index
  const out = [];
  const queue = [{ id: rootId, path: '' }];
  let folders = 0;
  while (queue.length) {
    const { id, path } = queue.shift();
    if (id === ownFolder) continue;
    if (++folders > 5000) throw new Error('โฟลเดอร์ย่อยมากเกิน 5,000 โฟลเดอร์');
    let pageToken;
    do {
      const page = Drive.Files.list({
        q: `'${id}' in parents and trashed = false`,
        fields: GDRIVE_FIELDS,
        pageSize: 1000,
        pageToken,
        supportsAllDrives: true,
        includeItemsFromAllDrives: true,
      });
      (page.files || []).forEach((f) => {
        if (f.mimeType === 'application/vnd.google-apps.folder') {
          queue.push({ id: f.id, path: path ? `${path}/${f.name}` : f.name });
          return;
        }
        if (f.mimeType === 'application/vnd.google-apps.shortcut') return; // ข้าม shortcut กันไฟล์ซ้ำ
        const native = GOOGLE_NATIVE_EXT[f.mimeType];
        out.push({
          id: f.id,
          name: f.name,
          ext: native || extOf_(f.name),
          path,
          size: Number(f.size) || 0,
          modified: f.modifiedTime || '',
          modifiedBy: (f.lastModifyingUser && f.lastModifyingUser.displayName) || '',
          webUrl: f.webViewLink || `https://drive.google.com/file/d/${f.id}/view`,
          changeKey: f.md5Checksum || f.modifiedTime || '',
        });
      });
      pageToken = page.nextPageToken;
    } while (pageToken);
  }
  return out;
}

/** อ่านข้อความจาก Google Docs / Sheets / Slides โดยตรง ไม่ต้องแปลงไฟล์ */
function readGoogleFile_(id, ext) {
  if (ext === 'gdoc') return DocumentApp.openById(id).getBody().getText();
  if (ext === 'gsheet') return sheetText_(SpreadsheetApp.openById(id));
  return slidesText_(SlidesApp.openById(id));
}

// ---------- OneDrive (Graph) ----------

function listOneDrive_() {
  return listAllFiles_(getSource_()).map(({ item, path }) => ({
    id: item.id,
    name: item.name,
    ext: extOf_(item.name),
    path,
    size: Number(item.size) || 0,
    modified: (item.fileSystemInfo && item.fileSystemInfo.lastModifiedDateTime) || '',
    modifiedBy: (item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName) || '',
    webUrl: item.webUrl || '',
    changeKey: item.cTag || item.eTag || `${item.fileSystemInfo && item.fileSystemInfo.lastModifiedDateTime}|${item.size}`,
  }));
}

function extOf_(name) {
  return (String(name).match(/\.([A-Za-z0-9]{1,6})$/) || ['', ''])[1].toLowerCase();
}
