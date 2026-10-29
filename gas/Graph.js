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
 * แปลงค่า SOURCE เป็นโฟลเดอร์ใน OneDrive รองรับ 3 แบบ
 *   1) ลิงก์แชร์ OneDrive/SharePoint  https://x-my.sharepoint.com/:f:/g/personal/...
 *   2) URL จากแถบที่อยู่ของ browser ตอนเปิดโฟลเดอร์ (มี id=/personal/.../Documents/โฟลเดอร์)
 *   3) "อีเมลเจ้าของ:/path/โฟลเดอร์" เช่น user@daihen.co.th:/QA — path นับจาก "My files" (ไม่ต้องใส่ Documents)
 */
function resolveSource_(input) {
  input = String(input || '').trim();
  if (!input) throw new Error('ยังไม่ได้ตั้งค่าโฟลเดอร์ OneDrive (เมนู 2. ตั้งค่า)');
  let item;
  const browse = /^https?:\/\//i.test(input) && !/\/:[a-z]:\//i.test(input) && input.match(/[?&]id=([^&#]+)/);
  if (browse) {
    item = resolveBrowserUrl_(input, decodeURIComponent(browse[1].replace(/\+/g, ' ')));
  } else if (/^https?:\/\//i.test(input)) {
    const encoded = 'u!' + Utilities.base64EncodeWebSafe(input, Utilities.Charset.UTF_8).replace(/=+$/, '');
    item = graphJson_(`/shares/${encoded}/driveItem?$select=${ITEM_SELECT}`);
  } else {
    const m = input.match(/^([^\s:/]+@[^\s:/]+):(.*)$/);
    if (!m) throw new Error('รูปแบบโฟลเดอร์ไม่ถูกต้อง: ใส่ลิงก์ OneDrive หรือ อีเมล:/path/โฟลเดอร์');
    const path = m[2].replace(/^\/+|\/+$/g, '');
    const upn = encodeURIComponent(m[1]);
    const byPath = (p) => graphJson_(p
      ? `/users/${upn}/drive/root:/${encodePath_(p)}?$select=${ITEM_SELECT}`
      : `/users/${upn}/drive/root?$select=${ITEM_SELECT}`);
    try {
      item = byPath(path);
    } catch (err) {
      // "Documents" คือตัว drive เอง — ถ้าผู้ใช้ใส่มาด้วยแล้วหาไม่เจอ ลองตัดออก
      if (!/^Graph 404/.test(err.message) || !/^Documents(\/|$)/i.test(path)) throw err;
      item = byPath(path.replace(/^Documents\/?/i, ''));
    }
  }
  if (!item.folder) throw new Error(`"${item.name}" ไม่ใช่โฟลเดอร์`);
  return { driveId: item.parentReference && item.parentReference.driveId, itemId: item.id, name: item.name };
}

/**
 * URL จาก browser: id=/personal/user_co_th/Documents/โฟลเดอร์ (และ listurl=ที่อยู่ของ library ถ้ามี)
 * → หา site → หา drive (library) ที่ตรงกัน → หาโฟลเดอร์ตาม path ใน drive
 */
function resolveBrowserUrl_(url, serverPath) {
  const host = url.match(/^https?:\/\/([^/?#]+)/i)[1];
  serverPath = '/' + serverPath.replace(/^\/+|\/+$/g, '');
  const listUrlParam = url.match(/[?&]listurl=([^&#]+)/i);
  let libraryPath;
  if (listUrlParam) {
    libraryPath = decodeURIComponent(listUrlParam[1]).replace(/^https?:\/\/[^/]+/i, '').replace(/\/+$/, '');
  } else {
    const m = serverPath.match(/^\/(?:personal|sites|teams)\/[^/]+\/[^/]+/i);
    if (!m) throw new Error(`อ่าน path ของโฟลเดอร์จากลิงก์ไม่ได้: ${serverPath}`);
    libraryPath = m[0];
  }
  if (serverPath.toLowerCase().indexOf(libraryPath.toLowerCase()) !== 0) {
    throw new Error('ลิงก์ไม่ถูกรูปแบบ: path ของโฟลเดอร์ไม่อยู่ใน library ที่ระบุ');
  }
  const sitePath = libraryPath.replace(/\/[^/]+$/, '');
  const site = graphJson_(`/sites/${host}:${encodePath_(sitePath)}?$select=id`);
  const norm = (u) => decodeURIComponent(String(u || '')).replace(/^https?:\/\/[^/]+/i, '').replace(/\/+$/, '').toLowerCase();
  const drives = graphJson_(`/sites/${site.id}/drives?$select=id,name,webUrl`).value || [];
  const drive = drives.find((d) => norm(d.webUrl) === libraryPath.toLowerCase());
  if (!drive) throw new Error(`ไม่พบ library "${libraryPath}" ในไซต์ — ตรวจสอบลิงก์`);
  const rel = serverPath.slice(libraryPath.length).replace(/^\/+/, '');
  return graphJson_(rel
    ? `/drives/${drive.id}/root:/${encodePath_(rel)}?$select=${ITEM_SELECT}`
    : `/drives/${drive.id}/root?$select=${ITEM_SELECT}`);
}

function encodePath_(path) {
  return path.split('/').map((p) => (p ? encodeURIComponent(p) : p)).join('/');
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
