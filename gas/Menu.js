/**
 * เมนู "🔎 SearchDocSystem" ใน Google Sheet — ติดตั้ง ตั้งค่า sync และทดสอบ โดยไม่ต้องเข้า Apps Script editor
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔎 SearchDocSystem')
    .addItem('1. ติดตั้ง / อัปเดตระบบ', 'menuInstall')
    .addItem('2. ตั้งค่า (OneDrive, รหัสเข้าเว็บ)', 'menuSettings')
    .addItem('3. เปิดหน้าเว็บ', 'menuOpenWebApp')
    .addSeparator()
    .addItem('Sync ตอนนี้', 'menuSync')
    .addItem('ลองแตกข้อความไฟล์ที่ล้มเหลวใหม่', 'menuRetryFailed')
    .addSubMenu(ui.createMenu('ทดสอบ')
      .addItem('เชื่อมต่อ OneDrive', 'menuTestConnection')
      .addItem('แตกข้อความ (OCR) ไฟล์แรก', 'menuTestExtract'))
    .addToUi();
}

function runMenu_(title, fn) {
  const ui = SpreadsheetApp.getUi();
  try {
    const message = fn();
    if (message) ui.alert(title, message, ui.ButtonSet.OK);
  } catch (err) {
    ui.alert(`${title}: ไม่สำเร็จ`, err.message, ui.ButtonSet.OK);
  }
}

function menuInstall() { runMenu_('ติดตั้งระบบ', () => setup()); }
function menuTestConnection() { runMenu_('ทดสอบเชื่อมต่อ OneDrive', () => testConnection()); }
function menuTestExtract() { runMenu_('ทดสอบแตกข้อความ', () => testExtract()); }

function menuSync() {
  runMenu_('Sync', () => {
    const r = syncNow();
    if (r.skipped) return 'มีงาน sync อื่นกำลังทำงานอยู่ ลองใหม่ในอีกสักครู่';
    return `ไฟล์ทั้งหมด ${r.total}\nใหม่ ${r.added} · เปลี่ยนแปลง ${r.changed} · ย้าย/เปลี่ยนชื่อ ${r.moved} · ถูกลบ ${r.deleted}\n` +
      `แตกข้อความแล้ว ${r.extracted} · ล้มเหลว ${r.failed}` +
      (r.remaining ? `\nเหลืออีก ${r.remaining} ไฟล์ — ระบบจะทำต่อในรอบถัดไป` : '');
  });
}

function menuRetryFailed() {
  runMenu_('ลองใหม่', () => `ใส่ ${retryFailed()} ไฟล์กลับเข้าคิวแล้ว — จะแตกข้อความในรอบ sync ถัดไป`);
}

function menuOpenWebApp() {
  runMenu_('หน้าเว็บ', () => {
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      return 'ยังไม่ได้ Deploy หน้าเว็บ\n\nที่ Apps Script: การทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่ → เว็บแอป\n' +
        'ดำเนินการในฐานะ: ฉัน → ทำให้ใช้งานได้';
    }
    const html = HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif;font-size:14px;line-height:1.6">
        <p>ลิงก์หน้าเว็บ (ส่งให้ผู้ใช้ได้เลย):</p>
        <p><a href="${url}" target="_blank" style="color:#cc785c;word-break:break-all">${url}</a></p>
      </div>`).setWidth(460).setHeight(150);
    SpreadsheetApp.getUi().showModalDialog(html, 'เปิดหน้าเว็บ');
    return '';
  });
}

// ---------- หน้าต่างตั้งค่า ----------

const CONFIG_FIELDS = [
  { key: 'TENANT_ID', label: 'Directory (tenant) ID — จาก IT', secret: false },
  { key: 'CLIENT_ID', label: 'Application (client) ID — จาก IT', secret: false },
  { key: 'CLIENT_SECRET', label: 'Client secret (Value) — จาก IT', secret: true },
  { key: 'SOURCE', label: 'โฟลเดอร์ OneDrive: ลิงก์แชร์ หรือ อีเมล:/path', secret: false },
  { key: 'SITE_PASSWORD', label: 'รหัสเข้าหน้าเว็บค้นหา (ไม่บังคับ)', secret: true },
];

function menuSettings() {
  runMenu_('ตั้งค่า', () => {
    requireOwner_();
    const p = props_();
    const esc = (s) => String(s || '').replace(/[&<>"']/g, (c) => `&#${c.charCodeAt(0)};`);
    const fields = CONFIG_FIELDS.map((f) => {
      const current = p.getProperty(f.key);
      return `<label>${esc(f.label)}
        <input name="${f.key}" type="${f.secret ? 'password' : 'text'}" autocomplete="off"
          value="${f.secret ? '' : esc(current)}" placeholder="${f.secret && current ? 'ตั้งไว้แล้ว — เว้นว่างเพื่อใช้ค่าเดิม' : ''}">
      </label>`;
    }).join('');
    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: sans-serif; font-size: 14px; color: #141413; margin: 0; padding: 4px 2px; }
        label { display: block; margin-bottom: 12px; color: #6c6a64; }
        input { display: block; width: 100%; box-sizing: border-box; height: 36px; margin-top: 4px; padding: 0 10px;
          border: 1px solid #e6dfd8; border-radius: 8px; font-size: 14px; color: #141413; }
        p { color: #6c6a64; font-size: 13px; margin: 0 0 14px; }
        button { height: 36px; padding: 0 18px; border: 0; border-radius: 8px; background: #cc785c; color: #fff; font-size: 14px; cursor: pointer; }
        #msg { margin-left: 10px; font-size: 13px; }
      </style>
      <p>ใส่ค่าจาก IT ทีหลังได้ · ช่องรหัสที่เว้นว่างจะใช้ค่าเดิม · พิมพ์ - เพื่อลบรหัสเข้าเว็บ</p>
      <form id="f">${fields}<button type="submit">บันทึก</button><span id="msg"></span></form>
      <script>
        document.getElementById('f').onsubmit = (e) => {
          e.preventDefault();
          const data = Object.fromEntries(new FormData(e.target));
          document.getElementById('msg').textContent = 'กำลังบันทึก…';
          google.script.run
            .withSuccessHandler(() => google.script.host.close())
            .withFailureHandler((err) => { document.getElementById('msg').textContent = err.message; })
            .saveSettingsFromDialog(data);
        };
      </script>`).setWidth(440).setHeight(470);
    SpreadsheetApp.getUi().showModalDialog(html, 'ตั้งค่า SearchDocSystem');
    return '';
  });
}

/** เรียกจากหน้าต่างตั้งค่า — ช่องที่เว้นว่างจะไม่เปลี่ยนค่าเดิม, "-" = ลบค่า */
function saveSettingsFromDialog(values) {
  requireOwner_();
  const p = props_();
  CONFIG_FIELDS.forEach(({ key }) => {
    const v = String((values && values[key]) || '').trim();
    if (v === '-') p.deleteProperty(key);
    else if (v) p.setProperty(key, v);
  });
  CacheService.getScriptCache().remove('graph_token');
  return true;
}
