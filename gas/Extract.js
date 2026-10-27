/**
 * แตกข้อความจากไฟล์ด้วยการแปลงเป็นไฟล์ Google ชั่วคราว (ต้องเปิด Advanced Service: Drive API v3)
 *   PDF          → Google Docs + OCR (ได้ทั้งไฟล์พิมพ์และไฟล์สแกน)
 *   Word         → Google Docs
 *   Excel / CSV  → Google Sheets
 *   PowerPoint   → Google Slides
 *   txt / md     → อ่านตรง
 */

const CONVERT_TARGET = {
  pdf: 'application/vnd.google-apps.document',
  doc: 'application/vnd.google-apps.document',
  docx: 'application/vnd.google-apps.document',
  xls: 'application/vnd.google-apps.spreadsheet',
  xlsx: 'application/vnd.google-apps.spreadsheet',
  csv: 'application/vnd.google-apps.spreadsheet',
  ppt: 'application/vnd.google-apps.presentation',
  pptx: 'application/vnd.google-apps.presentation',
};

function extractText_(blob, ext) {
  if (ext === 'txt' || ext === 'md') return blob.getDataAsString('UTF-8');
  const target = CONVERT_TARGET[ext];
  if (!target) return '';
  const options = ext === 'pdf' ? { ocrLanguage: getSettings_().ocrLanguage } : {};
  const tmp = Drive.Files.create(
    { name: `_tmp_extract.${ext}`, mimeType: target, parents: [props_().getProperty('TMP_FOLDER_ID')] },
    blob,
    options
  );
  try {
    if (target === CONVERT_TARGET.docx) return DocumentApp.openById(tmp.id).getBody().getText();
    if (target === CONVERT_TARGET.xlsx) {
      return SpreadsheetApp.openById(tmp.id).getSheets()
        .map((s) => s.getDataRange().getDisplayValues().map((r) => r.filter(String).join(' ')).filter(Boolean).join('\n'))
        .join('\n');
    }
    return SlidesApp.openById(tmp.id).getSlides()
      .map((slide) => slide.getShapes().map((sh) => {
        try { return sh.getText().asString(); } catch (e) { return ''; }
      }).join('\n'))
      .join('\n');
  } finally {
    try { Drive.Files.remove(tmp.id); } catch (e) { log_('WARN', `ลบไฟล์ชั่วคราว ${tmp.id} ไม่ได้: ${e.message}`); }
  }
}

/** บันทึกข้อความลง Drive /text — ใช้ไฟล์เดิมถ้ามี */
function saveText_(fileRow, text) {
  if (fileRow.text_file_id) {
    try {
      const f = DriveApp.getFileById(fileRow.text_file_id);
      if (!f.isTrashed()) {
        f.setContent(text);
        return f.getId();
      }
    } catch (e) { /* ไฟล์ถูกลบไปแล้ว สร้างใหม่ */ }
  }
  const folder = DriveApp.getFolderById(props_().getProperty('TEXT_FOLDER_ID'));
  return folder.createFile(`${fileRow.id}.txt`, text, MimeType.PLAIN_TEXT).getId();
}

function readText_(textFileId) {
  if (!textFileId) return '';
  try {
    return DriveApp.getFileById(textFileId).getBlob().getDataAsString('UTF-8');
  } catch (e) {
    return '';
  }
}
