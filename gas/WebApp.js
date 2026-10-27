/**
 * หน้าเว็บ (Deploy → Web app) — ชื่อไฟล์ HTML ใน Apps Script ต้องตรงกับชื่อที่ส่งให้ createHtmlOutputFromFile
 * (ตัวพิมพ์เล็ก/ใหญ่มีผล: ไฟล์ "index.html" → 'index')
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
