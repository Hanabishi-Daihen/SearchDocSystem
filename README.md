# SearchDocSystem

ระบบค้นหาเอกสารจากโฟลเดอร์ OneDrive ขององค์กร โดยใช้ **Google Apps Script** เป็นตัวกลาง ค้นได้ทั้งชื่อไฟล์ โฟลเดอร์ และเนื้อหาในไฟล์ PDF (รวมถึงแบบสแกนภาษาไทย), Word, Excel และ PowerPoint — ใช้งานได้ฟรี

- **การออกแบบระบบ:** [docs/SYSTEM_DESIGN.md](docs/SYSTEM_DESIGN.md)
- **Design system (UI):** [docs/DESIGN.md](docs/DESIGN.md)
- **ต้นแบบหน้าค้นหา:** [prototype/index.html](prototype/index.html) — เปิดใน browser ได้เลย ใช้ข้อมูลตัวอย่าง

Stack: Google Apps Script · Microsoft Graph API · Google Sheets · Google Drive (OCR/แปลงไฟล์)

> ระบบนี้ **ไม่ใช้และไม่เก็บรหัสผ่านของลิงก์แชร์ OneDrive** — เข้าถึงไฟล์ด้วย App ที่ลงทะเบียนใน Microsoft Entra ID แทน (ดูข้อ 1 ในเอกสารออกแบบ) ห้าม commit รหัสผ่าน, client secret หรือ token ลงใน repo
