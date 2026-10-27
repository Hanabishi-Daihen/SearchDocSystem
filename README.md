# SearchDocSystem

ระบบค้นหาเอกสารจากโฟลเดอร์ **Google Drive** (ไม่ต้องขอสิทธิ์จาก IT) หรือ **OneDrive** ขององค์กร โดยใช้ **Google Apps Script** เป็นตัวกลาง ค้นได้ทั้งชื่อไฟล์ โฟลเดอร์ และเนื้อหาในไฟล์ PDF (รวมถึงแบบสแกนภาษาไทย), Word, Excel, PowerPoint และไฟล์ Google — ใช้งานได้ฟรี

- **ติดตั้ง:** วาง 3 ไฟล์จาก [`dist/`](dist) ลงใน Apps Script ของ Google Sheet แล้วใช้เมนู **🔎 SearchDocSystem** — ดู [docs/SETUP.md](docs/SETUP.md)
- **การออกแบบระบบ:** [docs/SYSTEM_DESIGN.md](docs/SYSTEM_DESIGN.md)
- **Design system (UI):** [docs/DESIGN.md](docs/DESIGN.md) · ต้นแบบหน้าค้นหาพร้อมข้อมูลตัวอย่าง: [prototype/index.html](prototype/index.html)
- **นักพัฒนา (Node 18+):** `npm test` → `npm run build` (ไม่มี dependency ต้องติดตั้ง)

Stack: Google Apps Script · Google Sheets · Google Drive (แหล่งเอกสาร + OCR/แปลงไฟล์) · Microsoft Graph API (เฉพาะ OneDrive)

> ระบบอ่านแหล่งเอกสารอย่างเดียว ไม่แก้หรือลบไฟล์ · OneDrive เข้าถึงด้วย App ที่ลงทะเบียนใน Microsoft Entra ID ผ่าน Graph API · ค่าลับทั้งหมด (client secret, รหัสเข้าเว็บ) เก็บใน Script Properties — ห้าม commit รหัสผ่าน, secret หรือ token ลงใน repo
