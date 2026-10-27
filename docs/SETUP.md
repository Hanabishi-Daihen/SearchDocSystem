# คู่มือติดตั้ง SearchDocSystem (Google Apps Script)

ใช้เวลาประมาณ 15 นาที (ไม่นับเวลารอ IT) ไม่มีค่าใช้จ่าย

**ภาพรวม:** [1. วางโค้ด](#1-วางโค้ด-3-ไฟล์) → [2. ติดตั้งจากเมนู](#2-ติดตั้งจากเมนูใน-google-sheet) → [3. ตั้งค่า](#3-ตั้งค่า) → [4. Deploy หน้าเว็บ](#4-deploy-หน้าเว็บ) → [5. ทดสอบและ sync](#5-ทดสอบและ-sync) · [6. ขอ App จาก IT](#6-ลงทะเบียน-app-ใน-microsoft-entra-id-ให้-it-admin-ทำ)

---

## 1. วางโค้ด 3 ไฟล์

ไฟล์อยู่ในโฟลเดอร์ **`dist/`** ของ repo (เปิดไฟล์บน GitHub → ปุ่ม **Copy raw file**)

1. เปิด Google Sheet → **ส่วนขยาย → Apps Script**
2. ⚙️ **การตั้งค่าโปรเจ็กต์** → ติ๊ก **แสดงไฟล์ Manifest "appsscript.json" ในเครื่องมือแก้ไข**
3. **ลบไฟล์ .gs อื่นทั้งหมดออกก่อน** (เช่น `Config.gs`, `Repo.gs`, `WebApp.gs` ที่วางไว้ก่อนหน้า) — โค้ดทั้งหมดรวมอยู่ใน `Code.gs` แล้ว ถ้ามีไฟล์ซ้ำจะขึ้น error *"Identifier ... has already been declared"*
4. วางเนื้อหาตามตาราง

| ไฟล์ใน repo | ใน Apps Script | ทำอย่างไร |
|---|---|---|
| `dist/Code.gs` | `Code.gs` | ลบของเดิมทั้งหมด → วาง (ถ้ายังไม่มี: **+ → สคริปต์** ตั้งชื่อ `Code`) |
| `dist/index.html` | `index.html` | ลบของเดิมทั้งหมด → วาง (ชื่อไฟล์ต้องเป็น `index` ตัวพิมพ์เล็ก) |
| `dist/appsscript.json` | `appsscript.json` | ลบของเดิมทั้งหมด → วาง |

5. กด 💾 บันทึก แล้ว **รีเฟรชหน้า Google Sheet**

ผลลัพธ์ในแถบ "ไฟล์" ควรเหลือแค่ `appsscript.json`, `Code.gs`, `index.html` และมี **Drive** ใต้ "บริการ"

---

## 2. ติดตั้งจากเมนูใน Google Sheet

หลังรีเฟรช จะมีเมนู **🔎 SearchDocSystem** ที่แถบเมนูด้านบน

1. **🔎 SearchDocSystem → 1. ติดตั้ง / อัปเดตระบบ**
2. ครั้งแรก Google จะขอสิทธิ์ → เลือกบัญชี → ถ้าขึ้น *"Google ไม่ได้ยืนยันแอปนี้"* กด **ขั้นสูง → ไปที่ … (ไม่ปลอดภัย)** → **อนุญาต** (ปกติสำหรับ Script ที่เขียนเอง)
3. ถ้าอนุญาตแล้วไม่มีอะไรเกิดขึ้น ให้กดเมนูเดิมอีกครั้ง → ขึ้น **✅ ติดตั้งเสร็จแล้ว**

ระบบสร้างให้:
- sheet **Files** (รายการไฟล์และสถานะ), **Settings**, **SearchLog** (สถิติคำค้น), **Logs**
- โฟลเดอร์ **SearchDocSystem/** ใน Google Drive (`text/`, `index/`, `tmp/`)
- sync อัตโนมัติทุก 30 นาที

> ⚠️ อย่าแก้ข้อมูลใน sheet อื่นนอกจาก `Settings` ด้วยมือ — ระบบเขียนทับทุกรอบ sync
> 🔒 ตั้งการแชร์ Google Sheet เป็น **จำกัด** — ผู้ใช้เว็บไม่ต้องมีสิทธิ์เข้า Sheet และใน Sheet มีข้อมูลรายชื่อไฟล์

---

## 3. ตั้งค่า

**🔎 SearchDocSystem → 2. ตั้งค่า**

| ช่อง | ค่า |
|---|---|
| Directory (tenant) ID | จาก IT ([ข้อ 6](#6-ลงทะเบียน-app-ใน-microsoft-entra-id-ให้-it-admin-ทำ)) |
| Application (client) ID | จาก IT |
| Client secret (Value) | จาก IT |
| โฟลเดอร์ OneDrive | ลิงก์แชร์โฟลเดอร์ หรือ `อีเมลเจ้าของ:/path` เช่น `supansa@daihen.co.th:/Documents/ชื่อโฟลเดอร์` |
| รหัสเข้าหน้าเว็บค้นหา | ไม่บังคับ — ถ้าใส่ ผู้ใช้ต้องพิมพ์รหัสนี้ก่อนค้นหา |

- ใส่ค่าจาก IT ทีหลังได้ · ช่องรหัสที่เว้นว่างจะใช้ค่าเดิม · พิมพ์ `-` เพื่อลบค่า
- ถ้าลิงก์แชร์เปิดผ่าน Graph ไม่ได้ (เช่นลิงก์ที่ตั้งรหัสผ่านไว้) ให้ใช้รูปแบบ `อีเมล:/path` แทน — path ดูได้จาก OneDrive ของเจ้าของโฟลเดอร์ (ไม่ต้องใส่ "My files")
- ค่าทั้งหมดเก็บใน Script Properties ไม่อยู่ในโค้ดหรือใน Sheet

---

## 4. Deploy หน้าเว็บ

1. Apps Script → **การทำให้ใช้งานได้ → การทำให้ใช้งานได้รายการใหม่** → ⚙️ **เว็บแอป**
2. ดำเนินการในฐานะ: **ฉัน** · ผู้มีสิทธิ์เข้าถึง: เลือกตามที่ต้องการ
   - **ทุกคน** — ผู้ใช้ไม่ต้องมีบัญชี Google → **ควรตั้ง "รหัสเข้าหน้าเว็บค้นหา" ในข้อ 3** เพราะเนื้อหาเอกสารค้นได้จากเว็บ
   - **ทุกคนที่มีบัญชี Google** — ต้อง login Google ก่อน
3. **ทำให้ใช้งานได้** → คัดลอก URL เว็บแอปให้ผู้ใช้ (ดูภายหลังได้จากเมนู **3. เปิดหน้าเว็บ**)

> หลังวางโค้ดใหม่ ต้อง **การทำให้ใช้งานได้ → จัดการการทำให้ใช้งานได้ → ✏️ → เวอร์ชัน: เวอร์ชันใหม่** เพื่อให้เว็บใช้โค้ดล่าสุด (URL เดิม)

---

## 5. ทดสอบและ sync

เมนู **🔎 SearchDocSystem → ทดสอบ**

| เมนู | ผลที่ควรได้ |
|---|---|
| เชื่อมต่อ OneDrive | ✅ ขอ token สำเร็จ · เปิดโฟลเดอร์ได้ · รายชื่อไฟล์ 15 ไฟล์แรก |
| แตกข้อความ (OCR) ไฟล์แรก | ข้อความช่วงต้นของไฟล์ + สัดส่วนภาษาไทย — ตรวจว่าอ่านภาษาไทยถูกต้อง |

จากนั้น **Sync ตอนนี้** — ครั้งแรกอ่านได้ประมาณ 15–40 ไฟล์ต่อรอบ (ขึ้นกับขนาดไฟล์) ไฟล์ที่เหลือจะทำต่อทุก 30 นาที ระหว่างนี้ไฟล์ที่ยังไม่ได้อ่านเนื้อหาก็ค้นจากชื่อได้แล้ว

| ปัญหา | วิธีแก้ |
|---|---|
| `ขอ token ไม่สำเร็จ: invalid_client` | Client secret ผิด หรือหมดอายุ — ใส่ค่า **Value** (ไม่ใช่ Secret ID) |
| `Graph 403` / `Access denied` | IT ยังไม่ได้ Grant admin consent หรือยังไม่ได้ให้สิทธิ์ site (`Sites.Selected`) |
| `Graph 404` / เปิดลิงก์ไม่ได้ | ใช้รูปแบบ `อีเมล:/path` ในช่องโฟลเดอร์แทนลิงก์ |
| ไฟล์สถานะ `error` ใน sheet Files | ดูคอลัมน์ `last_error` · แก้แล้วใช้เมนู **ลองแตกข้อความไฟล์ที่ล้มเหลวใหม่** |
| PDF ยาว อ่านได้แค่ช่วงต้น | ข้อจำกัด OCR ของ Google Drive (ประมาณ 10 หน้าแรก) — แผนแก้ใน Phase 3 |

สถานะในคอลัมน์ `status` ของ sheet Files: `queued` รออ่าน · `indexed` อ่านแล้ว · `name_only` ค้นได้จากชื่อ (ชนิดไฟล์ไม่รองรับ/ใหญ่เกิน) · `error` อ่านไม่สำเร็จ · `deleted` ถูกลบจาก OneDrive

ค่าใน sheet **Settings**: รอบ sync, นามสกุลไฟล์ที่อ่าน, ขนาดไฟล์สูงสุด, จำนวนตัวอักษรต่อไฟล์ใน index (แก้ `sync_interval_minutes` แล้วรันเมนู 1 ใหม่)

---

## 6. ลงทะเบียน App ใน Microsoft Entra ID (ให้ IT Admin ทำ)

1. [Entra admin center](https://entra.microsoft.com) → **Applications → App registrations → New registration**
   ชื่อ `SearchDocSystem` · Supported account types: **Single tenant** · Redirect URI: เว้นว่าง → Register
2. หน้า Overview: คัดลอก **Application (client) ID** และ **Directory (tenant) ID**
3. **Certificates & secrets → New client secret** → คัดลอกช่อง **Value** (แสดงครั้งเดียว)
4. **API permissions → Add a permission → Microsoft Graph → Application permissions** เลือกแบบใดแบบหนึ่ง
   - **`Sites.Selected`** (แนะนำ — อ่านได้เฉพาะ OneDrive ที่ให้สิทธิ์) แล้วให้สิทธิ์อ่าน OneDrive ของเจ้าของโฟลเดอร์กับ App นี้ (เช่นผ่าน PowerShell `Grant-PnPAzureADAppSitePermission -Permissions Read`)
   - หรือ **`Files.Read.All`** (อ่านได้ทุกไฟล์ในองค์กร)
5. **Grant admin consent for Daihen**

ระบบ **อ่านอย่างเดียว** — ไม่มีโค้ดที่เขียนหรือลบไฟล์ใน OneDrive

---

## สำหรับนักพัฒนา

```bash
npm test        # รันโค้ดกับ Apps Script + Graph จำลอง (tests/gas-harness.js)
npm run build   # gas/ → dist/
```
แก้โค้ดที่ `gas/` เท่านั้น แล้ว build — เทสต์ `build.test.js` จะล้มถ้า `dist/` ไม่ตรงกับซอร์ส
ใช้ [clasp](https://github.com/google/clasp) ได้: คัดลอก `.clasp.json.example` เป็น `.clasp.json` ใส่ Script ID แล้ว `npm run push`
