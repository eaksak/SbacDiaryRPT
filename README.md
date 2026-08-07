# บธว. Daily Revenue & Expense Report System (SbacDiaryRPT)

ระบบสรุปรายได้ค่าใช้จ่ายประจำวัน บธว. สหกรณ์ร้านค้าและดั่งใจเบเกอรี่ (Professional Daily Diary Report & Analytics System)

---

## 🌟 คุณสมบัติเด่น (Features)
- 📝 **บันทึกรายรับ-รายจ่ายประจำวัน**: รองรับข้อมูลสหกรณ์ร้านค้า และ ดั่งใจเบเกอรี่ คำนวณยอดรวมและกระทบยอดชำระเงินอัตโนมัติ
- 📄 **รายงานรูปแบบมาตรฐาน (Print / PDF)**: แสดงผลรายงาน accounting diary ถอดแบบจากแบบฟอร์ม บธว. สำหรับพิมพ์ลงกระดาษ A4 หรือบันทึกเป็น PDF
- 📊 **Dashboard & Analytics**: แสดงกราฟเปรียบเทียบรายได้-ค่าใช้จ่าย ยอดขายแยกตามช่องทางชำระเงิน (มั่งมี, ถุงเงิน, SCB, กรุงศรี, เงินสด)
- ☁️ **Google Sheets Cloud Integration**: บันทึกข้อมูลขึ้น Google Spreadsheet แบบเรียลไทม์ผ่าน Google Apps Script API พร้อมระบบสำรองข้อมูลในเครื่อง (Local Storage Cache)
- 🔍 **ค้นหาประวัติรายงาน**: กรองข้อมูลตามช่วงวันที่และหน่วยงาน พร้อมสรุปยอดรวม

---

## 🚀 การใช้งานผ่าน GitHub Pages
1. เข้าใช้งานผ่าน URL: `https://<your-github-username>.github.io/SbacDiaryRPT/`
2. ไปที่เมนู **ตั้งค่าระบบ & Cloud** เพื่อกรอก **Google Apps Script Web App URL**

---

## 🛠️ ขั้นตอนการติดตั้ง Google Apps Script (Backend Storage)

1. สร้าง **Google Spreadsheet** ใหม่ใน Google Drive (เช่น ตั้งชื่อ `บธว. Daily Diary Report`)
2. ไปที่เมนู **ส่วนขยาย (Extensions) > Apps Script**
3. คัดลอกโค้ดทั้งหมดจากไฟล์ [`google-apps-script.js`](./google-apps-script.js) ไปวางใน Apps Script Editor
4. กด **การทำให้ใช้งานได้ (Deploy) > การทำให้ใช้งานได้ใหม่ (New deployment)**
5. ตั้งค่าการ Deploy:
   - **เลือกประเภท**: เว็บแอป (Web App)
   - **การเรียกใช้เป็น**: ฉัน (Me)
   - **ผู้ที่มีสิทธิ์เข้าถึง**: ทุกคน (Anyone)
6. กด **Deploy**, อนุมัติสิทธิ์ (Authorize permissions)
7. คัดลอก **Web App URL** มาวางในเมนูตั้งค่าบนเว็บแอป แล้วกด **บันทึก & เชื่อมต่อ**

---

## 📁 โครงสร้างตารางใน Google Sheets
สคริปต์จะสร้าง Sheet ให้โดยอัตโนมัติ 4 หน้า:
1. `DailyReports`: สรุปหัวข้อรายงานประจำวัน (วันที่, สถานะ, เวลาบันทึก)
2. `ReportLines`: รายละเอียดรายรับ-รายจ่าย ต้นทุน และเจ้าหนี้การค้า
3. `Debtors`: รายชื่อและยอดขายเชื่อลูกหนี้
4. `PaymentChannels`: ยอดเงินตามช่องทางชำระเงิน (BAY PromptPay, BAY Transfer, KTB 60/40, SCB, Cash)
