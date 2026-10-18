# ปพ.5 · ปพ.6 ออนไลน์ (หลักสูตร 2551 ประถมศึกษา)

เว็บสำหรับทำสมุดบันทึกการพัฒนาคุณภาพผู้เรียน (ปพ.5) และแบบรายงานรายบุคคล (ปพ.6) แปลงมาจากไฟล์ Excel ต้นแบบ
ชีทและสูตรเหมือนไฟล์เดิม คำนวณในเบราว์เซอร์ด้วย `js/engine.js` กรอกข้อมูลแล้วบันทึกขึ้น Supabase
เข้าใช้งานด้วย **รหัสสิทธิ์โรงเรียน** (ไม่ต้องสมัครอีเมล) หลังบ้านออกรหัสที่ `admin.html`

## ไฟล์สำคัญ
| ไฟล์ | หน้าที่ |
|---|---|
| `index.html` `js/app.js` | หน้าใช้งาน: เข้าสู่ระบบ เลือกชั้น/ห้อง แก้ไขชีท พิมพ์ สำรองข้อมูล |
| `data/template.json` | เทมเพลตที่แปลงจาก Excel (83 ชีท) ใช้ร่วมกันทุกชั้น ข้อมูลของโรงเรียนเก็บแยกเฉพาะช่องที่ผู้ใช้กรอก |
| `js/engine.js` `js/render.js` | คำนวณสูตร Excel และวาดชีทเป็นตาราง HTML |
| `admin.html` `register.html` | หลังบ้านผู้ดูแล / หน้าสมัครและชำระเงิน |
| `sql/schema.sql` | โครงสร้างฐานข้อมูล Supabase ทั้งหมด (รันไฟล์เดียว) |
| `supabase/functions/auto-register` | ตรวจสลิปและออกรหัสอัตโนมัติ (ดู `docs/auto-register.md`) |
| `demo/` | เว็บตัวอย่างให้ลูกค้าทดลอง ใช้งานได้ด้วยตัวเอง (มีไฟล์ครบในโฟลเดอร์) deploy แยกเป็นเว็บเดี่ยวได้ ตั้ง publish directory = `demo` (ไม่ต้อง build) หรือเปิดที่ `/demo/` ของเว็บจริง มีลายน้ำ ไม่เชื่อม Supabase สร้างด้วย `python3 tools/build_demo.py` และ `node tools/make_demo_sample.js` (รันใหม่ทุกครั้งที่แก้โปรแกรมหลักหรือเทมเพลต) |
| `tools/build_template.py` | สร้าง `data/template.json` จากไฟล์ `.xlsb` |
| `tools/verify_engine.js` | ตรวจ engine เทียบค่าที่เก็บในไฟล์ |

## ตั้งค่า Supabase
1. สร้างโปรเจกต์ที่ supabase.com แล้วเปิด SQL Editor รัน `sql/schema.sql`
2. ตั้งรหัสผ่านผู้ดูแล: `select set_admin_password('รหัสผ่านที่ต้องการ');`
3. รัน `sql/keepalive.sql` (ให้ workflow `.github/workflows/keepalive.yml` ping ทุก 3 วัน กันโปรเจกต์ฟรีหยุดพัก)
4. ใส่ Project URL และ anon key ใน `js/pp-config.js` (ถ้าเว้นว่าง เว็บทำงานโหมดทดลอง ใช้รหัส `DEMO` เก็บข้อมูลในเบราว์เซอร์)
5. เปิด `admin.html` เข้าด้วยรหัสผ่านข้อ 2 → ตั้งราคา/บัญชีรับเงิน → ออกรหัสสิทธิ์ให้โรงเรียน

## Deploy บน Cloudflare Pages
- Build command: `bash tools/cf_build.sh`
- Build output directory: `dist`

## สร้างเทมเพลตใหม่จาก Excel
```
python3 tools/build_template.py "ไฟล์ต้นแบบ.xlsb" data/template.json
```
ต้องมี LibreOffice (`soffice`) และ Python (`openpyxl`, `lxml`) สคริปต์จะกู้ชื่อ `No` ที่หายตอนแปลง ดึงพื้นที่พิมพ์ ล้างข้อมูลติดต่อ/บัญชีของผู้จัดทำไฟล์ต้นฉบับ และล้างข้อมูลตัวอย่าง

## ทดสอบในเครื่อง
```
python3 -m http.server 8765
```
เปิด http://127.0.0.1:8765/ (ถ้า `js/pp-config.js` ไม่มี URL/key จะเข้าโหมดทดลองด้วยรหัส `DEMO`)
