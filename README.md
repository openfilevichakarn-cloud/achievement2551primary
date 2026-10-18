# ปพ.5 · ปพ.6 ออนไลน์ (หลักสูตร 2551 ประถมศึกษา)

เว็บสำหรับทำสมุดบันทึกการพัฒนาคุณภาพผู้เรียน (ปพ.5) และแบบรายงานรายบุคคล (ปพ.6) หน้ากรอกข้อมูลเป็นเว็บ หน้าพิมพ์จัดรูปแบบตามไฟล์ Excel ของโรงเรียน (A4 แนวตั้ง)

เป็นเว็บแบบ static (HTML + JavaScript ไม่มีขั้นตอน build) ข้อมูลเก็บใน Supabase

## ตั้งค่า Supabase
1. สร้างโปรเจกต์ที่ supabase.com
2. เปิด SQL Editor แล้วรันไฟล์ `supabase/schema.sql`
3. คัดลอก Project URL และ anon key (Project Settings > API) ไปใส่ใน `js/config.js`
4. Authentication > Providers เปิด Email (ถ้าไม่อยากยืนยันอีเมลตอนสมัคร ปิด “Confirm email”)

ถ้าเว้น `js/config.js` ว่าง เว็บจะทำงานในโหมดทดลอง เก็บข้อมูลในเบราว์เซอร์เครื่องนั้นเท่านั้น

## Deploy บน Cloudflare Pages
เชื่อม repo นี้ที่ Workers & Pages > Create > Pages > Connect to Git
- Framework preset: None
- Build command: (ว่าง)
- Build output directory: `/`

## ทดสอบในเครื่อง
```
python -m http.server 8765
```
แล้วเปิด http://127.0.0.1:8765/

## โครงสร้าง
- `js/calc.js` กติกาคำนวณ 2551 (ร้อยละ → เกรด 0–4, ผ/มผ ตัวชี้วัด, ฐานนิยมคุณลักษณะ/ค่านิยม)
- `js/print.js` แบบพิมพ์ทั้งหมด
- `js/app.js` หน้ากรอกข้อมูล
- `js/store.js` เชื่อม Supabase / โหมดทดลอง
- ข้อมูลแต่ละห้องเก็บเป็น JSON 1 แถวในตาราง `classrooms`
