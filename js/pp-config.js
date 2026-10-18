/* ตั้งค่าการเชื่อมต่อ Supabase ของเว็บรหัสสิทธิ์ (index/admin/register)
 * anon key ออกแบบมาให้เปิดเผยในเว็บได้ ความปลอดภัยอยู่ที่ RLS + ฟังก์ชันใน sql/schema.sql
 * ถ้าเว้นว่าง เว็บจะทำงานใน "โหมดทดลอง" (เก็บข้อมูลในเบราว์เซอร์เครื่องนี้ ใช้รหัส DEMO) */
window.PP_CONFIG = {
  SUPABASE_URL: 'https://jdcbhrkqjgakkxnoxwcd.supabase.co',
  SUPABASE_ANON_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImpkY2Jocmtxamdha2t4bm94d2NkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTIyOTc5NzgsImV4cCI6MjEwNzg3Mzk3OH0.W7oOR7UPe-QnNhwtGtGzh_uoXSI0akE7_va4JjMPOfQ',
  BUILD: '1',
  GRADES_READY: [1, 2, 3, 4, 5, 6],   // ชั้นที่เปิดใช้ (ทุกชั้นใช้เทมเพลตเดียวกัน data/template.json)
  SALES: {
    DEMO_URL: '',
    FACEBOOK: '',
    MESSENGER: '',
    LINE_ID: '',
    PHONE: '',
    PACKAGES: [],
    BANK: null,
  },
};
