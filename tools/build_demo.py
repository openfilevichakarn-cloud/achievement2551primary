#!/usr/bin/env python3
"""สร้าง "เว็บตัวอย่าง" (ให้ลูกค้าทดลองก่อนตัดสินใจ) ไว้ในโฟลเดอร์ demo/ ของเว็บจริง (เปิดที่ /demo/)
  python3 tools/build_demo.py              # สร้าง demo/index.html + demo/pp-config.js จาก index.html ล่าสุด
  node tools/make_demo_sample.js           # สร้างข้อมูลนักเรียนสมมติ 40 คน ทุกชั้น -> demo/sample/pN.json
- ใช้ไฟล์ js/ css/ fonts/ img/ และ data/template.json ร่วมกับเว็บจริง (อ้างด้วย ../) จึงต้อง deploy ทั้งโปรเจกต์
- ไม่เชื่อม Supabase (ไม่มีกุญแจ ไม่แตะข้อมูลโรงเรียนจริง) การแก้ไขเก็บในเบราว์เซอร์ของผู้ทดลองเท่านั้น
- มีลายน้ำทุกหน้าจอและทุกหน้าที่พิมพ์ / ไม่มีหน้าผู้ดูแลระบบ
รันใหม่ทุกครั้งที่แก้หน้า index.html (ข้อมูลตัวอย่างใน demo/sample ไม่ถูกลบ)"""
import os, re

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'demo'); os.makedirs(OUT, exist_ok=True)
cfg = open(os.path.join(ROOT, 'js', 'pp-config.js'), encoding='utf-8').read()
grades = re.search(r'GRADES_READY:\s*(\[[^\]]*\])', cfg)[1]
open(os.path.join(OUT, 'pp-config.js'), 'w', encoding='utf-8').write(f"""/* เว็บตัวอย่าง — สร้างโดย tools/build_demo.py (อย่าแก้ไฟล์นี้โดยตรง)
 * ไม่เชื่อมต่อ Supabase: ข้อมูลตัวอย่างโหลดจาก sample/ และการแก้ไขเก็บในเบราว์เซอร์เท่านั้น */
window.PP_CONFIG = {{
  SUPABASE_URL: '',
  SUPABASE_ANON_KEY: '',
  BUILD: '1',
  GRADES_READY: {grades},
  TEMPLATE_URL: '../data/template.json',
  SAMPLE_DIR: 'sample',
  DEMO_SCHOOL: 'โรงเรียนสาธิตตัวอย่าง (เว็บทดลอง)',
  WATERMARK: 'ตัวอย่าง · ห้ามใช้งานจริง',
}};
""")
html = open(os.path.join(ROOT, 'index.html'), encoding='utf-8').read()
html = re.sub(r'(href|src)="((?:js|css)/)', r'\1="../\2', html)
html = html.replace('src="../js/pp-config.js"', 'src="pp-config.js"')
html = html.replace('<title>โปรแกรมแจ้งผลการศึกษา ปพ.5 และ ปพ.6 หลักสูตร 2551</title>',
                    '<title>ทดลองใช้ — โปรแกรมแจ้งผลการศึกษา ปพ.5 และ ปพ.6 หลักสูตร 2551</title>\n<meta name="robots" content="noindex">')
html = re.sub(r'<p class="lead">.*?</p>', '<p class="lead">ทดลองใช้งานโปรแกรมได้ทุกเมนู พร้อมข้อมูลตัวอย่างครบทุกชั้น เพื่อประกอบการตัดสินใจก่อนสั่งซื้อ</p>', html, flags=re.S)
html = html.replace('<a href="register.html">ดูรายละเอียดและสมัครใช้งาน</a>', '<a href="../register.html">ดูรายละเอียดและสั่งซื้อ</a>')
html = re.sub(r'<p id="demo-note".*?</p>', '<p id="demo-note" class="demo"><b>เว็บตัวอย่างสำหรับทดลองใช้</b> — มีข้อมูลนักเรียนสมมติ 40 คน ครบทุกชั้น กด “เข้าสู่ระบบ” ได้เลย (รหัส <code>DEMO</code>) การแก้ไขเก็บในเบราว์เซอร์นี้เท่านั้น ไม่กระทบข้อมูลโรงเรียนจริง</p>', html, flags=re.S)
open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8').write(html)
print('built', OUT)
