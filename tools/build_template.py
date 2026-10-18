#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""แปลงไฟล์ต้นแบบ ปพ.5/ปพ.6 (.xlsb) เป็น data/template.json สำหรับเว็บ
ขั้นตอน: LibreOffice แปลง .xlsb -> .xlsx · ล้างชื่อที่หายตอนแปลง · แทนชื่อ No ด้วยเซลล์เลขที่นักเรียนของแต่ละชีท
        · ดึงพื้นที่พิมพ์จากไฟล์ .xlsb ตรง ๆ · extract_xlsx.py -> JSON
ใช้:  python3 tools/build_template.py "ไฟล์.xlsb" data/template.json [โฟลเดอร์ชั่วคราว]"""
import sys, os, re, json, struct, shutil, zipfile, subprocess, tempfile

# ชื่อ No (เลขที่นักเรียนที่เลือก) เป็นชื่อระดับชีทใน Excel แต่ LibreOffice ทำหายตอนอ่าน .xlsb — ชี้ไปเซลล์เดียวกับต้นฉบับ
NO_CELL = {'ปพ.6 เทอม 1': '$K$8', 'ปพ.6 เทอม 2': '$K$8', 'ปพ.6 รวม': '$K$8',
           'เล่ม ปพ.6 เทอม 1': '$CX$4', 'เล่ม ปพ.6 เทอม 2': '$CX$4', 'เล่ม ปพ.6 รวม': '$CX$4', 'ใบรับรอง': '$BA$7'}

# ข้อความขายสินค้า/บัญชีธนาคาร/เบอร์ติดต่อของผู้จัดทำไฟล์ต้นฉบับ — ไม่ให้ไปปรากฏในเว็บของเรา
SCRUB = {'ข้อมูลพื้นฐาน': ['C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C14'] + ['A%d' % r for r in range(15, 23)],
         'สารบัญ': ['D6', 'D7'], 'คู่มือ': ['A35']}

# ข้อมูลตัวอย่างในไฟล์ต้นฉบับที่ต้องล้างให้ห้องใหม่เริ่มจากว่าง (รายวิชาตัวอย่าง / เครื่องหมาย / ในช่องกิจกรรม)
CLEAR = {'รายวิชา': ['B6', 'C6', 'D6', 'E6', 'F6']}


def colnum(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n


def sheet_paths(z):
    wb = z.read('xl/workbook.xml').decode('utf8'); rels = z.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = {}
    for m in re.finditer(r'<Relationship ([^>]*)/?>', rels):
        a = dict(re.findall(r'(\w+)="([^"]*)"', m.group(1))); rid[a['Id']] = a['Target']
    out = {}
    for m in re.finditer(r'<sheet ([^>]*)/?>', wb):
        a = dict(re.findall(r'([\w:]+)="([^"]*)"', m.group(1)))
        out['xl/' + rid[a['r:id']].lstrip('/').replace('xl/', '')] = a['name'].replace('&amp;', '&')
    return out


def patch_xlsx(src, dst):
    zin = zipfile.ZipFile(src); paths = sheet_paths(zin)
    zout = zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED); fixed = 0
    for it in zin.infolist():
        b = zin.read(it.filename)
        if it.filename == 'xl/workbook.xml':   # openpyxl อ่านไม่ได้ถ้ามี definedName ว่าง
            b = re.sub(r'<definedName [^>]*?></definedName>', '', b.decode('utf8')).encode('utf8')
        nm = paths.get(it.filename)
        if nm in NO_CELL:
            def fx(m):
                nonlocal fixed
                fixed += m.group(2).count('#NAME?'); return m.group(1) + m.group(2).replace('#NAME?', NO_CELL[nm]) + m.group(3)
            b = re.sub(r'(<f[^>]*>)([^<]*)(</f>)', fx, b.decode('utf8')).encode('utf8')
        zout.writestr(it, b)
    zout.close(); print('แทนชื่อ No ใน', fixed, 'สูตร')


# ---- อ่านชื่อ (Print_Area) จาก workbook.bin ของ .xlsb ----
def read_print_areas(xlsb):
    data = zipfile.ZipFile(xlsb).read('xl/workbook.bin')
    def recs(b):
        i = 0
        while i < len(b):
            x = b[i]; i += 1; rid = x
            if x & 0x80: x2 = b[i]; i += 1; rid = (x & 0x7f) | ((x2 & 0x7f) << 7)
            sz = 0; sh = 0
            for _ in range(4):
                x = b[i]; i += 1; sz |= (x & 0x7f) << sh; sh += 7
                if not x & 0x80: break
            yield rid, b[i:i + sz]; i += sz
    def wstr(b, o):
        n = struct.unpack_from('<I', b, o)[0]; return b[o + 4:o + 4 + 2 * n].decode('utf-16le'), o + 4 + 2 * n
    def col(n):
        s = ''; n += 1
        while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
        return s
    sheets, xti, names = [], [], []
    for rid, b in recs(data):
        if rid == 0x9c: o = 8; _, o = wstr(b, o); nm, o = wstr(b, o); sheets.append(nm)
        elif rid == 0x16a:
            c = struct.unpack_from('<I', b, 0)[0]; xti = [struct.unpack_from('<iii', b, 4 + 12 * k) for k in range(c)]
        elif rid == 0x27:
            fl, ch, itab = struct.unpack_from('<IBI', b, 0); nm, o = wstr(b, 9)
            cce = struct.unpack_from('<I', b, o)[0]; names.append((nm, itab, b[o + 4:o + 4 + cce]))
    out = {}
    for nm, itab, rg in names:
        if nm != 'Print_Area' or itab >= len(sheets) or not rg or rg[0] not in (0x3b, 0x5b, 0x7b): continue
        ix, r1, r2, c1, c2 = struct.unpack_from('<HIIHH', rg, 1)
        out[sheets[itab]] = '%s%d:%s%d' % (col(c1 & 0x3fff), r1 + 1, col(c2 & 0x3fff), r2 + 1)
    return out


def main():
    src, dst = sys.argv[1], sys.argv[2]
    tmp = sys.argv[3] if len(sys.argv) > 3 else tempfile.mkdtemp()
    os.makedirs(tmp, exist_ok=True)
    shutil.copy(src, os.path.join(tmp, 'src.xlsb'))
    subprocess.run(['soffice', '--headless', '--convert-to', 'xlsx', '--outdir', os.path.join(tmp, 'lo'), os.path.join(tmp, 'src.xlsb')],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    fixed = os.path.join(tmp, 'fixed.xlsx')
    patch_xlsx(os.path.join(tmp, 'lo', 'src.xlsx'), fixed)
    raw = os.path.join(tmp, 'raw.json')
    subprocess.run([sys.executable, '-I', os.path.join(os.path.dirname(__file__), 'extract_xlsx.py'), fixed, raw, '--skip', ''], check=True)
    d = json.load(open(raw, encoding='utf-8'))
    areas = read_print_areas(src); n = 0
    for s in d['sheets']:
        if s['name'] in areas: s['print']['area'] = areas[s['name']]; n += 1
    for s in d['sheets']:
        for ref in SCRUB.get(s['name'], []):
            if ref in s['cells']: s['cells'][ref].pop('v', None)
    for s in d['sheets']:
        for ref in CLEAR.get(s['name'], []):
            if ref in s['cells']: s['cells'][ref].pop('v', None)
        if s['name'] == 'ปฏิทิน':   # วันเรียนและหมายเหตุวันหยุดของปีการศึกษา 2566 ในไฟล์ต้นฉบับ — ให้ผู้ใช้สร้างปฏิทินของปีตัวเองในหน้า ข้อมูล > ปฏิทิน
            for k, c in s['cells'].items():
                m = re.fullmatch(r'([A-Z]+)(\d+)', k); col, row = m[1], int(m[2])
                if 'f' in c: continue
                if (5 <= row <= 17 and row != 11 and 2 <= colnum(col) <= 32) or k == 'A2' or (21 <= row <= 30 and col in ('M', 'Z')): c.pop('v', None)
        if re.match(r'^(คุณลักษณะ|ค่านิยม) ', s['name']):   # ค่าตัวอย่างของนักเรียนคนแรก ๆ (แถวข้อมูลตั้งแต่ 5)
            for k, c in s['cells'].items():
                if c.get('u') and 'f' not in c and 'v' in c and int(re.sub(r'\D', '', k)) >= 5: c.pop('v')
        if s['name'].startswith('กิจกรรม '):
            for c in s['cells'].values():
                if c.get('u') and c.get('v') == '/': c.pop('v')
    d['source'] = os.path.basename(src)
    json.dump(d, open(dst, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
    print('พื้นที่พิมพ์', n, 'ชีท ->', dst, round(os.path.getsize(dst) / 1e6, 1), 'MB')


main()
