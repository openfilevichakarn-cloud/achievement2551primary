# -*- coding: utf-8 -*-
"""ชีท คุณลักษณะ/ค่านิยม รวม ในไฟล์ต้นฉบับใช้คอลัมน์ช่วยคำนวณจำนวนมาก (คอลัมน์ AF–CF) ผูกกับจำนวนหัวข้อตายตัว
ปรับให้ช่องหัวข้อแต่ละช่องคำนวณตรงจากเทอม 1/2 (ค่าสูงสุดของสองภาคเรียน ซึ่งเท่ากับผล MODE(LARGE()) เดิม)
และช่องช่วยของแต่ละกลุ่มอ้างช่วงหัวข้อโดยตรง เพื่อให้เพิ่ม/ลดหัวข้อในเว็บได้ (js/assess.js)
ใช้: python3 tools/patch_assess.py data/template.json  (และถูกเรียกจาก build_template.py)"""
import json, re, sys

def colnum(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n
def colname(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

# kind -> (ชื่อชีท, [(คอลัมน์หัวข้อแรก, สุดท้าย, ช่องช่วยเดิมของ รวม, ช่องผลลัพธ์ 4 ช่อง)])
KINDS = {
    'คุณลักษณะ': [('C', 'L', 'BT'), ('Q', 'U', 'BZ')],
    'ค่านิยม': [('C', 'N', 'BZ'), ('S', 'W', 'CF')],
}
HELPER = ['AF', 'AG']
STU = "'ข้อมูลนักเรียน-กรอง'!A%d"

def patch(sheets):
    for kind, groups in KINDS.items():
        sh = sheets[kind + ' รวม']; term = sheets[kind + ' 1']; c = sh['cells']
        for k in [k for k in c if colnum(re.match(r'[A-Z]+', k)[0]) >= colnum('AB') and int(re.sub(r'\D', '', k)) >= 5]: del c[k]
        for r in range(5, 50):
            stu = STU % (r - 3)
            for gi, (a, b, old) in enumerate(groups):
                for ci in range(colnum(a), colnum(b) + 1):
                    col = colname(ci); cur = c.get(f'{col}{r}', {})
                    c[f'{col}{r}'] = {**{x: y for x, y in cur.items() if x in ('s', 'u')},
                        'f': f'IF({stu}="","",IF(COUNT(\'{kind} 1\'!{col}{r},\'{kind} 2\'!{col}{r})=0,"",MAX(\'{kind} 1\'!{col}{r},\'{kind} 2\'!{col}{r})))'}
                h = HELPER[gi]; rng = f'{a}{r}:{b}{r}'
                c[f'{h}{r}'] = {'s': term['cells'][f'{h}5'].get('s', 0),
                    'f': f'IF({stu}="","",IFERROR(MODE(LARGE({rng},ROW(INDIRECT("1:"&COUNT({rng}))))),IF(MAX({rng})=0,"",MAX({rng}))))'}
                for x in range(1, 5):   # ช่องผลลัพธ์ที่อ้างช่องช่วยเดิม
                    pass
            for k, v in list(c.items()):
                if int(re.sub(r'\D', '', k)) == r and v.get('f'):
                    for gi, (a, b, old) in enumerate(groups):
                        v['f'] = re.sub(rf'\b{old}{r}\b', f'{HELPER[gi]}{r}', v['f'])
        # ชีทรวมไม่ใช้ความกว้างคอลัมน์ช่วยแล้ว
        sh['dim'] = [sh['dim'][0], max(33, 0)]

if __name__ == '__main__':
    d = json.load(open(sys.argv[1], encoding='utf-8'))
    patch({s['name']: s for s in d['sheets']})
    json.dump(d, open(sys.argv[1], 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
