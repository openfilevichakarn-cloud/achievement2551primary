# -*- coding: utf-8 -*-
"""ปก ปพ.5: ชื่อผู้อำนวยการ (AD:BJ) กับตำแหน่ง (ทั้งความกว้าง A:CN) จัดกึ่งกลางคนละตำแหน่ง ทำให้ไม่ตรงกัน
ย้ายตำแหน่งมาอยู่ช่วง AD:BJ เดียวกับชื่อ (และแก้สูตรที่อ้างถึงเซลล์ตำแหน่ง)
ใช้: python3 tools/patch_cover.py data/template.json  (และถูกเรียกจาก build_template.py)"""
import json, re, sys

def patch(sheets):
    moved = {}   # 'ชื่อชีท!A56' -> 'AD56'
    for name, d in sheets.items():
        if not name.startswith('ปก ปพ.5 '): continue
        c = d['cells']
        for k, v in list(c.items()):
            m = re.fullmatch(r'A(\d+)', k)
            if not m: continue
            r = int(m[1]); f = v.get('f') or ''
            if not (f == 'ข้อมูลพื้นฐาน!B10' or re.fullmatch(r"'ปก ปพ\.5 รายชั้น [^']+'!A\d+", f)): continue
            if f'A{r}:CN{r}' not in d['merges'] or f'AD{r-1}:BJ{r-1}' not in d['merges']: continue
            d['merges'] = [x for x in d['merges'] if x != f'A{r}:CN{r}'] + [f'AD{r}:BJ{r}']
            nv = dict(v); nv['s'] = c.get(f'AD{r-1}', {}).get('s', v.get('s')); c[f'AD{r}'] = nv   # สไตล์เดียวกับช่องชื่อ (ไม่มีเส้นกรอบ)
            c[k] = {'s': c.get(f'A{r-1}', {}).get('s', v.get('s'))}   # คอลัมน์ A คงเส้นกรอบด้านซ้ายของปก
            moved[(name, f'A{r}')] = f'AD{r}'
    # ช่องสี่เหลี่ยม อนุมัติ/ไม่อนุมัติ ให้ขนาดเท่ากันทุกปก: กว้าง 5 คอลัมน์เท่ากัน (Y:AC, BF:BJ) สูงเท่ากันเมื่อพิมพ์ (ปกรายชั้น 1149px / รายวิชา 1104px)
    for name, d in sheets.items():
        if re.match(r'ปก ปพ\.5 รายชั้น ', name):
            for c in list(range(25, 30)) + list(range(58, 63)):   # ช่องละ 12/13px รวม 63px ให้สัดส่วนตรงกับปกรายวิชา
                d['cw'][str(c)] = {**d['cw'].get(str(c), {}), 'w': 1.08 if c in (25, 26, 27, 58, 59, 60) else 1.0}
            if '52' in d['rh']: d['rh']['52'] = {**d['rh']['52'], 'h': 47.25}   # สูง 63px = กว้างช่อง → สี่เหลี่ยมจัตุรัส
        elif re.match(r'ปก ปพ\.5 รายวิชา ', name) and '36' in d['rh']: d['rh']['36'] = {**d['rh']['36'], 'h': 45.0}   # กว้างช่อง 60px
    for name, d in sheets.items():
        for v in d['cells'].values():
            f = v.get('f')
            if not f: continue
            for (sn, ref), new in moved.items():
                f = re.sub(rf"('{re.escape(sn)}'!)\$?{ref}\b", rf"\g<1>{new}", f)
                if name == sn: f = re.sub(rf"(?<![A-Z!'])\$?{ref}\b(?!:)", new, f) if f.startswith(ref) else f
            if f != v['f']: v['f'] = f; v.pop('c', None)

if __name__ == '__main__':
    d = json.load(open(sys.argv[1], encoding='utf-8'))
    patch({s['name']: s for s in d['sheets']})
    json.dump(d, open(sys.argv[1], 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
