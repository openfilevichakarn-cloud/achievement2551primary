# -*- coding: utf-8 -*-
"""ปพ.6 (เทอม 1/2/รวม): หัวคอลัมน์ "เวลาเรียน (ชั่วโมง/ปี)" ให้ขึ้น 2 บรรทัดพอดี ไม่ตัดเป็น "(ชั่วโมง/" "ปี)"
ขยายคอลัมน์ D (เวลาเรียน) จาก 8.69 เป็น 12.4 (คอลัมน์ C รายวิชาคงเดิม เพื่อให้ชื่อวิชายาวไม่ถูกตัด)
ใช้: python3 tools/patch_p6.py data/template.json  (และถูกเรียกจาก build_template.py)"""
import json, sys

def patch(sheets):
    for name, d in sheets.items():
        if not name.startswith('ปพ.6 '): continue
        cw = d['cw']
        if '3' in cw and '4' in cw and cw['4']['w'] < 10:
            cw['4']['w'] = 12.4

if __name__ == '__main__':
    p = sys.argv[1]; d = json.load(open(p, encoding='utf-8'))
    patch({s['name']: s for s in d['sheets']})
    json.dump(d, open(p, 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
