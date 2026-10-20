# -*- coding: utf-8 -*-
"""เล่ม ปพ.6 — เตรียมแม่แบบให้กรอกข้อมูลส่วนตัวนักเรียนในเว็บ (ข้อมูล › ข้อมูลส่วนตัว)
 - ล้างข้อความตัวอย่าง 'กรอกที่แผ่นงานข้อมูลเพิ่มเติม คอลัมน์ ..' ที่ติดมากับช่องความเห็นครูในชีท ข้อมูลเพิ่มเติม
 - ล้างวันที่ออกเอกสาร (CX20) ของเล่ม ปพ.6 ที่เป็นค่าเก่าจากไฟล์ต้นฉบับ (เว็บตั้งเป็นวันนี้ให้ และกรอกใหม่ได้ใน ข้อมูล › โรงเรียน)
 - อายุ (ปี/เดือน) ใน เล่ม ปพ.6 ใช้ IFERROR เพื่อไม่แสดงรหัสผิดพลาดเมื่อยังไม่กรอกวันเกิด
ใช้: python3 tools/patch_profile.py data/template.json  (และถูกเรียกจาก build_template.py)"""
import json, re, sys

def patch(sheets, styles=None):
    if styles is not None and 622 < len(styles): styles[622]['n'] = 'd\\ mmm\\ yyyy'   # วันที่ชั่งน้ำหนัก/วัดส่วนสูงในเล่ม ปพ.6 แบบย่อให้พอดีช่อง
    c = sheets['ข้อมูลเพิ่มเติม']['cells']
    for k, v in c.items():
        if isinstance(v.get('v'), str) and v['v'].startswith('กรอกที่แผ่นงานข้อมูลเพิ่มเติม'): v.pop('v')
    for name, d in sheets.items():
        if not name.startswith('เล่ม ปพ.6'): continue
        if 'CX20' in d['cells']: d['cells']['CX20'].pop('v', None)
        for k, v in d['cells'].items():
            f = v.get('f')
            if f and 'DATEDIF' in f and not f.startswith('IFERROR('):
                v['f'] = 'IFERROR(' + f + ',"-")'; v.pop('c', None)

if __name__ == '__main__':
    d = json.load(open(sys.argv[1], encoding='utf-8'))
    patch({s['name']: s for s in d['sheets']}, d['styles'])
    json.dump(d, open(sys.argv[1], 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
