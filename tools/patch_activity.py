# -*- coding: utf-8 -*-
"""กิจกรรม รวม ในไฟล์ต้นฉบับเป็นช่องกรอกเปล่า (ปพ.6 รวมอ่านจากชีทนี้) — ใส่สูตรสรุปทั้งปีจากกิจกรรม 1 และ 2
ผ่านทั้งปี = ผ่านทั้ง 2 ภาคเรียน · ไม่ผ่านทั้งปี = ไม่ผ่านภาคใดภาคหนึ่ง · ชื่อชุมนุม = ของภาคเรียนที่ 2 (ถ้าไม่มีใช้ภาคเรียนที่ 1)
ใช้: python3 tools/patch_activity.py data/template.json  (และถูกเรียกจาก build_template.py)"""
import json, sys

def patch(s):
    c = s['cells']
    def one(src, col, r):
        return f"'กิจกรรม {src}'!{col}{r}"
    for r in range(5, 50):
        name = f"IF($B{r}=\"\",\"\","
        for fail, ok in (('C', 'D'), ('E', 'F'), ('H', 'I'), ('J', 'K')):
            a1, a2 = one(1, fail, r), one(2, fail, r)
            p1, p2 = one(1, ok, r), one(2, ok, r)
            c[f'{fail}{r}'] = {**{k: v for k, v in c.get(f'{fail}{r}', {}).items() if k in ('s',)},
                'f': f'IF($B{r}="","",IF(OR({a1}="/",{a2}="/"),"/",IF({a1}<>"",{a1},IF({a2}<>"",{a2},""))))'}
            c[f'{ok}{r}'] = {**{k: v for k, v in c.get(f'{ok}{r}', {}).items() if k in ('s',)},
                'f': f'IF($B{r}="","",IF(AND({p1}="/",{p2}="/"),"/",IF(AND({p1}<>"",{p1}<>"/"),{p1},IF(AND({p2}<>"",{p2}<>"/"),{p2},""))))'}
        c[f'G{r}'] = {**{k: v for k, v in c.get(f'G{r}', {}).items() if k in ('s',)},
            'f': f'IF($B{r}="","",IF({one(2, "G", r)}<>"",{one(2, "G", r)},{one(1, "G", r)}))'}

if __name__ == '__main__':
    d = json.load(open(sys.argv[1], encoding='utf-8'))
    for s in d['sheets']:
        if s['name'] == 'กิจกรรม รวม': patch(s)
    json.dump(d, open(sys.argv[1], 'w', encoding='utf-8'), ensure_ascii=False, separators=(',', ':'))
