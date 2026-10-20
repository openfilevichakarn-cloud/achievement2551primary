/* อ่านไฟล์รายชื่อนักเรียนจากระบบ DMC (.xlsx) — ไม่ต้องใช้ไลบรารีภายนอก (ใช้ DecompressionStream ของเบราว์เซอร์)
 * parse(file) → { classes: [{ key, label, grade, room, rows: [ { prefix, first, last, ... } ] }] } */
(function (root) {
  'use strict';
  const td = new TextDecoder('utf-8');
  const colNum = (s) => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const unesc = (s) => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(+d)).replace(/&#x([0-9a-f]+);/gi, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&amp;/g, '&');

  async function inflateRaw(bytes) {
    const ds = new DecompressionStream('deflate-raw');
    const out = new Response(new Blob([bytes]).stream().pipeThrough(ds));
    return new Uint8Array(await out.arrayBuffer());
  }
  // แตกไฟล์ zip (xlsx) → Map ชื่อไฟล์ → Uint8Array เฉพาะที่ขอ
  async function unzip(buf, want) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let e = u8.length - 22; while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('ไม่ใช่ไฟล์ .xlsx');
    const n = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true); const out = new Map();
    for (let i = 0; i < n; i++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), nl = dv.getUint16(p + 28, true), el = dv.getUint16(p + 30, true), cl = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
      const name = td.decode(u8.subarray(p + 46, p + 46 + nl)); p += 46 + nl + el + cl;
      if (!want(name)) continue;
      const lnl = dv.getUint16(off + 26, true), lel = dv.getUint16(off + 28, true), start = off + 30 + lnl + lel, data = u8.subarray(start, start + csize);
      out.set(name, method === 0 ? data : await inflateRaw(data));
    }
    return out;
  }
  function sharedStrings(xml) {
    const out = [], re = /<si\b[^>]*>([\s\S]*?)<\/si>/g; let m;
    while ((m = re.exec(xml))) { let s = '', t; const tr = /<t\b[^>]*>([\s\S]*?)<\/t>/g; while ((t = tr.exec(m[1]))) s += t[1]; out.push(unesc(s)); }
    return out;
  }
  function sheetRows(xml, ss) {
    const rows = [], re = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g; let m;
    while ((m = re.exec(xml))) {
      const at = m[1], body = m[2] || '', ref = /\br="([A-Z]+)(\d+)"/.exec(at); if (!ref) continue;
      const t = (/\bt="(\w+)"/.exec(at) || [])[1], v = /<v>([\s\S]*?)<\/v>/.exec(body);
      let val = null;
      if (t === 'inlineStr') { let s = '', x; const tr = /<t\b[^>]*>([\s\S]*?)<\/t>/g; while ((x = tr.exec(body))) s += x[1]; val = unesc(s); }
      else if (v) val = t === 's' ? ss[+v[1]] : t === 'str' || t === 'e' ? unesc(v[1]) : v[1];
      if (val === null || val === undefined) continue;
      const r = +ref[2] - 1, c = colNum(ref[1]) - 1;
      (rows[r] = rows[r] || [])[c] = String(val).trim();
    }
    return rows;
  }
  async function readXlsx(buf) {
    const files = await unzip(buf, (n) => n === 'xl/sharedStrings.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
    const names = [...files.keys()].filter(n => n.startsWith('xl/worksheets/')).sort();
    if (!names.length) throw new Error('ไม่พบแผ่นงานในไฟล์');
    const ss = files.has('xl/sharedStrings.xml') ? sharedStrings(td.decode(files.get('xl/sharedStrings.xml'))) : [];
    return sheetRows(td.decode(files.get(names[0])), ss);
  }

  // หัวคอลัมน์ DMC → ชื่อฟิลด์ (เลขประจำตัวนักเรียน ซ้ำ 2 คอลัมน์: คอลัมน์แรกคือเลข 13 หลัก คอลัมน์ที่สองคือรหัสนักเรียน)
  const HEAD = [
    ['ชั้น', 'cls'], ['ห้อง', 'room'], ['เพศ', 'sex'], ['คำนำหน้าชื่อ', 'prefix'], ['ชื่อ', 'first'], ['นามสกุล', 'last'], ['วันเกิด', 'birth'],
    ['กลุ่มเลือด', 'blood'], ['ศาสนา', 'religion'], ['เชื้อชาติ', 'race'], ['สัญชาติ', 'nation'],
    ['บ้านเลขที่', 'house'], ['หมู่', 'moo'], ['ถนน/ซอย', 'road'], ['ตำบล', 'tambon'], ['อำเภอ', 'amphoe'], ['จังหวัด', 'province'],
    ['คำนำหน้าชื่อผู้ปกครอง', 'gPrefix'], ['ชื่อผู้ปกครอง', 'gFirst'], ['นามสกุลผู้ปกครอง', 'gLast'], ['อาชีพของผู้ปกครอง', 'gJob'], ['ความเกี่ยวข้อง', 'gRel'],
    ['คำนำหน้าชื่อบิดา', 'fPrefix'], ['ชื่อบิดา', 'fFirst'], ['นามสกุลบิดา', 'fLast'], ['อาชีพของบิดา', 'fJob'],
    ['คำนำหน้าชื่อมารดา', 'mPrefix'], ['ชื่อมารดา', 'mFirst'], ['นามสกุลมารดา', 'mLast'], ['อาชีพของมารดา', 'mJob'],
  ];
  function parseRows(rows) {
    let hi = rows.findIndex(r => r && r.includes('คำนำหน้าชื่อ') && r.includes('นามสกุล'));
    if (hi < 0) throw new Error('ไม่ใช่ไฟล์รายชื่อนักเรียนจาก DMC (ไม่พบหัวคอลัมน์ คำนำหน้าชื่อ / นามสกุล)');
    const head = rows[hi], idx = {}, ids = [];
    head.forEach((h, c) => {
      if (!h) return; h = h.trim();
      if (h === 'เลขประจำตัวนักเรียน') { ids.push(c); return; }
      const f = HEAD.find(x => x[0] === h || (x[1] === 'gRel' && h.startsWith('ความเกี่ยวข้อง')));
      if (f && idx[f[1]] === undefined) idx[f[1]] = c;
    });
    idx.pid = ids[0]; idx.code = ids[1] !== undefined ? ids[1] : ids[0];
    const clean = (s) => (s === undefined || s === null || s === '-' || s === '' ? '' : String(s).trim());
    const classes = new Map();
    for (let r = hi + 1; r < rows.length; r++) {
      const row = rows[r]; if (!row || !clean(row[idx.first])) continue;
      const o = {}; for (const k of Object.keys(idx)) o[k] = clean(row[idx[k]]);
      const key = o.cls + '|' + o.room;
      if (!classes.has(key)) { const g = /^ป\.?\s*(\d)$/.exec(o.cls); classes.set(key, { key, label: o.cls + (o.room ? '/' + o.room : ''), cls: o.cls, room: o.room, grade: g ? +g[1] : null, rows: [] }); }
      classes.get(key).rows.push(o);
    }
    return { classes: [...classes.values()].sort((a, b) => a.label.localeCompare(b.label, 'th', { numeric: true })) };
  }
  async function parse(file) {
    const buf = file.arrayBuffer ? await file.arrayBuffer() : file;
    return parseRows(await readXlsx(buf));
  }
  // ฟิลด์ → คอลัมน์ในชีท ข้อมูลนักเรียน (ป้าย เชื้อชาติ/สัญชาติ ตรงกับที่เล่ม ปพ.6 พิมพ์)
  const COL = { pid: 3, code: 6, sex: 7, prefix: 8, first: 9, last: 10, birth: 13, blood: 16, race: 17, nation: 18, religion: 19,
    house: 46, moo: 47, road: 48, tambon: 49, amphoe: 50, province: 51,
    gRel: 38, gPrefix: 40, gFirst: 41, gLast: 42, gJob: 83, fPrefix: 27, fFirst: 28, fLast: 29, fJob: 81, mPrefix: 33, mFirst: 34, mLast: 35, mJob: 82 };
  root.PPDmc = { parse, parseRows, readXlsx, COL };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.PPDmc;
})(typeof window !== 'undefined' ? window : globalThis);
