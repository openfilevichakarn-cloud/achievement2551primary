/* เพิ่ม/ลด/ตั้งชื่อหัวข้อของ คุณลักษณะ · ค่านิยม (2 กลุ่มต่อชีท)
 * ชีทในไฟล์ต้นฉบับมีจำนวนคอลัมน์หัวข้อตายตัว — โมดูลนี้ปรับ "แม่แบบ" ก่อนสร้างเวิร์กบุ๊ก:
 * แทรก/ลบคอลัมน์หัวข้อ (ทั้ง เทอม 1 · เทอม 2 · รวม พร้อมกัน) แล้วปรับสูตรที่อ้างถึงทุกชีท (รวมถึงปก ปพ.5 และ ปพ.6)
 * ค่าตั้งค่าเก็บในเอกสาร doc.items — ใช้ซ้ำได้ทุกครั้งที่เปิดห้อง */
(function (root) {
  'use strict';
  const colNum = (s) => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const numCol = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  const parseRef = (k) => { const m = /^([A-Z]+)(\d+)$/.exec(k); return { c: colNum(m[1]), r: +m[2] }; };

  // base = จำนวนช่องหัวข้อในแม่แบบเดิม · def = จำนวนตั้งต้นที่แสดง · origG = ชื่อกลุ่มในแม่แบบ
  const KINDS = {
    trait: { prefix: 'คุณลักษณะ', base: [10, 5], def: [8, 5], max: [20, 20], origG: ['คุณลักษณะอันพึงประสงค์', 'การอ่าน คิดวิเคราะห์ และเขียนสื่อความ'], p6rows: [39, 40], covers: [['AB39', 0]] },
    value: { prefix: 'ค่านิยม', base: [12, 5], def: [12, 5], max: [20, 20], origG: ['ค่านิยมหลัก 12 ประการ', 'สมรรถนะสำคัญของผู้เรียน'], p6rows: [41, 42], covers: [['AB43', 0], ['BF43', 1]] },
  };
  const SUFFIX = ['1', '2', 'รวม'], RES = 4;
  const sheetsOf = (kind) => SUFFIX.map(s => KINDS[kind].prefix + ' ' + s);

  // ตำแหน่งคอลัมน์ตามจำนวนหัวข้อ n = [n1, n2]
  function layout(kind, n) {
    const items = [[], []], res = [[], []]; let c = 3;
    for (let g = 0; g < 2; g++) {
      for (let i = 0; i < n[g]; i++) items[g].push(c++);
      for (let i = 0; i < RES; i++) res[g].push(c++);
    }
    return { items, res, last: c - 1 };
  }

  // ---------------------------------------------------------------- เขียนสูตรใหม่เมื่อแทรก/ลบคอลัมน์
  const SHEET = "(?:(?:'([^']+)'|([A-Za-z0-9_\\u0E00-\\u0E7F]+))!)?";
  const RE_REF = new RegExp("(?<![A-Za-z0-9_\\u0E00-\\u0E7F.$'])" + SHEET + "(\\$?)([A-Z]{1,3})(\\$?)(\\d+)(?::(\\$?)([A-Z]{1,3})(\\$?)(\\d+))?(?![A-Za-z0-9_(])", 'g');
  const RE_COLS = new RegExp("(?<![A-Za-z0-9_\\u0E00-\\u0E7F.$'])" + SHEET + "(\\$?)([A-Z]{1,3}):(\\$?)([A-Z]{1,3})(?![A-Za-z0-9_(!])", 'g');
  // op: { t: 'ins'|'del'|'clone', p, targets:Set(ชื่อชีทตัวพิมพ์เล็ก) }
  // ins = แทรกคอลัมน์ใหม่ต่อจากคอลัมน์ p (ช่วงที่ลงท้ายที่ p ขยายตาม) · del = ลบคอลัมน์ p · clone = ย้ายการอ้างคอลัมน์ p ไปคอลัมน์ p+1
  function mapSingle(c, op) {
    if (op.t === 'clone') return c === op.p ? c + 1 : c;
    if (op.t === 'ins') return c > op.p ? c + 1 : c;
    return c > op.p ? c - 1 : c;
  }
  function mapRange(a, b, op) {
    if (op.t === 'clone') return a === op.p && b === op.p ? [a + 1, b + 1] : [a, b];
    if (op.t === 'ins') return [a > op.p ? a + 1 : a, (b > op.p || (b === op.p && a < op.p)) ? b + 1 : b];
    if (a === b && a === op.p) return [a, b];
    return [a > op.p ? a - 1 : a, b >= op.p ? b - 1 : b];
  }
  function rewrite(f, ctx, op) {
    const parts = f.split(/("(?:[^"]|"")*")/);
    for (let i = 0; i < parts.length; i += 2) {
      let s = parts[i];
      s = s.replace(RE_REF, (m, q1, q2, d1, c1, d2, r1, d3, c2, d4, r2) => {
        const sh = (q1 || q2 || ctx).toLowerCase(); if (!op.targets.has(sh)) return m;
        const prefix = q1 ? "'" + q1 + "'!" : q2 ? q2 + '!' : '';
        if (c2 === undefined) return prefix + d1 + numCol(mapSingle(colNum(c1), op)) + d2 + r1;
        const [a, b] = mapRange(colNum(c1), colNum(c2), op);
        return prefix + d1 + numCol(a) + d2 + r1 + ':' + d3 + numCol(b) + d4 + r2;
      });
      s = s.replace(RE_COLS, (m, q1, q2, d1, c1, d2, c2) => {
        const sh = (q1 || q2 || ctx).toLowerCase(); if (!op.targets.has(sh)) return m;
        const prefix = q1 ? "'" + q1 + "'!" : q2 ? q2 + '!' : '';
        const [a, b] = mapRange(colNum(c1), colNum(c2), op);
        return prefix + d1 + numCol(a) + ':' + d2 + numCol(b);
      });
      parts[i] = s;
    }
    return parts.join('');
  }

  // ---------------------------------------------------------------- เตรียมสำเนาแม่แบบ (copy-on-write)
  const refCache = new WeakMap();
  function relevantSheets(tpl) {   // ชีทที่มีสูตรอ้างถึง คุณลักษณะ/ค่านิยม
    if (refCache.has(tpl)) return refCache.get(tpl);
    const out = new Set();
    for (const s of tpl.sheets) {
      if (/^(คุณลักษณะ|ค่านิยม) /.test(s.name)) { out.add(s.name); continue; }
      for (const c of Object.values(s.cells)) if (c.f && (c.f.indexOf('คุณลักษณะ ') >= 0 || c.f.indexOf('ค่านิยม ') >= 0)) { out.add(s.name); break; }
    }
    refCache.set(tpl, out); return out;
  }
  function shiftMerge(m, op) {
    const [x, y] = m.split(':'), A = parseRef(x), B = parseRef(y || x);
    if (op.t === 'del' && A.c === B.c && A.c === op.p) return null;
    const [a, b] = mapRange(A.c, B.c, op);
    return numCol(a) + A.r + ':' + numCol(b) + B.r;
  }
  function colOp(W, kind, op) {   // W: Map ชื่อ -> def (สำเนาใช้แก้ได้)
    const names = sheetsOf(kind), targets = new Set(names.map(n => n.toLowerCase()));
    const o = { ...op, targets };
    // 1) เขียนสูตรทุกชีทที่เกี่ยวข้อง
    for (const [name, d] of W) {
      for (const [k, c] of Object.entries(d.cells)) {
        if (!c.f) continue;
        const nf = rewrite(c.f, name, o);
        if (nf !== c.f) d.cells[k] = { ...c, f: nf };
      }
    }
    // 2) ย้ายเซลล์/ผสาน/ความกว้างของชีทหัวข้อ
    for (const name of names) {
      const d = W.get(name), p = op.p, nc = {};
      for (const [k, c] of Object.entries(d.cells)) {
        const a = parseRef(k);
        if (op.t === 'ins') nc[a.c > p ? numCol(a.c + 1) + a.r : k] = c;
        else if (a.c === p) continue;
        else nc[a.c > p ? numCol(a.c - 1) + a.r : k] = c;
      }
      if (op.t === 'ins') {   // คอลัมน์ใหม่ = สำเนาของคอลัมน์ p
        const q = numCol(p + 1), ctx = { ...o, t: 'clone' };
        for (const [k, c] of Object.entries(d.cells)) {
          const a = parseRef(k); if (a.c !== p) continue;
          const cc = { ...c };
          if (cc.f) cc.f = rewrite(cc.f, name, ctx);
          else if (a.r >= 5) delete cc.v;
          nc[q + a.r] = cc;
        }
      }
      d.cells = nc;
      const mg = [];
      for (const m of d.merges || []) { const x = shiftMerge(m, o); if (x) mg.push(x); }
      if (op.t === 'ins') for (const m of d.merges || []) {   // ผสานแนวตั้งของช่องหัวข้อ (เช่น K3:K4) ทำซ้ำให้คอลัมน์ใหม่
        const [x, y] = m.split(':'), A = parseRef(x), B = parseRef(y || x);
        if (A.c === op.p && B.c === op.p) mg.push(numCol(op.p + 1) + A.r + ':' + numCol(op.p + 1) + B.r);
      }
      d.merges = mg;
      const cw = {};
      for (const [k, v] of Object.entries(d.cw || {})) {
        const c = +k;
        if (op.t === 'ins') { cw[c > p ? c + 1 : c] = v; if (c === p) cw[p + 1] = v; }
        else if (c !== p) cw[c > p ? c - 1 : c] = v;
      }
      d.cw = cw;
      d.dim = [d.dim[0], d.dim[1] + (op.t === 'ins' ? 1 : -1)];
      if (d.print && d.print.area) {
        const m = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(d.print.area);
        if (m) { const [a, b] = mapRange(colNum(m[1]), colNum(m[3]), o); d.print = { ...d.print, area: numCol(a) + m[2] + ':' + numCol(b) + m[4] }; }
      }
    }
  }

  // ---------------------------------------------------------------- ค่าตั้งต้น / ตรวจสอบ
  const txt = (d, ref) => { const c = d.cells[ref]; return c && c.v !== undefined && c.v !== null ? String(c.v) : ''; };
  function defaultCfg(tpl) {
    const out = {};
    for (const kind of Object.keys(KINDS)) {
      const K = KINDS[kind], d = tpl.sheets.find(s => s.name === K.prefix + ' 1'), L = layout(kind, K.base);
      out[kind] = { n: K.def.slice(), g: K.origG.slice(), names: [0, 1].map(g => L.items[g].slice(0, K.def[g]).map(c => txt(d, numCol(c) + '3'))) };
    }
    return out;
  }
  function normalize(tpl, cfg) {
    const def = defaultCfg(tpl), out = {};
    for (const kind of Object.keys(KINDS)) {
      const c = (cfg && cfg[kind]) || def[kind], K = KINDS[kind];
      const n = [0, 1].map(g => Math.max(1, Math.min(K.max[g], Math.floor(+(c.n && c.n[g]) || def[kind].n[g]))));
      out[kind] = { n, g: [0, 1].map(g => String((c.g && c.g[g]) || def[kind].g[g])), names: [0, 1].map(g => Array.from({ length: n[g] }, (_, i) => String((c.names && c.names[g] && c.names[g][i] !== undefined) ? c.names[g][i] : (i + 1)))) };
    }
    return out;
  }

  // ---------------------------------------------------------------- สร้างแม่แบบตามค่าตั้ง
  function apply(tpl, cfgIn) {
    const cfg = normalize(tpl, cfgIn), rel = relevantSheets(tpl), W = new Map();
    for (const s of tpl.sheets) if (rel.has(s.name)) W.set(s.name, { ...s, cells: { ...s.cells }, merges: (s.merges || []).slice(), cw: { ...s.cw }, print: s.print && { ...s.print }, dim: s.dim.slice() });
    for (const kind of Object.keys(KINDS)) {
      const K = KINDS[kind], c = cfg[kind];
      for (const g of [1, 0]) {
        const diff = c.n[g] - K.base[g]; let p = layout(kind, K.base).items[g].slice(-1)[0];
        if (diff > 0) for (let i = 0; i < diff; i++) { colOp(W, kind, { t: 'ins', p }); p++; }
        else for (let i = 0; i < -diff; i++) { colOp(W, kind, { t: 'del', p }); p--; }
      }
      const L = layout(kind, c.n);
      for (const name of sheetsOf(kind)) {
        const d = W.get(name);
        [0, 1].forEach(g => {
          L.items[g].forEach((col, i) => {
            const ref = numCol(col) + '3', cur = d.cells[ref] || {}, nm = c.names[g][i];
            const o = { ...cur, v: /^\d+$/.test(nm) ? +nm : nm }; delete o.f; delete o.c; d.cells[ref] = o;
          });
          const gh = numCol(L.items[g][0]) + '2';
          if (d.cells[gh]) d.cells[gh] = { ...d.cells[gh], v: c.g[g] };
        });
        const t1 = d.cells.B1;   // ชื่อกลุ่มที่อยู่ในข้อความหัวเรื่องและเกณฑ์
        if (t1 && t1.f) { let f = t1.f; [0, 1].forEach(g => { if (c.g[g] !== K.origG[g]) f = f.split(K.origG[g]).join(c.g[g]); }); d.cells.B1 = { ...t1, f }; delete d.cells.B1.c; }
        const t52 = d.cells.B52;
        if (t52 && t52.v) { let v = String(t52.v); [0, 1].forEach(g => { if (c.g[g] !== K.origG[g]) v = v.split(K.origG[g]).join(c.g[g]); }); d.cells.B52 = { ...t52, v }; }
      }
      // ปพ.6: ดัชนีคอลัมน์ใน VLOOKUP + ป้ายชื่อกลุ่ม · ปก ปพ.5: ป้ายชื่อกลุ่ม
      const delta = [c.n[0] - K.base[0], (c.n[0] - K.base[0]) + (c.n[1] - K.base[1])];
      K.p6rows.forEach((row, g) => {
        for (const sn of ['ปพ.6 เทอม 1', 'ปพ.6 เทอม 2', 'ปพ.6 รวม']) {
          const d = W.get(sn); if (!d) continue;
          const ref = 'D' + row, cell = d.cells[ref];
          if (cell && cell.f) d.cells[ref] = { ...cell, f: cell.f.replace(/,(\d+)\)="\/"/g, (m, n) => ',' + (+n + delta[g]) + ')="/"') };
          const lab = d.cells['A' + row];
          if (lab && c.g[g] !== K.origG[g]) d.cells['A' + row] = { ...lab, v: 'ผลการประเมิน' + c.g[g] };
        }
      });
    }
    // ชื่อกลุ่มบนปก ปพ.5 (รายชั้น/รายวิชา)
    const cover = { 'AB39': ['trait', 0], 'AB43': ['value', 0], 'BF43': ['value', 1] };
    for (const s of tpl.sheets) {
      if (!/^ปก ปพ\.5 /.test(s.name) || !W.has(s.name)) continue;
      const d = W.get(s.name);
      for (const [ref, [kind, g]] of Object.entries(cover)) {
        const cell = d.cells[ref]; if (!cell || cell.v === undefined || cfg[kind].g[g] === KINDS[kind].origG[g]) continue;
        d.cells[ref] = { ...cell, v: cfg[kind].g[g] };
      }
    }
    return { ...tpl, sheets: tpl.sheets.map(s => W.get(s.name) || s) };
  }

  // ---------------------------------------------------------------- ย้ายข้อมูลที่กรอกไว้เมื่อโครงสร้างเปลี่ยน
  // edits: { ชื่อชีท: { 'C5': ค่า } } ตามผังเดิม (oldN) → ผังใหม่ (newN) · del = { g, i } ลบหัวข้อลำดับ i ของกลุ่ม g
  function remapEdits(edits, kind, oldN, newN, del) {
    const Lo = layout(kind, oldN), Ln = layout(kind, newN), map = new Map();
    for (let g = 0; g < 2; g++) for (let i = 0; i < oldN[g]; i++) {
      let j = i; if (del && del.g === g) { if (i === del.i) continue; if (i > del.i) j = i - 1; }
      if (j < newN[g]) map.set(Lo.items[g][i], Ln.items[g][j]);
    }
    const itemCols = new Set(Lo.items[0].concat(Lo.items[1]));
    for (const name of sheetsOf(kind)) {
      const m = edits[name]; if (!m) continue;
      const out = {};
      for (const [ref, v] of Object.entries(m)) {
        const a = parseRef(ref);
        if (a.r < 5) continue;
        if (itemCols.has(a.c)) { const nc = map.get(a.c); if (nc) out[numCol(nc) + a.r] = v; } else out[ref] = v;
      }
      edits[name] = out;
    }
    return edits;
  }

  root.PPAssess = { KINDS, layout, apply, normalize, defaultCfg, remapEdits, sheetsOf, numCol, colNum };
  if (typeof module !== 'undefined' && module.exports) module.exports = root.PPAssess;
})(typeof window !== 'undefined' ? window : globalThis);
