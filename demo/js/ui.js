/* ui.js — หน้าใช้งานแบบเว็บ (ฟอร์ม/การ์ด/ตารางกรอกข้อมูล) ที่อ่านและเขียนข้อมูลชุดเดียวกับชีท Excel
 * ทุกค่าที่กรอกส่งผ่าน PPApp.setCell → เก็บลงเอกสาร + บันทึกอัตโนมัติ  ·  ผลคำนวณมาจากสูตรในชีทเดิม (engine.js)
 * ชีทแบบ Excel ยังเปิดดู/แก้/พิมพ์ได้จากเมนู "เอกสาร" */
(function (root) {
  'use strict';
  const P = root.PPEngine, R = root.PPRender;
  const $ = (s, e) => (e || document).querySelector(s);
  const $$ = (s, e) => Array.from((e || document).querySelectorAll(s));
  const esc = R.esc;
  const A = () => root.PPApp;
  const WB = () => A().S.wb;
  const C = P.colToNum, L = P.numToCol;
  const MAX_STU = 45, MAX_SUBJ = 20;
  const U = { tab: 'home', dsub: 'school', subj: 1, term: 1, akind: 'trait', aterm: 1, p6term: 1, p6no: 1, docq: '' };

  // ---------------------------------------------------------------- เข้าถึงข้อมูลในเวิร์กบุ๊ก
  function get(n, r, c) {
    const wb = WB(), sh = wb.sheet(n); if (!sh) return null;
    let v = wb.get(sh, r, typeof c === 'string' ? C(c) : c); if (v && v.isRange) v = wb.scalar(v);
    return v === undefined || v === null || P.isErr(v) ? null : v;
  }
  const txt = (n, r, c) => { const v = get(n, r, c); return v === null ? '' : String(v); };
  const num = (n, r, c) => { const v = get(n, r, c); return typeof v === 'number' ? v : null; };
  const setv = (n, r, c, v) => A().setCell(n, r, typeof c === 'string' ? C(c) : c, v);
  const fmt2 = (v) => typeof v === 'number' ? (Number.isInteger(v) ? String(v) : String(Math.round(v * 100) / 100)) : (v === null ? '' : String(v));

  // นักเรียน: แถวของชีทข้อมูลนักเรียนที่มีชื่อ (แถว 3..)
  const ST = 'ข้อมูลนักเรียน', SUBJ = 'รายวิชา', INFO = 'ข้อมูลพื้นฐาน';
  function nStudents() { let n = 0; for (let i = 0; i < MAX_STU; i++) if (txt(ST, 3 + i, 9).trim() || txt(ST, 3 + i, 10).trim()) n = i + 1; return n; }
  const stuName = (i) => { const p = (c) => txt(ST, 3 + i, c); return (p(8) + p(9) + ' ' + p(10)).trim(); };
  const subjCount = () => { let n = 0; for (let i = 0; i < MAX_SUBJ; i++) if (txt(SUBJ, 6 + i, 3).trim()) n = i + 1; return n; };
  const subjects = () => Array.from({ length: MAX_SUBJ }, (_, i) => ({ slot: i + 1, row: 6 + i, name: txt(SUBJ, 6 + i, 3) })).filter(s => s.name.trim());

  // ---------------------------------------------------------------- ไอคอนและเมนู
  const IC = {
    home: '<path d="M3 11l9-8 9 8M5 10v10h14V10"/>', data: '<rect x="5" y="3" width="14" height="18" rx="2"/><path d="M9 8h6M9 12h6M9 16h4"/>',
    scores: '<path d="M4 20h4L19 9l-4-4L4 16z"/><path d="M13 7l4 4"/>', assess: '<circle cx="12" cy="12" r="9"/><path d="M8 12l3 3 5-6"/>',
    report5: '<path d="M4 5a2 2 0 012-2h5v18H6a2 2 0 01-2-2zM11 3h7a2 2 0 012 2v14a2 2 0 01-2 2h-7"/>', report: '<path d="M6 3h8l4 4v14H6z"/><path d="M14 3v4h4M9 12h6M9 16h6"/>',
    docs: '<path d="M7 9V3h10v6M7 17H5a2 2 0 01-2-2v-4a2 2 0 012-2h14a2 2 0 012 2v4a2 2 0 01-2 2h-2"/><rect x="7" y="14" width="10" height="7"/>',
    manual: '<circle cx="12" cy="12" r="9"/><path d="M9.5 9.5a2.5 2.5 0 114 2c-.9.6-1.5 1-1.5 2M12 17h.01"/>',
  };
  const TABS = [['home', 'หน้าแรก', 'home', 't-home'], ['data', 'ข้อมูล', 'data', 't-data'], ['scores', 'คะแนน', 'scores', 't-scores'], ['assess', 'ประเมิน', 'assess', 't-assess'],
    ['report5', 'รายงาน ปพ.5', 'report5', 't-report5'], ['report6', 'รายงาน ปพ.6', 'report', 't-report'], ['docs', 'เอกสาร', 'docs', 't-docs'], ['manual', 'คู่มือ', 'manual', 't-manual']];
  const svg = (k) => `<svg viewBox="0 0 24 24" aria-hidden="true">${IC[k]}</svg>`;

  function mount() {
    $('#tabbar').innerHTML = TABS.map(t => `<button data-tab="${t[0]}" class="${t[3]}">${svg(t[2])}<span>${t[1]}</span></button>`).join('');
    $('#tabbar').onclick = (e) => { const b = e.target.closest('button'); if (b) go(b.dataset.tab); };
    const v = $('#view');
    v.oninput = onInput; v.onchange = onChange; v.onclick = onClick; v.onkeydown = onKey; v.onpaste = onPaste;
    go('home');
  }
  function go(tab, sub) {
    U.tab = tab; if (sub) { if (tab === 'data') U.dsub = sub; }
    $$('#tabbar button').forEach(b => b.classList.toggle('on', b.dataset.tab === tab));
    const v = $('#view'); v.scrollTop = 0; render();
  }
  function refresh() { const v = $('#view'), top = v.scrollTop; render(); v.scrollTop = top; }
  function render() {
    const v = $('#view'); const f = { home: vHome, data: vData, scores: vScores, assess: vAssess, report5: vReport5, report6: vReport6, docs: vDocs, manual: vManual }[U.tab];
    v.innerHTML = '<div class="page">' + f() + '</div>';
    updateRo(v);
  }
  const seg = (items, cur, act) => `<div class="seg" role="tablist">${items.map(([k, t]) => `<button role="tab" data-act="${act}" data-k="${k}" class="${String(k) === String(cur) ? 'on' : ''}">${esc(t)}</button>`).join('')}</div>`;
  const field = (label, inner, cls) => `<label class="fld ${cls || ''}"><span>${esc(label)}</span>${inner}</label>`;

  // ---------------------------------------------------------------- ช่องตาราง (กรอกแบบชีท)
  // data-t: ชนิดของค่า  num | text | date | sex | pass   data-gr/gc: ตำแหน่งในตาราง (ใช้เลื่อนช่องและวางจาก Excel)
  function gcell(sheet, r, c, t, gr, gc, o) {
    o = o || {};
    const v = get(sheet, r, c); let shown = v === null ? '' : String(v);
    if (t === 'date' && typeof v === 'number') shown = R.fmtValue(v, { n: 'm/d/yyyy' });
    const lock = A().S.view.locked.has(sheet + '!' + L(typeof c === 'string' ? C(c) : c) + r);
    return `<td class="gtd${o.cls ? ' ' + o.cls : ''}"><input class="gc${o.cls2 ? ' ' + o.cls2 : ''}" data-sh="${esc(sheet)}" data-r="${r}" data-c="${typeof c === 'string' ? C(c) : c}" data-t="${t}" data-gr="${gr}" data-gc="${gc}" value="${esc(shown)}"${lock ? ' disabled' : ''}${o.ph ? ` placeholder="${esc(o.ph)}"` : ''}${t === 'num' ? ' inputmode="decimal"' : ''}></td>`;
  }
  // ช่องผลลัพธ์ (อ่านอย่างเดียว) อัปเดตเองเมื่อมีการกรอก
  const rocell = (sheet, r, c, kind) => `<td class="ro" data-ro="${esc(sheet)}|${r}|${typeof c === 'string' ? C(c) : c}|${kind || ''}"></td>`;
  function roText(sheet, r, c, kind) {
    const v = get(sheet, r, c);
    if (v === null || v === '') return '';
    if (kind === 'lvl') return String(v);
    return typeof v === 'number' ? fmt2(v) : String(v);
  }
  function updateRo(scope) {
    $$('[data-ro]', scope || document).forEach(td => {
      const [sh, r, c, k] = td.dataset.ro.split('|'); const t = roText(sh, +r, +c, k);
      if (td.textContent !== t) td.textContent = t;
      if (k === 'grade') td.className = 'ro ' + (t === '' ? '' : gradeCls(t));
    });
    $$('[data-roh]', scope || document).forEach(el => { const [sh, r, c] = el.dataset.roh.split('|'); el.textContent = roText(sh, +r, +c); });
  }
  const gradeCls = (g) => { const n = parseFloat(g); return isNaN(n) ? '' : n >= 3 ? 'gok' : n >= 1 ? 'gmid' : 'gbad'; };

  function parseCell(t, text) {
    text = String(text).replace(/\r/g, '').trim();
    if (t === 'sex') { if (/^(ช|ชาย|ด\.?ช\.?|เด็กชาย|m|male)$/i.test(text)) return 'ช'; if (/^(ญ|หญิง|ด\.?ญ\.?|เด็กหญิง|f|female)$/i.test(text)) return 'ญ'; return text === '' ? null : text; }
    if (t === 'pass') { if (/^(ผ|ผ่าน|\/|✓|✔|1|p|pass)$/i.test(text)) return 'ผ'; if (/^(มผ|ไม่ผ่าน|x|×|✗|0|f|fail)$/i.test(text)) return 'มผ'; return text === '' ? null : text; }
    if (t === 'num') { if (text === '') return null; const n = Number(text.replace(/,/g, '')); return isNaN(n) ? text : n; }
    if (t === 'date') return text === '' ? null : A().parseInput(ST, 13, text);
    return text === '' ? null : text;
  }
  const rawOf = (v, t) => v === null ? '' : (t === 'date' && typeof v === 'number' ? R.fmtValue(v, { n: 'm/d/yyyy' }) : String(v));

  function commitInput(inp) {
    const t = inp.dataset.t, sh = inp.dataset.sh, r = +inp.dataset.r, c = +inp.dataset.c;
    let val = parseCell(t, inp.value);
    if (t === 'num' && typeof val === 'string') { A().toast('กรุณากรอกเป็นตัวเลข', 'warn'); inp.value = rawOf(get(sh, r, c), t); return false; }
    const old = get(sh, r, c);
    if (!(val === null && (old === null || old === '')) && val !== old) setv(sh, r, c, val);
    inp.value = rawOf(get(sh, r, c), t);
    if (t === 'pass') markPass(inp);
    return true;
  }
  const markPass = (inp) => { inp.classList.toggle('ok', inp.value === 'ผ'); inp.classList.toggle('no', inp.value === 'มผ'); };
  function onChange(e) {
    const inp = e.target.closest('input.gc'); if (inp) { commitInput(inp); updateRo($('#view')); return; }
    const x = e.target.closest('[data-f]'); if (x) fieldChange(x);
  }
  function onInput(e) { const s = e.target.closest('input[data-search]'); if (s) { U.docq = s.value; docFilter(); } }
  const gridAt = (tbl, gr, gc) => $(`input.gc[data-gr="${gr}"][data-gc="${gc}"]`, tbl);
  function gridMove(el, dr, dc) {
    const tbl = el.closest('table'); let gr = +el.dataset.gr, gc = +el.dataset.gc;
    for (let i = 0; i < 80; i++) {
      gr += dr; gc += dc;
      const nx = gridAt(tbl, gr, gc);
      if (nx && !nx.disabled) { nx.focus(); nx.select(); return; }
      if (!nx && !tbl.querySelector(`input.gc[data-gr="${gr}"]`) && dr) return;
      if (!nx && dc && !tbl.querySelector(`input.gc[data-gc="${gc}"]`)) return;
    }
  }
  function onKey(e) {
    const el = e.target.closest('input.gc'); if (!el) return;
    if (e.key === 'Enter' || e.key === 'ArrowDown') { e.preventDefault(); commitInput(el); updateRo($('#view')); gridMove(el, 1, 0); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); commitInput(el); updateRo($('#view')); gridMove(el, -1, 0); }
    else if (e.key === 'Tab') { e.preventDefault(); commitInput(el); updateRo($('#view')); gridMove(el, 0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'ArrowRight' && el.selectionStart === el.value.length) { e.preventDefault(); commitInput(el); gridMove(el, 0, 1); }
    else if (e.key === 'ArrowLeft' && el.selectionEnd === 0) { e.preventDefault(); commitInput(el); gridMove(el, 0, -1); }
    else if (el.dataset.t === 'pass' && e.key === ' ') { e.preventDefault(); el.value = el.value === 'ผ' ? 'มผ' : el.value === 'มผ' ? '' : 'ผ'; commitInput(el); updateRo($('#view')); }
  }
  // วางหลายช่องจาก Excel (แถว/คอลัมน์) เริ่มที่ช่องปัจจุบัน
  function onPaste(e) {
    const el = e.target.closest('input.gc'); if (!el) return;
    const text = (e.clipboardData || window.clipboardData).getData('text');
    if (!/[\t\n]/.test(text.replace(/\n$/, ''))) return;
    e.preventDefault();
    const tbl = el.closest('table'), rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
    let n = 0;
    rows.forEach((row, i) => row.forEach((cell, j) => {
      const t = gridAt(tbl, +el.dataset.gr + i, +el.dataset.gc + j); if (!t || t.disabled) return;
      t.value = cell; if (commitInput(t)) n++;
    }));
    updateRo($('#view')); A().toast('วางข้อมูล ' + n + ' ช่อง', 'ok');
  }

  // ---------------------------------------------------------------- การกดปุ่มต่าง ๆ
  function onClick(e) {
    const b = e.target.closest('[data-act]'); if (!b) return;
    const act = b.dataset.act, k = b.dataset.k;
    const set = (o) => { Object.assign(U, o); refresh(); };
    switch (act) {
      case 'go': go(b.dataset.tab, b.dataset.sub); break;
      case 'dsub': set({ dsub: k }); break;
      case 'term': set({ term: isNaN(+k) ? k : +k }); break;
      case 'akind': set({ akind: k }); break;
      case 'aterm': set({ aterm: isNaN(+k) ? k : +k }); break;
      case 'p6term': set({ p6term: isNaN(+k) ? k : +k }); break;
      case 'subj-step': { const list = subjects(); const i = list.findIndex(s => s.slot === U.subj); const n = list[i + (+k)]; if (n) set({ subj: n.slot }); break; }
      case 'doc': A().openSheet(b.dataset.name); break;
      case 'print-doc': A().printSheets([b.dataset.name]); break;
      case 'cyc': cycle(b); break;
      case 'setall': setAllLevels(+k); break;
      case 'act': toggleActivity(b); break;
      case 'ind-add': indAdd(1); break;
      case 'ind-del': indAdd(-1); break;
      case 'work-add': workAdd(1); break;
      case 'work-del': workAdd(-1); break;
      case 'addrows': set({ [b.dataset.key]: (U[b.dataset.key] || 0) + 5 }); break;
      case 'p6step': { const n = nStudents(); if (n) set({ p6no: ((U.p6no - 1 + (+k) + n) % n) + 1 }); break; }
      case 'p6print': printStudent(); break;
      case 'p6all': A().printAllStudents(p6Sheet()); break;
      case 'backup': A().exportBackup(); break;
      case 'restore': A().pickBackup(); break;
      case 'rooms': A().manageRooms(); break;
      case 'clearall': clearScores(); break;
      case 'toc': { const d = $('#man-' + k); if (d) { d.open = true; d.scrollIntoView({ behavior: 'smooth', block: 'start' }); } break; }
    }
  }
  function fieldChange(x) {
    const [sh, r, c, t] = x.dataset.f.split('|');
    if (x.dataset.f === 'subj') { U.subj = +x.value; refresh(); return; }
    if (x.dataset.f === 'p6no') { U.p6no = +x.value; refresh(); return; }
    let val = t === 'num' ? parseCell('num', x.value) : (String(x.value).trim() === '' ? null : String(x.value).trim());
    if (t === 'date') val = parseCell('date', x.value);
    if (t === 'num' && typeof val === 'string') { A().toast('กรุณากรอกเป็นตัวเลข', 'warn'); return; }
    setv(sh, +r, +c, val);
    if (t === 'date' && val !== null) x.value = rawOf(get(sh, +r, +c), 'date');
    updateRo($('#view'));
  }
  const bindF = (sh, r, c, t) => `data-f="${esc(sh)}|${r}|${typeof c === 'string' ? C(c) : c}|${t || 'text'}"`;
  const ph = (v) => /^กรอก/.test(String(v || ''));
  function formInput(sh, r, c, label, o) {
    o = o || {}; const v = get(sh, r, c); const shown = v === null || ph(v) ? '' : (o.t === 'date' ? rawOf(v, 'date') : String(v));
    const locked = A().S.view.locked.has(sh + '!' + L(typeof c === 'string' ? C(c) : c) + r);
    return field(label, `<input ${bindF(sh, r, c, o.t)} value="${esc(shown)}" placeholder="${esc(ph(v) ? v : (o.ph || ''))}"${locked ? ' disabled' : ''}${o.t === 'num' ? ' inputmode="numeric"' : ''}>`, o.cls);
  }

  // ---------------------------------------------------------------- หน้าแรก
  function scoreProgress() {
    const n = nStudents(), subs = subjects(); if (!n || !subs.length) return 0;
    let done = 0, all = 0;
    for (const s of subs) for (const ex of ['DA', 'EE']) for (let i = 0; i < n; i++) { all++; if (num('วิชาที่ ' + s.slot, 7 + i, ex) !== null) done++; }
    return Math.round(done * 100 / all);
  }
  function assessProgress() {
    const n = nStudents(); if (!n) return 0; let done = 0, all = 0;
    for (const t of [1, 2]) for (let i = 0; i < n; i++) for (let c = 3; c <= 10; c++) { all++; if (num('คุณลักษณะ ' + t, 5 + i, c) !== null) done++; }
    return Math.round(done * 100 / all);
  }
  function vHome() {
    const S = A().S, n = nStudents(), ns = subjCount(), sp = scoreProgress(), ap = assessProgress();
    const info = !ph(get(INFO, 9, 2)) && get(INFO, 9, 2) !== null && !ph(get(INFO, 13, 2)) && get(INFO, 13, 2) !== null;
    const step = (no, done, title, sub, tab, subk, prog) => `<button class="step${done ? ' done' : ''}" data-act="go" data-tab="${tab}"${subk ? ` data-sub="${subk}"` : ''}><span class="no">${done ? '✓' : no}</span><span style="flex:1;min-width:0"><b>${title}</b><small>${sub}</small>${prog !== undefined ? `<div class="prog"><i style="width:${prog}%"></i></div>` : ''}</span><span class="go">›</span></button>`;
    return `<div class="card hello"><div><h2>${esc(S.school.name)} 🔒</h2><p>${esc(txt(SUBJ, 2, 2))} · ปีการศึกษา ${esc(txt(SUBJ, 4, 2))}</p></div></div>
      <div class="stat"><div><b>${n}</b><span>นักเรียน</span></div><div><b>${ns}</b><span>รายวิชา</span></div><div><b>${sp}%</b><span>คะแนนกรอกแล้ว</span></div></div>
      <div class="steps">
        ${step(1, info, 'ตั้งค่าโรงเรียน', info ? 'กรอกชื่อผู้อำนวยการและครูประจำชั้นแล้ว' : 'กรอกชื่อ ผอ. ครูประจำชั้น และปีการศึกษา', 'data', 'school')}
        ${step(2, n > 0, 'รายชื่อนักเรียน', n ? n + ' คน' : 'กรอกหรือวางรายชื่อจาก Excel', 'data', 'students')}
        ${step(3, ns > 0, 'รายวิชา', ns ? ns + ' วิชา' : 'เพิ่มรหัสวิชา ชื่อวิชา เวลาเรียน และครูผู้สอน', 'data', 'subjects')}
        ${step(4, sp >= 100, 'บันทึกคะแนน', 'เลือกวิชาแล้วกรอกตัวชี้วัด คะแนนระหว่างเรียน และคะแนนสอบ ทีละภาคเรียน', 'scores', '', sp)}
        ${step(5, ap >= 100, 'ประเมินคุณลักษณะ ค่านิยม และกิจกรรม', 'แตะช่องเพื่อเลือกระดับ / ผ่าน / ไม่ผ่าน', 'assess', '', ap)}
        ${step(6, false, 'รายงาน ปพ.5 และ ปพ.6', 'ดูสรุปผลรายห้อง รายบุคคล และพิมพ์', 'report6')}
      </div>`;
  }

  // ---------------------------------------------------------------- ข้อมูล
  function vData() {
    const sub = U.dsub;
    return `<div class="bar">${seg([['school', 'โรงเรียน'], ['students', 'นักเรียน'], ['subjects', 'รายวิชา'], ['body', 'น้ำหนัก-ส่วนสูง'], ['time', 'เวลาเรียน'], ['backup', 'สำรองข้อมูล']], sub, 'dsub')}</div>` +
      ({ school: vSchool, students: vStudents, subjects: vSubjects, body: vBody, time: vTime, backup: vBackup }[sub] || vSchool)();
  }
  function vSchool() {
    const S = A().S;
    return `<div class="card"><h3>ข้อมูลโรงเรียนและชั้นเรียน</h3><div class="grid2">
      ${field('ชื่อโรงเรียน (ล็อกตามรหัสสิทธิ์)', `<input value="${esc(S.school.name)}" disabled>`)}
      ${field('ชั้น (ตามห้องที่เปิดอยู่)', `<input value="${esc(txt(SUBJ, 2, 2))}" disabled>`)}
      ${formInput(SUBJ, 4, 2, 'ปีการศึกษา', { t: 'num', ph: 'เช่น 2566' })}
      ${formInput(SUBJ, 3, 2, 'ภาคเรียนที่ (ใช้ในส่วนหัวเอกสาร)', { t: 'num', ph: '1 หรือ 2' })}
      ${formInput(INFO, 4, 2, 'สังกัดระดับกรม')}${formInput(INFO, 8, 2, 'สังกัดระดับจังหวัด / เขตพื้นที่การศึกษา')}
      ${formInput(INFO, 5, 2, 'ตำบล / แขวง')}${formInput(INFO, 6, 2, 'อำเภอ / เขต')}${formInput(INFO, 7, 2, 'จังหวัด')}
    </div></div>
    <div class="card"><h3>ผู้ลงนามในเอกสาร</h3><div class="grid2">
      ${formInput(INFO, 9, 2, 'ชื่อ-สกุล ผู้อำนวยการ')}${formInput(INFO, 10, 2, 'ตำแหน่งผู้อำนวยการ / รักษาการ')}
      ${formInput(INFO, 11, 2, 'ชื่อ-สกุล นายทะเบียน')}${formInput(INFO, 12, 2, 'ชื่อ-สกุล หัวหน้างานวิชาการ')}
      ${formInput(INFO, 13, 2, 'ชื่อ-สกุล ครูประจำชั้น / ครูที่ปรึกษา')}
    </div><p class="hint">ข้อมูลนี้เก็บแยกตามห้อง ถ้าโรงเรียนมีหลายห้องต้องกรอกในแต่ละห้อง</p></div>`;
  }
  const PREFIX = ['เด็กชาย', 'เด็กหญิง', 'นาย', 'นางสาว', 'นาง'];
  function vStudents() {
    const n = nStudents(), rows = Math.min(MAX_STU, Math.max(n + 3, 10, U.stuRows || 0));
    let body = '';
    for (let i = 0; i < rows; i++) {
      const r = 3 + i;
      body += `<tr><td class="no">${i + 1}</td>${gcell(ST, r, 6, 'text', i, 0, { ph: 'รหัส' })}${gcell(ST, r, 3, 'text', i, 1, { ph: '13 หลัก' })}${gcell(ST, r, 7, 'sex', i, 2, { cls: 'sex' })}${gcell(ST, r, 8, 'text', i, 3)}${gcell(ST, r, 9, 'text', i, 4)}${gcell(ST, r, 10, 'text', i, 5)}${gcell(ST, r, 13, 'date', i, 6, { ph: 'วว/ดด/ปปปป' })}</tr>`;
    }
    return `<datalist id="prefixes">${PREFIX.map(p => `<option value="${p}">`).join('')}</datalist>
      <div class="card soft"><b>รายชื่อนักเรียน ${n} คน</b> (สูงสุด ${MAX_STU} คน)<p class="hint" style="margin:.3rem 0 0">คลิกช่องแล้วพิมพ์ · ลูกศร/Enter เลื่อนช่อง · คัดลอกตารางจาก Excel มาวาง (Ctrl+V) ที่ช่องแรกได้ทั้งตาราง · เพศกรอก ช / ญ · วันเกิดกรอก วว/ดด/ปปปป (พ.ศ.)</p></div>
      <div class="tw fit"><table class="t sheet fit"><thead><tr><th class="no">ที่</th><th>รหัสนักเรียน</th><th>เลขประจำตัวประชาชน</th><th class="sex">เพศ</th><th>คำนำหน้า</th><th>ชื่อ</th><th>นามสกุล</th><th>วันเกิด</th></tr></thead><tbody>${body}</tbody></table></div>
      ${rows < MAX_STU ? '<p><button class="btn" data-act="addrows" data-key="stuRows">＋ เพิ่ม 5 แถว</button></p>' : ''}`.replace(/(<input class="gc"[^>]*data-gc="3")/g, '$1 list="prefixes"');
  }
  const GROUPS8 = ['ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์และเทคโนโลยี', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 'สุขศึกษาและพลศึกษา', 'ศิลปะ', 'การงานอาชีพและเทคโนโลยี', 'ภาษาต่างประเทศ', 'รายวิชาเพิ่มเติม'];
  function vSubjects() {
    const n = subjCount(), rows = Math.min(MAX_SUBJ, Math.max(n + 2, 8, U.subjRows || 0));
    let body = '';
    for (let i = 0; i < rows; i++) {
      const r = 6 + i;
      body += `<tr><td class="no">${i + 1}</td>${gcell(SUBJ, r, 2, 'text', i, 0, { ph: 'เช่น ท11101' })}${gcell(SUBJ, r, 3, 'text', i, 1, { cls: 'wide', ph: 'ชื่อวิชา' })}${gcell(SUBJ, r, 4, 'num', i, 2)}${gcell(SUBJ, r, 5, 'text', i, 3, { cls: 'wide' })}${gcell(SUBJ, r, 6, 'text', i, 4, { cls: 'wide' })}</tr>`;
    }
    return `<datalist id="groups">${GROUPS8.map(p => `<option value="${p}">`).join('')}</datalist>
      <div class="card soft"><b>รายวิชา ${n} วิชา</b> (สูงสุด ${MAX_SUBJ} วิชา)<p class="hint" style="margin:.3rem 0 0">ชื่อวิชาที่กรอกจะไปปรากฏในเมนู “คะแนน” และในเอกสาร ปพ.5 / ปพ.6 ทุกฉบับ เวลาเรียนคือชั่วโมงต่อปี</p></div>
      <div class="tw fit"><table class="t sheet fit"><thead><tr><th class="no">ที่</th><th>รหัสวิชา</th><th>รายวิชา</th><th>เวลาเรียน (ชม./ปี)</th><th>กลุ่มสาระการเรียนรู้</th><th>ครูประจำวิชา</th></tr></thead><tbody>${body}</tbody></table></div>
      ${rows < MAX_SUBJ ? '<p><button class="btn" data-act="addrows" data-key="subjRows">＋ เพิ่ม 2 แถว</button></p>' : ''}`.replace(/(<input class="gc"[^>]*data-gc="3")/g, '$1 list="groups"');
  }
  function vBody() {
    const n = nStudents(); if (!n) return '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่แท็บ “นักเรียน” ก่อน</div>';
    const BS = 'น้ำหนัก-ส่วนสูง'; let body = '';
    for (let i = 0; i < n; i++) body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>${[3, 4, 5, 6, 7, 8, 9, 10].map((c, j) => gcell(BS, 7 + i, c, 'num', i, j)).join('')}</tr>`;
    const dts = [3, 5, 7, 9].map(c => `<th colspan="2">ครั้งที่ ${(c - 1) / 2}<br><input class="in" style="min-height:30px;width:9rem" ${bindF(BS, 6, c, 'date')} value="${esc(rawOf(get(BS, 6, c), 'date'))}" placeholder="วว/ดด/ปปปป"></th>`).join('');
    return `<div class="card soft"><b>น้ำหนัก (กก.) · ส่วนสูง (ซม.) 4 ครั้งต่อปี</b><p class="hint" style="margin:.3rem 0 0">กรอกวันที่ชั่งวัดไว้ที่หัวตาราง แล้วกรอกค่าของนักเรียนแต่ละคน</p></div>
      <div class="tw"><table class="t sheet"><thead><tr><th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>${dts}</tr><tr><th class="no"></th><th class="stk"></th>${'<th>น้ำหนัก</th><th>ส่วนสูง</th>'.repeat(4)}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  function vTime() {
    return `<div class="card"><h3>ปฏิทินและเวลาเรียน</h3>
      <p class="hint">ตั้งวันเปิด-ปิดเรียนและวันหยุดในชีทปฏิทินก่อน แล้วจึงบันทึกเวลาเรียนรายวันของนักเรียน (กรอก 1 = มา, ป = ป่วย, ล = ลา, ข = ขาด) ทั้งสองชีทกรอกแบบตาราง Excel</p>
      <div class="row"><button class="btn primary" data-act="doc" data-name="ปฏิทิน">เปิดชีทปฏิทิน</button><button class="btn primary" data-act="doc" data-name="เวลาเรียน">เปิดชีทเวลาเรียน</button><button class="btn" data-act="doc" data-name="ข้อมูลเพิ่มเติม">ข้อมูลนักเรียนเพิ่มเติม (บิดามารดา ผู้ปกครอง)</button></div></div>`;
  }
  function vBackup() {
    return `<div class="card"><h3>สำรองข้อมูล</h3><p class="hint">ดาวน์โหลดข้อมูลทุกห้องของโรงเรียนเป็นไฟล์ .json เก็บไว้ในเครื่อง และนำเข้ากลับได้เมื่อต้องการ</p>
      <div class="row"><button class="btn primary" data-act="backup">สำรองข้อมูล (ดาวน์โหลด)</button><button class="btn" data-act="restore">นำเข้าไฟล์สำรอง</button><button class="btn" data-act="rooms">จัดการห้องเรียน</button></div></div>`;
  }

  // ---------------------------------------------------------------- คะแนนรายวิชา
  // ผังชีท "วิชาที่ N": ตัวชี้วัด ผ/มผ · คะแนนระหว่างเรียน (แถว 5 = ชื่อ แถว 6 = คะแนนเต็ม) · สอบปลายภาค · ผลลัพธ์
  const SC = {
    1: { ind: [C('C'), C('AK')], work: [C('CA'), C('CY')], exam: C('DA'), btw: C('CZ'), total: C('DB'), pct: C('DC'), grade: C('DD') },
    2: { ind: [C('AO'), C('BW')], work: [C('DE'), C('EB')], exam: C('EE'), btw: C('ED'), total: C('EF'), pct: C('EG'), grade: C('EH') },
  };
  const IND_ROW = 7, SC_NAME = 5, SC_MAX = 6;
  function scCfg() {
    const T = SC[U.term]; if (!T) return null;
    const sh = 'วิชาที่ ' + U.subj;
    let ni = 0, nw = 0;
    for (let c = T.ind[0]; c <= T.ind[1]; c++) if (txt(sh, SC_NAME, c).trim()) ni = c - T.ind[0] + 1;
    for (let c = T.work[0]; c <= T.work[1]; c++) if ((num(sh, SC_MAX, c) || 0) > 0) nw = c - T.work[0] + 1;
    return { sh, T, ni, nw };
  }
  function indAdd(d) {
    const k = scCfg(); if (!k) return; const max = k.T.ind[1] - k.T.ind[0] + 1;
    if (d > 0 && k.ni < max) setv(k.sh, SC_NAME, k.T.ind[0] + k.ni, 'ตัวชี้วัดที่ ' + (k.ni + 1));
    else if (d < 0 && k.ni > 0) setv(k.sh, SC_NAME, k.T.ind[0] + k.ni - 1, null);
    refresh();
  }
  function workAdd(d) {
    const k = scCfg(); if (!k) return; const max = k.T.work[1] - k.T.work[0] + 1;
    if (d > 0 && k.nw < max) { const c = k.T.work[0] + k.nw; setv(k.sh, SC_NAME, c, 'งานที่ ' + (k.nw + 1)); setv(k.sh, SC_MAX, c, 10); }
    else if (d < 0 && k.nw > 0) { const c = k.T.work[0] + k.nw - 1; setv(k.sh, SC_NAME, c, null); setv(k.sh, SC_MAX, c, null); }
    refresh();
  }
  async function clearScores() {
    const k = scCfg(); if (!k) return;
    const yes = await A().modal('ล้างคะแนนภาคเรียนนี้', `<p>ล้างตัวชี้วัด คะแนนระหว่างเรียน และคะแนนสอบของ <b>${esc(txt(SUBJ, 5 + U.subj, 3))}</b> ภาคเรียนที่ ${U.term} ของนักเรียนทุกคน ย้อนกลับไม่ได้</p>`, [{ label: 'ยกเลิก', value: false }, { label: 'ล้างคะแนน', value: true, primary: true }]);
    if (!yes) return;
    for (let i = 0; i < MAX_STU; i++) {
      const r = IND_ROW + i;
      for (let c = k.T.ind[0]; c <= k.T.ind[1]; c++) if (get(k.sh, r, c) !== null) setv(k.sh, r, c, null);
      for (let c = k.T.work[0]; c <= k.T.work[1]; c++) if (get(k.sh, r, c) !== null) setv(k.sh, r, c, null);
      if (get(k.sh, r, k.T.exam) !== null) setv(k.sh, r, k.T.exam, null);
    }
    refresh();
  }
  function vScores() {
    const list = subjects();
    if (!list.length) return '<div class="card empty"><b>ยังไม่มีรายวิชา</b>เพิ่มรายวิชาที่แท็บ “ข้อมูล › รายวิชา” ก่อน<br><br><button class="btn primary" data-act="go" data-tab="data" data-sub="subjects">ไปเพิ่มรายวิชา</button></div>';
    if (!list.some(s => s.slot === U.subj)) U.subj = list[0].slot;
    const n = nStudents(), idx = list.findIndex(s => s.slot === U.subj);
    const head = `<div class="pick"><button class="btn" data-act="subj-step" data-k="-1" ${idx <= 0 ? 'disabled' : ''}>‹</button>
      <select data-f="subj" aria-label="เลือกวิชา">${list.map(s => `<option value="${s.slot}"${s.slot === U.subj ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}</select>
      <button class="btn" data-act="subj-step" data-k="1" ${idx >= list.length - 1 ? 'disabled' : ''}>›</button></div>
      <div class="bar" style="margin-top:10px">${seg([[1, 'เทอม 1'], [2, 'เทอม 2'], ['sum', 'สรุปทั้งปี']], U.term, 'term')}</div>`;
    if (!n) return head + '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่แท็บ “ข้อมูล › นักเรียน” ก่อน</div>';
    return head + (U.term === 'sum' ? vScoreYear(n) : vScoreTerm(n));
  }
  function vScoreTerm(n) {
    const k = scCfg(), { sh, T } = k, ratio = [num(sh, 3, 'DA'), num(sh, 3, 'DC')];
    const indIn = Array.from({ length: k.ni }, (_, j) => `<label class="fld"><span>ตัวชี้วัดที่ ${j + 1}</span><input ${bindF(sh, SC_NAME, T.ind[0] + j)} value="${esc(txt(sh, SC_NAME, T.ind[0] + j))}"></label>`).join('');
    const workIn = Array.from({ length: k.nw }, (_, j) => `<div class="indset"><b>งานที่ ${j + 1}</b>${field('ชื่อ', `<input ${bindF(sh, SC_NAME, T.work[0] + j)} value="${esc(txt(sh, SC_NAME, T.work[0] + j))}">`)}${field('คะแนนเต็ม', `<input inputmode="numeric" ${bindF(sh, SC_MAX, T.work[0] + j, 'num')} value="${esc(txt(sh, SC_MAX, T.work[0] + j))}">`)}</div>`).join('');
    const none = !k.ni && !k.nw;
    const cfg = `<details class="card soft more" ${none ? 'open' : ''}><summary><b>ตั้งค่าตัวชี้วัดและคะแนนเต็ม</b> <span class="hint">ใช้ร่วมกันทั้ง 2 ภาคเรียนของวิชานี้ ยกเว้นรายการตัวชี้วัด/งานที่แยกตามภาค</span></summary>
      <div class="maxbox" style="margin-top:8px">
        <label class="fld"><span>สัดส่วน ระหว่างเรียน</span><input inputmode="numeric" ${bindF(sh, 3, 'DA', 'num')} value="${esc(txt(sh, 3, 'DA'))}"></label>
        <label class="fld"><span>สัดส่วน ปลายภาค (= คะแนนสอบเต็ม)</span><input inputmode="numeric" ${bindF(sh, 3, 'DC', 'num')} value="${esc(txt(sh, 3, 'DC'))}"></label>
      </div>
      <p class="hint">ระหว่างเรียน : ปลายภาค = ${fmt2(ratio[0])} : ${fmt2(ratio[1])} (รวม ${fmt2((ratio[0] || 0) + (ratio[1] || 0))} คะแนน) คะแนนระหว่างเรียนทุกงานจะถูกปรับให้เต็มตามสัดส่วนระหว่างเรียนอัตโนมัติ</p>
      <h4 style="margin:.6rem 0 .3rem">ตัวชี้วัด (ผ/มผ) เทอม ${U.term}</h4><div class="maxbox">${indIn}</div>
      <div class="row" style="margin:6px 0"><button class="btn sm" data-act="ind-add">＋ เพิ่มตัวชี้วัด</button><button class="btn sm" data-act="ind-del" ${k.ni ? '' : 'disabled'}>－ ลดตัวชี้วัด</button></div>
      <h4 style="margin:.6rem 0 .3rem">งาน / คะแนนระหว่างเรียน เทอม ${U.term}</h4><div class="maxbox">${workIn}</div>
      <div class="row" style="margin:6px 0"><button class="btn sm" data-act="work-add">＋ เพิ่มงาน</button><button class="btn sm" data-act="work-del" ${k.nw ? '' : 'disabled'}>－ ลดงาน</button></div></details>`;
    if (none) return cfg + '<div class="card empty"><b>ยังไม่ได้กำหนดตัวชี้วัดและงาน</b>กดปุ่มด้านบนเพื่อเพิ่ม หรือ<br><br><button class="btn primary" data-act="work-add">เพิ่มงานแรก</button></div>';
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>', body = '';
    for (let j = 0; j < k.ni; j++) th += `<th>${esc(txt(sh, SC_NAME, T.ind[0] + j) || 'ตัวชี้วัด')}<br><small>ผ / มผ</small></th>`;
    for (let j = 0; j < k.nw; j++) th += `<th>${esc(txt(sh, SC_NAME, T.work[0] + j) || 'งาน')}<br><small>(${fmt2(num(sh, SC_MAX, T.work[0] + j))})</small></th>`;
    th += `<th>สอบ<br><small>(${esc(txt(sh, 3, 'DC'))})</small></th><th class="ro">ระหว่างเรียน</th><th class="ro">รวม</th><th class="ro">ร้อยละ</th><th class="ro">ระดับ</th>`;
    for (let i = 0; i < n; i++) {
      const r = IND_ROW + i; let gc = 0;
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`;
      for (let j = 0; j < k.ni; j++) body += gcell(sh, r, T.ind[0] + j, 'pass', i, gc++, { cls2: get(sh, r, T.ind[0] + j) === 'ผ' ? 'ok' : get(sh, r, T.ind[0] + j) === 'มผ' ? 'no' : '' });
      for (let j = 0; j < k.nw; j++) body += gcell(sh, r, T.work[0] + j, 'num', i, gc++);
      body += gcell(sh, r, T.exam, 'num', i, gc++) + rocell(sh, r, T.btw) + rocell(sh, r, T.total) + rocell(sh, r, T.pct) + rocell(sh, r, T.grade, 'grade') + '</tr>';
    }
    return cfg + `<p class="hint">คลิกช่องแล้วพิมพ์ · ตัวชี้วัดพิมพ์ ผ / มผ (หรือกด Space สลับ) · ลูกศร/Enter เลื่อนช่อง · วางคะแนนหลายช่องจาก Excel ได้ (Ctrl+V)</p>
      <div class="tw"><table class="t sheet"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>
      <p style="margin-top:10px"><button class="btn danger sm" data-act="clearall">ล้างคะแนนภาคเรียนนี้</button> <button class="btn sm" data-act="doc" data-name="${esc(sh)}">เปิดชีทเต็มรูปแบบ (Excel)</button></p>`;
  }
  function vScoreYear(n) {
    const sh = 'วิชาที่ ' + U.subj; let body = '';
    for (let i = 0; i < n; i++) {
      const r = IND_ROW + i;
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>${rocell(sh, r, SC[1].total)}${rocell(sh, r, SC[1].grade, 'grade')}${rocell(sh, r, SC[2].total)}${rocell(sh, r, SC[2].grade, 'grade')}${rocell(sh, r, 'EQ')}${rocell(sh, r, 'ER')}${rocell(sh, r, 'ES', 'grade')}</tr>`;
    }
    return `<p class="hint">สรุปผลทั้งปี คำนวณจากคะแนนภาคเรียนที่ 1 และ 2 (ดูอย่างเดียว)</p><div class="tw"><table class="t sheet"><thead><tr><th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th><th>เทอม 1 รวม</th><th>ระดับ</th><th>เทอม 2 รวม</th><th>ระดับ</th><th>รวมทั้งปี</th><th>ร้อยละ</th><th>ระดับผลการเรียน</th></tr></thead><tbody>${body}</tbody></table></div>`;
  }

  // ---------------------------------------------------------------- ประเมิน (คุณลักษณะ · ค่านิยม · กิจกรรม)
  const LV_TRAIT = ['ไม่ผ่าน', 'ผ่าน', 'ดี', 'ดีเยี่ยม'], LV_COMP = ['ปรับปรุง', 'พอใช้', 'ดี', 'ดีเยี่ยม'];
  const AS = {
    trait: { title: 'คุณลักษณะอันพึงประสงค์', base: 'คุณลักษณะ', groups: [{ name: 'คุณลักษณะอันพึงประสงค์ 8 ประการ', cols: [3, 10], lv: LV_TRAIT, hr: 3 }, { name: 'การอ่าน คิดวิเคราะห์ และเขียน', cols: [17, 21], lv: LV_TRAIT, hr: 3 }] },
    value: { title: 'ค่านิยม · สมรรถนะ', base: 'ค่านิยม', groups: [{ name: 'ค่านิยมหลัก 12 ประการ', cols: [3, 14], lv: LV_TRAIT, hr: 3 }, { name: 'สมรรถนะสำคัญของผู้เรียน', cols: [19, 23], lv: LV_COMP, hr: 3 }] },
    act: { title: 'กิจกรรมพัฒนาผู้เรียน', base: 'กิจกรรม' },
  };
  const assessSheet = () => AS[U.akind].base + ' ' + (U.aterm === 'sum' ? 'รวม' : U.aterm);
  function vAssess() {
    const n = nStudents();
    const head = `<div class="bar">${seg([['trait', 'คุณลักษณะ'], ['value', 'ค่านิยม · สมรรถนะ'], ['act', 'กิจกรรมพัฒนาผู้เรียน']], U.akind, 'akind')}${seg([[1, 'เทอม 1'], [2, 'เทอม 2'], ['sum', 'ทั้งปี']], U.aterm, 'aterm')}</div>`;
    if (!n) return head + '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่แท็บ “ข้อมูล › นักเรียน” ก่อน</div>';
    const edit = U.aterm !== 'sum', sh = assessSheet();
    if (U.akind === 'act') return head + vActivities(n, sh, edit);
    const K = AS[U.akind]; let h1 = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>', h2 = '<th class="no"></th><th class="stk"></th>', cols = [];
    for (const g of K.groups) {
      h1 += `<th colspan="${g.cols[1] - g.cols[0] + 1}">${esc(g.name)}</th>`;
      for (let c = g.cols[0]; c <= g.cols[1]; c++) { h2 += `<th>${esc(txt(sh, g.hr, c) || String(c - g.cols[0] + 1))}</th>`; cols.push({ c, lv: g.lv }); }
    }
    let body = '';
    for (let i = 0; i < n; i++) {
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`;
      for (const x of cols) body += lvCell(sh, 5 + i, x.c, x.lv, edit);
      body += '</tr>';
    }
    const tools = edit ? `<div class="bar"><span class="hint">ตั้งทั้งห้อง:</span><button class="btn sm" data-act="setall" data-k="1">ผ่าน</button><button class="btn sm" data-act="setall" data-k="2">ดี</button><button class="btn sm" data-act="setall" data-k="3">ดีเยี่ยม</button><button class="btn sm danger" data-act="setall" data-k="-1">ล้างทั้งหมด</button></div>` : '';
    return head + `<p class="hint">แตะช่องเพื่อเปลี่ยนระดับ (ว่าง → ผ่าน → ดี → ดีเยี่ยม → ไม่ผ่าน → ว่าง)${edit ? '' : ' · ทั้งปีคำนวณจากเทอม 1 และ 2 ดูอย่างเดียว'}</p>${tools}<div class="tw"><table class="t sheet"><thead><tr>${h1}</tr><tr>${h2}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  function lvCell(sh, r, c, labels, edit) {
    const v = get(sh, r, c), ok = typeof v === 'number' && v >= 0 && v <= 3;
    const cls = ok ? ' l' + v : '', label = ok ? labels[v] : '';
    if (!edit) return `<td>${ok ? `<span class="lvl l${v}">${labels[v]}</span>` : ''}</td>`;
    return `<td class="gtd"><button class="pc${cls}" data-act="cyc" data-sh="${esc(sh)}" data-r="${r}" data-c="${c}" data-lv="${labels === LV_COMP ? 'c' : 't'}">${label || '&nbsp;'}</button></td>`;
  }
  function cycle(b) {
    const sh = b.dataset.sh, r = +b.dataset.r, c = +b.dataset.c, labels = b.dataset.lv === 'c' ? LV_COMP : LV_TRAIT;
    const v = get(sh, r, c), next = typeof v !== 'number' ? 1 : v === 3 ? 0 : v === 0 ? null : v + 1;
    setv(sh, r, c, next);
    const ok = typeof next === 'number';
    b.className = 'pc' + (ok ? ' l' + next : ''); b.innerHTML = ok ? labels[next] : '&nbsp;';
  }
  function setAllLevels(level) {
    const n = nStudents(), sh = assessSheet(); if (!n || U.aterm === 'sum') return;
    if (U.akind === 'act') { for (let i = 0; i < n; i++) setActivity(sh, 5 + i, level === -1 ? null : 'pass'); refresh(); return; }
    for (const g of AS[U.akind].groups) for (let i = 0; i < n; i++) for (let c = g.cols[0]; c <= g.cols[1]; c++) setv(sh, 5 + i, c, level === -1 ? null : level);
    refresh();
  }
  // กิจกรรม: คู่คอลัมน์ [ไม่ผ่าน, ผ่าน] ของแต่ละกิจกรรม (ชุมนุมมีช่องชื่อชุมนุมอยู่ก่อน)
  const ACT = [[3, 4], [5, 6], [8, 9], [10, 11]];
  function actState(sh, r, [bad, ok]) { return txt(sh, r, ok).trim() ? 'pass' : txt(sh, r, bad).trim() ? 'fail' : null; }
  function setActivity(sh, r, state, pairs) {
    for (const [bad, ok] of (pairs || ACT)) { setv(sh, r, bad, state === 'fail' ? '/' : null); setv(sh, r, ok, state === 'pass' ? '/' : null); }
  }
  function toggleActivity(b) {
    const sh = b.dataset.sh, r = +b.dataset.r, p = ACT[+b.dataset.p], cur = actState(sh, r, p);
    const next = cur === null ? 'pass' : cur === 'pass' ? 'fail' : null;
    setActivity(sh, r, next, [p]);
    b.className = 'pc' + (next === 'pass' ? ' ok' : next === 'fail' ? ' no' : ''); b.textContent = next === 'pass' ? 'ผ่าน' : next === 'fail' ? 'ไม่ผ่าน' : ' ';
  }
  function vActivities(n, sh, edit) {
    const btn = (r, pi) => { const s = actState(sh, r, ACT[pi]); return edit ? `<td class="gtd"><button class="pc${s === 'pass' ? ' ok' : s === 'fail' ? ' no' : ''}" data-act="act" data-sh="${esc(sh)}" data-r="${r}" data-p="${pi}">${s === 'pass' ? 'ผ่าน' : s === 'fail' ? 'ไม่ผ่าน' : '&nbsp;'}</button></td>` : `<td>${s === 'pass' ? '<span class="pass ok">ผ่าน</span>' : s === 'fail' ? '<span class="pass no">ไม่ผ่าน</span>' : ''}</td>`; };
    let body = '';
    for (let i = 0; i < n; i++) {
      const r = 5 + i;
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>${btn(r, 0)}${btn(r, 1)}${edit ? gcell(sh, r, 7, 'text', i, 0, { ph: 'ชื่อชุมนุม' }) : `<td>${esc(txt(sh, r, 7))}</td>`}${btn(r, 2)}${btn(r, 3)}</tr>`;
    }
    const tools = edit ? `<div class="bar"><span class="hint">ตั้งทั้งห้อง:</span><button class="btn sm" data-act="setall" data-k="1">ผ่านทุกกิจกรรม</button><button class="btn sm danger" data-act="setall" data-k="-1">ล้างทั้งหมด</button></div>` : '';
    return `<p class="hint">แตะช่องเพื่อสลับ ว่าง → ผ่าน → ไม่ผ่าน → ว่าง</p>${tools}<div class="tw"><table class="t sheet"><thead><tr><th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th><th>${esc(txt(sh, 2, 3) || 'แนะแนว')}</th><th>${esc(txt(sh, 2, 5) || 'ลูกเสือ-เนตรนารี')}</th><th>ชื่อชุมนุม</th><th>${esc(txt(sh, 2, 7) || 'ชุมนุม')}</th><th>${esc((txt(sh, 2, 10) || 'กิจกรรมเพื่อสังคมฯ').replace(/\n/g, ' '))}</th></tr></thead><tbody>${body}</tbody></table></div>`;
  }

  // ---------------------------------------------------------------- รายงาน ปพ.5 (สรุปผลรายห้อง)
  const GRADES = [4, 3.5, 3, 2.5, 2, 1.5, 1, 0];
  const TCOL = { 1: { pct: 'DC', grade: 'DD', tot: 'DB' }, 2: { pct: 'EG', grade: 'EH', tot: 'EF' }, sum: { pct: 'ER', grade: 'ES', tot: 'EQ' } };
  const TERMSUF = { 1: '1', 2: '2', sum: 'รวม' };
  function vReport5() {
    const n = nStudents(), list = subjects(), tc = TCOL[U.term];
    const head = `<div class="bar">${seg([[1, 'เทอม 1'], [2, 'เทอม 2'], ['sum', 'ทั้งปี']], U.term, 'term')}</div>`;
    if (!n || !list.length) return head + '<div class="card empty"><b>ยังไม่มีข้อมูลสำหรับสรุป</b>กรอกรายชื่อนักเรียน รายวิชา และคะแนนก่อน</div>';
    let rows = '', sumAvg = 0, cntAvg = 0; const tot = {};
    for (const s of list) {
      const sh = 'วิชาที่ ' + s.slot, dist = {}; let sp = 0, np = 0;
      for (let i = 0; i < n; i++) {
        const p = num(sh, 7 + i, tc.pct), g = get(sh, 7 + i, tc.grade);
        if (p !== null) { sp += p; np++; }
        if (typeof g === 'number') { dist[g] = (dist[g] || 0) + 1; tot[g] = (tot[g] || 0) + 1; }
      }
      const avg = np ? sp / np : null; if (avg !== null) { sumAvg += avg; cntAvg++; }
      rows += `<tr><td class="stk">${esc(s.name)}</td><td class="num">${np}/${n}</td><td class="num">${avg === null ? '-' : fmt2(avg)}</td>${GRADES.map(g => `<td class="num">${dist[g] || '-'}</td>`).join('')}</tr>`;
    }
    const foot = `<tr><td class="stk">รวมทุกวิชา</td><td class="num"></td><td class="num">${cntAvg ? fmt2(sumAvg / cntAvg) : '-'}</td>${GRADES.map(g => `<td class="num">${tot[g] || '-'}</td>`).join('')}</tr>`;
    const suf = TERMSUF[U.term];
    const docs = [['ปก ปพ.5 รายชั้น ' + suf, 'ปก ปพ.5 (รายชั้น)'], ['ปก ปพ.5 รายวิชา ' + suf, 'ปก ปพ.5 (รายวิชา)'], ['คะแนนรายชั้น ' + suf, 'คะแนนรายชั้น'], ['เกรดรายชั้น ' + suf, 'เกรดรายชั้น'], ['คะแนนร้อยละ ' + suf, 'คะแนนร้อยละ'], ['เกรดร้อยละ ' + suf, 'เกรดร้อยละ']];
    return head + `<div class="p5stats"><div class="stat"><b>${n}</b><span>นักเรียน</span></div><div class="stat"><b>${list.length}</b><span>รายวิชา</span></div><div class="stat"><b>${cntAvg ? fmt2(sumAvg / cntAvg) : '-'}</b><span>ค่าเฉลี่ยร้อยละทุกวิชา</span></div></div>
      <div class="tw p5tw"><table class="t p5t"><thead><tr><th class="stk">รายวิชา</th><th>มีคะแนน</th><th>เฉลี่ย (%)</th>${GRADES.map(g => `<th>เกรด ${g}</th>`).join('')}</tr></thead><tbody>${rows}</tbody><tfoot>${foot}</tfoot></table></div>
      <div class="card" style="margin-top:12px"><h3>เอกสารสรุป (ดู / พิมพ์)</h3><div class="dlist">${docs.filter(d => WB().sheet(d[0])).map(d => `<div class="ditem"><button class="open" data-act="doc" data-name="${esc(d[0])}">${esc(d[1])}</button><button class="btn" data-act="print-doc" data-name="${esc(d[0])}">พิมพ์</button></div>`).join('')}</div></div>`;
  }

  // ---------------------------------------------------------------- รายงาน ปพ.6 (รายบุคคล)
  const p6Sheet = () => 'ปพ.6 ' + (U.p6term === 'sum' ? 'รวม' : 'เทอม ' + U.p6term);
  const NO = { r: 8, c: 11 };
  function printStudent() { selectStudent(); A().printSheets([p6Sheet()]); }
  function selectStudent() { WB().setValue(p6Sheet(), L(NO.c) + NO.r, U.p6no); }
  function vReport6() {
    const n = nStudents();
    const head = `<div class="bar">${seg([[1, 'เทอม 1'], [2, 'เทอม 2'], ['sum', 'ทั้งปี']], U.p6term, 'p6term')}</div>`;
    if (!n) return head + '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่แท็บ “ข้อมูล › นักเรียน” ก่อน</div>';
    if (U.p6no > n || U.p6no < 1) U.p6no = 1;
    const sh = p6Sheet(); selectStudent();
    const pick = `<div class="pick"><button class="btn" data-act="p6step" data-k="-1">‹</button><select data-f="p6no" aria-label="เลือกนักเรียน">${Array.from({ length: n }, (_, i) => `<option value="${i + 1}"${i + 1 === U.p6no ? ' selected' : ''}>${i + 1}. ${esc(stuName(i))}</option>`).join('')}</select><button class="btn" data-act="p6step" data-k="1">›</button></div>`;
    let rows = '';
    for (let r = 9; r <= 28; r++) {
      if (!txt(sh, r, 3)) continue;
      const g = get(sh, r, 9);
      rows += `<tr><td>${esc(txt(sh, r, 2))}</td><td class="stk" style="max-width:none">${esc(txt(sh, r, 3))}</td><td class="num">${esc(txt(sh, r, 4))}</td><td class="num">${esc(fmt2(get(sh, r, 5)))}</td><td class="num">${esc(fmt2(get(sh, r, 6)))}</td><td class="num">${esc(fmt2(get(sh, r, 7)))}</td><td class="num">${esc(fmt2(get(sh, r, 8)))}</td><td class="num"><span class="lvn ${typeof g === 'number' ? (g >= 3 ? 'l3' : g >= 2 ? 'l2' : g >= 1 ? 'l1' : 'l0') : ''}">${g === null ? '-' : esc(String(g))}</span></td></tr>`;
    }
    const chip = (v) => v === null || v === '' ? '<span class="lvl none">-</span>' : `<span class="pass ${v === 'ไม่ผ่าน' || v === 'ปรับปรุง' ? 'no' : 'ok'}">${esc(String(v))}</span>`;
    const act = [35, 36, 37, 38].map(r => `<div class="row sp"><span>${esc(txt(sh, r, 2))}${r === 37 && txt(sh, r, 3) ? ' (' + esc(txt(sh, r, 3)) + ')' : ''}</span>${chip(get(sh, r, 4))}</div>`).join('');
    const ass = [[39, 'คุณลักษณะอันพึงประสงค์'], [40, 'การอ่าน คิดวิเคราะห์ และเขียน'], [41, 'ค่านิยมหลัก 12 ประการ'], [42, 'สมรรถนะสำคัญของผู้เรียน']].map(([r, t]) => `<div class="row sp"><span>${t}</span>${chip(get(sh, r, 4))}</div>`).join('');
    return head + pick + `<div class="card" style="margin-top:12px"><div class="rhead"><div><h3>${esc(stuName(U.p6no - 1))}</h3><span class="hint">เลขที่ ${U.p6no} · รหัส ${esc(txt(ST, 3 + U.p6no - 1, 6))} · ${esc(txt(SUBJ, 2, 2))}</span></div>
        <div class="row"><button class="btn" data-act="doc" data-name="${esc(sh)}">ดูเอกสาร</button><button class="btn primary" data-act="p6print">พิมพ์คนนี้</button><button class="btn" data-act="p6all">พิมพ์ทั้งห้อง</button></div></div></div>
      <div class="tw p5tw"><table class="t p5t"><thead><tr><th>รหัส</th><th class="stk">รายวิชา</th><th>เวลาเรียน</th><th>คะแนนเต็ม</th><th>เฉลี่ยชั้น</th><th>คะแนนที่ได้</th><th>ร้อยละ</th><th>ระดับ</th></tr></thead><tbody>${rows || '<tr><td colspan="8" class="hint">ยังไม่มีรายวิชา</td></tr>'}</tbody>
        <tfoot><tr><td colspan="3" class="stk">รวม · ร้อยละ ${esc(fmt2(get(sh, 30, 4)))}</td><td class="num">${esc(fmt2(get(sh, 29, 5)))}</td><td class="num">${esc(fmt2(get(sh, 29, 6)))}</td><td class="num">${esc(fmt2(get(sh, 29, 7)))}</td><td class="num">${esc(fmt2(get(sh, 29, 8)))}</td><td></td></tr></tfoot></table></div>
      <div class="grid2" style="margin-top:12px"><div class="card"><h3>ผลการเรียนเฉลี่ย</h3><div class="row sp"><span>ผลการเรียนเฉลี่ย</span><b>${esc(fmt2(get(sh, 32, 4)))}</b></div><div class="row sp"><span>คะแนนรวมได้ลำดับที่</span><b>${esc(fmt2(get(sh, 31, 4)))}</b></div><div class="row sp"><span>ผลการเรียนเฉลี่ยได้ลำดับที่</span><b>${esc(fmt2(get(sh, 33, 4)))}</b></div></div>
        <div class="card"><h3>กิจกรรมพัฒนาผู้เรียน</h3>${act}</div><div class="card full"><h3>ผลการประเมิน</h3>${ass}</div></div>`;
  }

  // ---------------------------------------------------------------- เอกสาร (ชีทแบบ Excel)
  function vDocs() {
    const groups = A().sheetGroups();
    return `<div class="bar"><input class="in" data-search placeholder="ค้นหาเอกสาร เช่น ปพ.6, คะแนน, วิชาที่ 3" value="${esc(U.docq)}" style="max-width:420px"></div>
      <p class="hint">เปิดชีทแบบ Excel เพื่อดู แก้ไขช่องที่กรอกได้ หรือพิมพ์ — ข้อมูลเดียวกับที่กรอกในเมนูอื่น</p>
      ${groups.map(g => `<div class="card doc-g"><h3>${esc(g.title)}</h3><div class="dlist">${g.items.map(nm => { const m = /^วิชาที่ (\d+)$/.exec(nm), sn = m ? A().subjName(+m[1]) : ''; return `<div class="ditem" data-q="${esc((nm + ' ' + sn).toLowerCase())}"><button class="open" data-act="doc" data-name="${esc(nm)}">${esc(nm)}${sn ? ` <span class="hint">${esc(sn)}</span>` : ''}</button><button class="btn" data-act="print-doc" data-name="${esc(nm)}">พิมพ์</button></div>`; }).join('')}</div></div>`).join('')}`;
  }
  function docFilter() {
    const q = U.docq.trim().toLowerCase();
    $$('#view .ditem').forEach(d => { d.hidden = !!q && !d.dataset.q.includes(q); });
    $$('#view .doc-g').forEach(g => { g.hidden = !$$('.ditem', g).some(d => !d.hidden); });
  }

  // ---------------------------------------------------------------- คู่มือ
  function vManual() {
    const S = [
      ['เริ่มต้นใช้งาน', `<ol><li>แท็บ <b>ข้อมูล › โรงเรียน</b> กรอกปีการศึกษา ชื่อ ผอ. และครูประจำชั้น</li><li><b>ข้อมูล › นักเรียน</b> กรอกหรือคัดลอกรายชื่อจาก Excel มาวาง (สูงสุด ${MAX_STU} คน)</li><li><b>ข้อมูล › รายวิชา</b> กรอกรหัสวิชา ชื่อวิชา เวลาเรียน และครูผู้สอน</li><li>แท็บ <b>คะแนน</b> เลือกวิชา ตั้งตัวชี้วัดและงาน แล้วกรอกคะแนนทีละภาคเรียน</li><li>แท็บ <b>ประเมิน</b> เลือกระดับคุณลักษณะ ค่านิยม และกิจกรรม</li><li>แท็บ <b>รายงาน ปพ.5 / ปพ.6</b> ดูสรุปผลและพิมพ์</li></ol>`],
      ['การกรอกข้อมูลในตาราง', `<ul><li>คลิกช่องแล้วพิมพ์ · <kbd>Enter</kbd> / <kbd>↓</kbd> ลงช่องล่าง · <kbd>Tab</kbd> ไปช่องขวา</li><li>ตัวชี้วัดพิมพ์ <b>ผ</b> หรือ <b>มผ</b> (หรือกด <kbd>Space</kbd> สลับ)</li><li>คัดลอกหลายช่องจาก Excel แล้ววาง (<kbd>Ctrl</kbd>+<kbd>V</kbd>) ที่ช่องแรกได้ทั้งตาราง</li><li>ระบบบันทึกอัตโนมัติ ดูสถานะที่มุมขวาบน หรือกดปุ่ม “บันทึก”</li></ul>`],
      ['เวลาเรียนและปฏิทิน', `<p>ตั้งปฏิทินการเรียนในชีทปฏิทินก่อน (ข้อมูล › เวลาเรียน) แล้วจึงกรอกเวลาเรียนรายวันในชีทเวลาเรียน กรอก 1 = มา · ป = ป่วย · ล = ลา · ข = ขาด</p>`],
      ['พิมพ์เอกสาร', `<p>เมนู <b>เอกสาร</b> เปิดชีทแบบทางการและกดพิมพ์ได้ หน้า ปพ.6 เลือกนักเรียนแล้วพิมพ์คนเดียว หรือพิมพ์ทั้งห้อง ระบบแสดงตัวอย่างก่อนพิมพ์เสมอ</p>`],
      ['สำรองข้อมูล', `<p>ข้อมูลเก็บบนระบบอัตโนมัติ ควรกด <b>ข้อมูล › สำรองข้อมูล</b> เพื่อดาวน์โหลดไฟล์เก็บไว้เป็นระยะ</p>`],
    ];
    return `<h2>คู่มือการใช้งาน</h2><p class="sub">โปรแกรมแจ้งผลการศึกษา ปพ.5 และ ปพ.6 หลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน พ.ศ. 2551 ระดับประถมศึกษา</p>
      <div class="man-toc">${S.map((s, i) => `<button class="chip" data-act="toc" data-k="${i}">${esc(s[0])}</button>`).join('')}</div>
      ${S.map((s, i) => `<details class="card man" id="man-${i}"${i === 0 ? ' open' : ''}><summary><span class="no">${i + 1}</span>${esc(s[0])}</summary><div class="man-body">${s[1]}</div></details>`).join('')}`;
  }

  root.PPUI = { mount, go, refresh };
})(window);
