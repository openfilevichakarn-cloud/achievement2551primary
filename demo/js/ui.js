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
  const ST = 'ข้อมูลนักเรียน', SUBJ = 'รายวิชา', INFO = 'ข้อมูลพื้นฐาน', ROSTER = 'ข้อมูลเพิ่มเติม';
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
    $('#modal').addEventListener('change', (e) => { if (e.target.id === 'dmc-class') window.__dmcPick = e.target.value; });
    $('#dmc-file').onchange = (e) => { importDmc(e.target.files[0]); e.target.value = ''; };
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
    $$('[data-sum]', scope || document).forEach(td => {
      const [r, c1, c2] = td.dataset.sum.split('|'); const sh = 'วิชาที่ ' + U.subj;
      const x = num(sh, +r, c1), y = num(sh, +r, c2), t = x === null && y === null ? '' : fmt2((x || 0) + (y || 0));
      if (td.textContent !== t) td.textContent = t;
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
    if (sh === ST && r >= 3 && (c === 8 || c === 9)) deriveSex(r);
    return true;
  }
  const markPass = (inp) => { inp.classList.toggle('ok', inp.value === 'ผ'); inp.classList.toggle('no', inp.value === 'มผ'); };
  function onChange(e) {
    if (e.target.dataset && e.target.dataset.attp) { schedChange(e.target); return; }
    if (e.target.dataset && (e.target.dataset.aitem !== undefined || e.target.dataset.agrp !== undefined)) { itemChange(e.target); return; }
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
    if (el.dataset.attp) { if (e.key === 'Enter') { e.preventDefault(); el.blur(); } return; }
    if (e.key === 'Enter' || e.key === 'ArrowDown') { e.preventDefault(); commitInput(el); updateRo($('#view')); gridMove(el, 1, 0); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); commitInput(el); updateRo($('#view')); gridMove(el, -1, 0); }
    else if (e.key === 'Tab') { e.preventDefault(); commitInput(el); updateRo($('#view')); gridMove(el, 0, e.shiftKey ? -1 : 1); }
    else if (e.key === 'ArrowRight' && el.selectionStart === el.value.length) { e.preventDefault(); commitInput(el); gridMove(el, 0, 1); }
    else if (e.key === 'ArrowLeft' && el.selectionEnd === 0) { e.preventDefault(); commitInput(el); gridMove(el, 0, -1); }
    else if (el.dataset.t === 'pass' && e.key === ' ') { e.preventDefault(); el.value = el.value === 'ผ' ? 'มผ' : el.value === 'มผ' ? '' : 'ผ'; commitInput(el); updateRo($('#view')); }
  }
  // วางหลายช่องจาก Excel (แถว/คอลัมน์) เริ่มที่ช่องปัจจุบัน
  function onPaste(e) {
    const el = e.target.closest('input.gc'); if (!el || el.dataset.attp) return;
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
      case 'doc': { if (FORM_DOC.test(b.dataset.name)) { A().showPreview(formDocPages(b.dataset.name)); break; } if (b.dataset.name === 'บิดามารดา' || b.dataset.name === 'ผู้ปกครอง') { A().showPreview(A().pagesFor(b.dataset.name)); break; } if (b.dataset.name === 'ชื่อนักเรียน') { A().showPreview([rosterPage()]); break; } const mv = /^เวลาเรียนวิชา:(\d+)$/.exec(b.dataset.name); if (mv) { A().showPreview(sPages(+mv[1], 'all')); break; } if (b.dataset.name === ATT && calTotal()) { A().showPreview(attPages('all')); break; } if (ASSESS_DOCS[b.dataset.name]) { A().showPreview(assessDocPages(b.dataset.name)); break; } const ms = /^วิชาที่ (\d+)$/.exec(b.dataset.name); if (ms && nStudents()) A().showPreview(subjDocPages(+ms[1])); else A().openSheet(b.dataset.name); break; }
      case 'docx': A().openSheet(b.dataset.name); break;
      case 'print-doc': if (FORM_DOC.test(b.dataset.name)) A().showPreview(formDocPages(b.dataset.name)); else if (b.dataset.name === 'ชื่อนักเรียน') A().showPreview([rosterPage()]); else if (/^เวลาเรียนวิชา:(\d+)$/.test(b.dataset.name)) A().showPreview(sPages(+b.dataset.name.split(':')[1], 'all')); else if (ASSESS_DOCS[b.dataset.name]) A().showPreview(assessDocPages(b.dataset.name)); else A().printSheets([b.dataset.name]); break;
      case 'logo': A().pickLogo(); break;
      case 'logo-del': A().removeLogo(); break;
      case 'dmc': $('#dmc-file').click(); break;
      case 'stdfill': stdFill(k); break;
      case 'stdnorm': stdNormalize(); break;
      case 'stdsort': stdSort(); break;
      case 'cyc': cycle(b); break;
      case 'aedit': set({ aedit: !U.aedit }); break;
      case 'aadd': itemAdd(+b.dataset.g); break;
      case 'adel': itemDel(+b.dataset.g, +b.dataset.i); break;
      case 'setall': setAllLevels(+k); break;
      case 'act': toggleActivity(b); break;
      case 'work-add': workAdd(1); break;
      case 'work-del': workAdd(-1); break;
      case 'stupre': { const cur = U.stuPrefix === undefined ? (() => { for (let i = 0; i < MAX_STU; i++) if (txt(ST, 3 + i, 8).trim()) return true; return false; })() : U.stuPrefix; set({ stuPrefix: !cur }); break; }
      case 'tsub': set({ tsub: k }); break;
      case 'asub-step': { const list = subjects().filter(x => true); const i = list.findIndex(x => x.slot === U.asub); const nx = list[i + (+k)]; if (nx) set({ asub: nx.slot }); break; }
      case 'amonth': set({ amonth: isNaN(+k) ? k : +k }); break;
      case 'aoff': subjOff(U.asub, k); break;
      case 'aov': {   // วนตามลำดับ: ตามรายห้อง → ข → ป → ล → 1 → 2 .. (จำนวนคาบของวันนั้น) → ตามรายห้อง
        const i = +b.dataset.i, per = +b.dataset.per || 1, cur = ovOf(U.asub, k, i);
        const seq = [null, 'ข', 'ป', 'ล', ...Array.from({ length: per }, (_, j) => String(j + 1))];
        setOv(U.asub, k, i, seq[(Math.max(0, seq.indexOf(cur)) + 1) % seq.length]); refresh(); break;
      }
      case 'aovclear': { const a = subjAtt(U.asub); for (const x of sessionsOf(U.asub, 'sum', true).filter(q => q.m === U.amonth)) if (a.ov) delete a.ov[x.k]; A().markDirty(); refresh(); break; }
      case 'aprint': A().showPreview(sPages(U.asub, U.amonth)); break;
      case 'aprintall': A().showPreview(sPages(U.asub, 'all')); break;
      case 'p5att': set({ p5att: k }); break;
      case 'p5attprint': A().showPreview(k === 'class' ? [classAttPage(U.term)] : sPages(U.p5subj, 'all')); break;
      case 'p6attprint': A().showPreview([studentAttPage(U.p6no - 1, U.p6term)]); break;
      case 'calgen': onGenCal(); break;
      case 'calsample': { const y = yearBE(); ['16/05/' + y, '31/10/' + y, '01/11/' + y, '31/03/' + (y + 1)].forEach((v, i) => setv(CAL, TERM_CELLS[i][1], TERM_CELLS[i][0], v)); refresh(); break; }
      case 'calnext': if (!calTotal()) { A().toast('ยังไม่ได้เลือกวันเรียนในปฏิทิน', 'warn'); break; } set({ tsub: 'att' }); break;
      case 'calprint': A().printSheets([CAL]); break;
      case 'caltog': calToggle(b); break;
      case 'attm': set({ attMonth: isNaN(+k) ? k : +k }); break;
      case 'attcyc': attCycle(b); break;
      case 'attday': attFill(U.attMonth, +b.dataset.d); refresh(); break;
      case 'attfill': attFill(U.attMonth); refresh(); break;
      case 'attclear': attClear(U.attMonth); break;
      case 'attprint': A().showPreview(attPages(U.attMonth)); break;
      case 'attprintall': A().showPreview(attPages('all')); break;
      case 'addrows': set({ [b.dataset.key]: (U[b.dataset.key] || 0) + 5 }); break;
      case 'p6step': { const n = nStudents(); if (n) set({ p6no: ((U.p6no - 1 + (+k) + n) % n) + 1 }); break; }
      case 'p6print': printStudent(); break;
      case 'p6all': A().printAllStudents(p6Sheet()); break;
      case 'p6book': { const b = 'เล่ม ปพ.6 รวม'; WB().setValue(b, 'CX4', U.p6no); A().showPreview(formDocPages(b)); break; }
      case 'p6bookall': A().printAllStudents('เล่ม ปพ.6 รวม'); break;
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
    if (x.dataset.f === 'asub') { U.asub = +x.value; refresh(); return; }
    if (x.dataset.f === 'p5subj') { U.p5subj = +x.value; refresh(); return; }
    if (x.dataset.f === 'cal-notes') { setHolidayLines(x.value); return; }
    if (x.dataset.f.startsWith('cal-term')) {
      const i = +x.dataset.f.split('|')[1], k = parseTh(x.value);
      if (x.value.trim() && k === null) { A().toast('กรอกวันที่เป็น วว/ดด/ปปปป (พ.ศ.)', 'warn'); return; }
      const v = k === null ? null : String(k % 100).padStart(2, '0') + '/' + String(Math.floor(k / 100) % 100).padStart(2, '0') + '/' + Math.floor(k / 10000);
      setv(CAL, TERM_CELLS[i][1], TERM_CELLS[i][0], v); x.value = v || ''; return;
    }
    if (x.dataset.f === 'p6no') { U.p6no = +x.value; refresh(); return; }
    if (x.dataset.f === 'bookdate1' || x.dataset.f === 'bookdate2') {
      const term = x.dataset.f === 'bookdate1' ? 1 : 2;
      const v = parseCell('date', x.value);
      if (x.value.trim() && typeof v !== 'number') { A().toast('กรอกวันที่เป็น วว/ดด/ปปปป (พ.ศ.)', 'warn'); return; }
      for (const b of (term === 1 ? [BOOKS[0]] : [BOOKS[1], BOOKS[2]])) setv(b, BOOK_DATE.r, BOOK_DATE.c, v);
      for (const [sh, ref, val] of A().issueTargets(v, term)) if (WB().sheet(sh)) { const m = /^([A-Z]+)(\d+)$/.exec(ref); setv(sh, +m[2], C(m[1]), val); }   // ปก ปพ.5 และ ปพ.6 ตามภาคเรียน
      x.value = rawOf(v, 'date'); return;
    }
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
    for (const t of [1, 2]) for (let i = 0; i < n; i++) for (const c of items('trait', 0)) { all++; if (num('คุณลักษณะ ' + t, 5 + i, c) !== null) done++; }
    return Math.round(done * 100 / all);
  }
  function vHome() {
    const S = A().S, n = nStudents(), ns = subjCount(), sp = scoreProgress(), ap = assessProgress();
    const info = !ph(get(INFO, 9, 2)) && get(INFO, 9, 2) !== null && !ph(get(INFO, 13, 2)) && get(INFO, 13, 2) !== null;
    const step = (no, done, title, sub, tab, subk, prog) => `<button class="step${done ? ' done' : ''}" data-act="go" data-tab="${tab}"${subk ? ` data-sub="${subk}"` : ''}><span class="no">${done ? '✓' : no}</span><span style="flex:1;min-width:0"><b>${title}</b><small>${sub}</small>${prog !== undefined ? `<div class="prog"><i style="width:${prog}%"></i></div>` : ''}</span><span class="go">›</span></button>`;
    return `<div class="card hello">${A().S.view.logo ? `<img class="logo" src="${A().S.view.logo}" alt="ตราโรงเรียน">` : ''}<div><h2>${esc(S.school.name)} 🔒</h2><p>${esc(txt(SUBJ, 2, 2))} · ปีการศึกษา ${esc(txt(SUBJ, 4, 2))}</p></div></div>
      <div class="stat"><div><b>${n}</b><span>นักเรียน</span></div><div><b>${ns}</b><span>รายวิชา</span></div><div><b>${sp}%</b><span>คะแนนกรอกแล้ว</span></div></div>
      <div class="steps">
        ${step(1, info, 'ตั้งค่าโรงเรียน', info ? 'กรอกชื่อผู้อำนวยการและครูประจำชั้นแล้ว' : 'กรอกชื่อ ผอ. ครูประจำชั้น และปีการศึกษา', 'data', 'school')}
        ${step(2, n > 0, 'รายชื่อนักเรียน', n ? n + ' คน' : 'กรอกหรือวางรายชื่อจาก Excel', 'data', 'students')}
        ${step(3, ns > 0, 'รายวิชา', ns ? ns + ' วิชา' : 'เพิ่มรหัสวิชา ชื่อวิชา เวลาเรียน และครูผู้สอน', 'data', 'subjects')}
        ${step(4, sp >= 100, 'บันทึกคะแนน', 'เลือกวิชาแล้วกรอกคะแนนตัวชี้วัดและคะแนนสอบ ทีละภาคเรียน', 'scores', '', sp)}
        ${step(5, ap >= 100, 'ประเมินคุณลักษณะ ค่านิยม และกิจกรรม', 'แตะช่องเพื่อเลือกระดับ / ผ่าน / ไม่ผ่าน', 'assess', '', ap)}
        ${step(6, false, 'รายงาน ปพ.5 และ ปพ.6', 'ดูสรุปผลรายห้อง รายบุคคล และพิมพ์', 'report6')}
      </div>`;
  }

  // ---------------------------------------------------------------- ข้อมูล
  function vData() {
    const sub = U.dsub;
    return `<div class="bar">${seg([['school', 'โรงเรียน'], ['students', 'นักเรียน'], ['note', 'ความเห็นครู'], ['subjects', 'รายวิชา'], ['body', 'น้ำหนัก-ส่วนสูง'], ['time', 'เวลาเรียน'], ['backup', 'สำรองข้อมูล']], sub, 'dsub')}</div>` +
      ({ school: vSchool, students: vStudents, note: vNote, subjects: vSubjects, body: vBody, time: vTime, backup: vBackup }[sub] || vSchool)();
  }
  function vSchool() {
    const S = A().S;
    return `<div class="card"><h3>ข้อมูลโรงเรียนและชั้นเรียน</h3><div class="grid2">
      ${field('ชื่อโรงเรียน (ล็อกตามรหัสสิทธิ์)', `<input value="${esc(S.school.name)}" disabled>`)}
      ${field('ชั้น (ตามห้องที่เปิดอยู่)', `<input value="${esc(txt(SUBJ, 2, 2))}" disabled>`)}
      ${formInput(SUBJ, 4, 2, 'ปีการศึกษา', { t: 'num', ph: 'เช่น 2566' })}
      ${formInput(SUBJ, 3, 2, 'ภาคเรียนที่ (ใช้ในส่วนหัวเอกสาร)', { t: 'num', ph: '1 หรือ 2' })}
      ${field('วันที่ออกเอกสาร ภาคเรียนที่ 1 (วว/ดด/ปปปป พ.ศ.)', `<input data-f="bookdate1" value="${esc(rawOf(bookDate(1), 'date'))}" placeholder="เช่น 31/10/2566">`)}
      ${field('วันที่ออกเอกสาร ภาคเรียนที่ 2 และรวมทั้งปี (วว/ดด/ปปปป พ.ศ.)', `<input data-f="bookdate2" value="${esc(rawOf(bookDate(2), 'date'))}" placeholder="เช่น 31/03/2567">`)}
      ${formInput(INFO, 4, 2, 'สังกัดระดับกรม')}${formInput(INFO, 8, 2, 'สังกัดระดับจังหวัด / เขตพื้นที่การศึกษา')}
      ${formInput(INFO, 5, 2, 'ตำบล / แขวง')}${formInput(INFO, 6, 2, 'อำเภอ / เขต')}${formInput(INFO, 7, 2, 'จังหวัด')}
    </div></div>
    <div class="card"><h3>ผู้ลงนามในเอกสาร</h3><div class="grid2">
      ${formInput(INFO, 9, 2, 'ชื่อ-สกุล ผู้อำนวยการ')}${formInput(INFO, 10, 2, 'ตำแหน่งผู้อำนวยการ / รักษาการ')}
      ${formInput(INFO, 11, 2, 'ชื่อ-สกุล นายทะเบียน')}${formInput(INFO, 12, 2, 'ชื่อ-สกุล หัวหน้างานวิชาการ')}
      ${formInput(INFO, 13, 2, 'ชื่อ-สกุล ครูประจำชั้น / ครูที่ปรึกษา')}
    </div><p class="hint">ข้อมูลนี้เก็บแยกตามห้อง ถ้าโรงเรียนมีหลายห้องต้องกรอกในแต่ละห้อง</p></div>
    <div class="card"><h3>ตราโรงเรียน</h3>
      <div class="row">${S.view.logo ? `<img src="${S.view.logo}" alt="ตราโรงเรียน" style="height:84px;width:84px;object-fit:contain;background:#fff;border:1px solid var(--line);border-radius:12px">` : '<span class="hint">ยังไม่ได้ใส่ตรา</span>'}
        <button class="btn primary" data-act="logo">เลือกรูปตรา</button>${S.view.logo ? '<button class="btn danger" data-act="logo-del">ลบตรา</button>' : ''}</div>
      <p class="hint">ตราจะแสดงที่มุมซ้ายบนของปก ปพ.5 และ ปพ.6 ใช้ร่วมกันทุกชั้นทุกห้อง (ไฟล์ PNG/JPG)</p></div>`;
  }
  const PREFIX = ['เด็กชาย', 'เด็กหญิง', 'นาย', 'นางสาว', 'นาง'];
  // เพศ (คอลัมน์ G) ใช้นับชาย/หญิงในปก ปพ.5 — เดาจากคำนำหน้า/ต้นชื่อ เมื่อยังไม่ได้กรอก
  function deriveSex(r) {
    if (A().S.wb.get(A().S.wb.sheet(ST), r, 7)) return;
    const t = (txt(ST, r, 8) + txt(ST, r, 9)).replace(/\s/g, '');
    const sex = /^(เด็กชาย|ด\.?ช\.?|นาย|master)/i.test(t) ? 'ช' : /^(เด็กหญิง|ด\.?ญ\.?|นางสาว|นาง|น\.ส\.|miss)/i.test(t) ? 'ญ' : null;
    if (sex) setv(ST, r, 7, sex);
  }
  // รายชื่อ + ข้อมูลส่วนตัวนักเรียน เรียงหัวข้อตามไฟล์ DMC (กรอกเองหรือนำเข้าไฟล์ DMC .xlsx)
  // [ป้าย, คอลัมน์ในชีท ข้อมูลนักเรียน, ชนิด, class]
  const STU_GROUPS = [
    ['ข้อมูลนักเรียน', [['เลขประจำตัวประชาชน', 3, 'text', 'wide'], ['เลขประจำตัวนักเรียน', 6, 'text'], ['เพศ (ช/ญ)', 7, 'sex'], ['คำนำหน้า', 8, 'text', 'pre'], ['ชื่อ', 9, 'text', 'wide'], ['นามสกุล', 10, 'text', 'wide'], ['วันเกิด (วว/ดด/ปปปป)', 13, 'date'], ['กลุ่มเลือด', 16, 'text'], ['ศาสนา', 19, 'text'], ['เชื้อชาติ', 17, 'text'], ['สัญชาติ', 18, 'text']]],
    ['ที่อยู่', [['บ้านเลขที่', 46, 'text'], ['หมู่', 47, 'text'], ['ถนน/ซอย', 48, 'text'], ['ตำบล', 49, 'text'], ['อำเภอ', 50, 'text'], ['จังหวัด', 51, 'text'], ['รหัสไปรษณีย์', 52, 'text']]],
    ['ผู้ปกครอง', [['คำนำหน้า', 40, 'text'], ['ชื่อ', 41, 'text', 'wide'], ['นามสกุล', 42, 'text', 'wide'], ['อาชีพ', 83, 'text', 'wide'], ['ความเกี่ยวข้องกับนักเรียน', 38, 'text']]],
    ['บิดา', [['คำนำหน้า', 27, 'text'], ['ชื่อ', 28, 'text', 'wide'], ['นามสกุล', 29, 'text', 'wide'], ['อาชีพ', 81, 'text', 'wide']]],
    ['มารดา', [['คำนำหน้า', 33, 'text'], ['ชื่อ', 34, 'text', 'wide'], ['นามสกุล', 35, 'text', 'wide'], ['อาชีพ', 82, 'text', 'wide']]],
    ['อื่น ๆ', [['สถานภาพสมรสของบิดา-มารดา', 25, 'text', 'wide']]],
  ];
  function vStudents() {
    const n = nStudents(), rows = Math.min(MAX_STU, Math.max(n + 3, 10, U.stuRows || 0));
    let hasPre = false; for (let i = 0; i < MAX_STU; i++) if (txt(ST, 3 + i, 8).trim()) { hasPre = true; break; }
    const showPre = U.stuPrefix === undefined ? hasPre : U.stuPrefix;
    const groups = STU_GROUPS.map(([g, cols]) => [g, cols.filter(c => c[3] !== 'pre' || showPre)]);
    let h1 = '<th class="no" rowspan="2">ที่</th>', h2 = '';
    for (const [g, cols] of groups) { h1 += `<th colspan="${cols.length}">${esc(g)}</th>`; h2 += cols.map(c => `<th>${esc(c[0] === 'ชื่อ' && g === 'ข้อมูลนักเรียน' && !showPre ? 'ชื่อ (รวมคำนำหน้า)' : c[0])}</th>`).join(''); }
    let body = '';
    for (let i = 0; i < rows; i++) {
      const r = 3 + i; let gc = 0; body += `<tr><td class="no">${i + 1}</td>`;
      for (const [, cols] of groups) for (const c of cols) body += gcell(ST, r, c[1], c[2], i, gc++, { cls: c[3] === 'pre' ? '' : c[3] || '', ph: c[3] === 'pre' ? 'เด็กชาย' : '' });
      body += '</tr>';
    }
    return `<datalist id="prefixes">${PREFIX.map(p => `<option value="${p}">`).join('')}</datalist>
      <div class="card soft"><b>รายชื่อนักเรียน ${n} คน</b> (สูงสุด ${MAX_STU} คน)
        <p class="hint" style="margin:.3rem 0 0">หัวข้อเรียงตามไฟล์ DMC · กรอกเอง หรือกด “นำเข้าไฟล์ DMC” · วางข้อมูลหลายช่องจาก Excel ได้ (Ctrl+V) · ข้อมูลทั้งหมดนี้ใช้ในเล่ม ปพ.6</p>
        <div class="row" style="margin-top:8px"><button class="btn primary" data-act="dmc">นำเข้าไฟล์ DMC (.xlsx)</button><button class="btn sm" data-act="stupre">${showPre ? '－ ซ่อนคำนำหน้า' : '＋ เพิ่มคำนำหน้า'}</button>
          <button class="btn sm" data-act="stdnorm" title="ล้างช่องว่างเกิน แปลงคำนำหน้าย่อ (ด.ช. ด.ญ. น.ส.) แยกคำนำหน้า/นามสกุลที่ติดกัน แปลงวันเกิดและเลขบัตร ฯลฯ">ปรับข้อมูล (หลังวาง)</button><button class="btn sm" data-act="stdsort">เรียงลำดับ (ชายก่อน · เลขประจำตัว)</button><button class="btn sm" data-act="stdfill" data-k="def">ตั้ง เชื้อชาติ/สัญชาติ = ไทย ศาสนา = พุทธ ให้ที่ว่าง</button><button class="btn sm" data-act="stdfill" data-k="addr">เติมตำบล/อำเภอ/จังหวัด/ไปรษณีย์ จากคนที่ 1</button></div></div>
      <div class="tw"><table class="t sheet prof"><thead><tr>${h1}</tr><tr>${h2}</tr></thead><tbody>${body}</tbody></table></div>
      ${rows < MAX_STU ? '<p><button class="btn" data-act="addrows" data-key="stuRows">＋ เพิ่ม 5 แถว</button></p>' : ''}`
      .replace(/(<input class="gc"[^>]*data-c="8")/g, '$1 list="prefixes"');
  }
  // เติมค่าให้ช่องที่ว่าง
  function stdFill(kind) {
    const n = nStudents(); if (!n) return;
    const specs = kind === 'def' ? [[17, 'ไทย'], [18, 'ไทย'], [19, 'พุทธ']] : [[49], [50], [51], [52]];
    for (const [col, dflt] of specs) {
      const base = dflt !== undefined ? dflt : get(ST, 3, col); if (base === null || base === '') continue;
      for (let i = 0; i < n; i++) { const v = get(ST, 3 + i, col); if (v === null || v === '') setv(ST, 3 + i, col, base); }
    }
    refresh();
  }
  // ---- ปรับข้อมูล (กรณีคัดลอกมาวาง) และเรียงลำดับ
  const STU_ALL = () => [...new Set([...Object.values(root.PPDmc.COL), 25, 52, 4, 5])];
  const PRE_MAP = [[/^(ด\.?\s*ช\.?|เด็กชาย|master)$/i, 'เด็กชาย'], [/^(ด\.?\s*ญ\.?|เด็กหญิง)$/i, 'เด็กหญิง'], [/^(น\.?\s*ส\.?|นางสาว|miss)$/i, 'นางสาว'], [/^(นาย|mr\.?)$/i, 'นาย'], [/^(นาง|mrs\.?)$/i, 'นาง']];
  const normPrefix = (s) => { for (const [re, v] of PRE_MAP) if (re.test(s)) return v; return s; };
  const PRE_START = /^(เด็กชาย|เด็กหญิง|ด\.ช\.|ด\.ญ\.|ดช\.?|ดญ\.?|นางสาว|น\.ส\.|นส\.|นาย|นาง)\s*/;
  const stuRowCount = () => { let m = 0; for (let i = 0; i < MAX_STU; i++) if (txt(ST, 3 + i, 8).trim() || txt(ST, 3 + i, 9).trim() || txt(ST, 3 + i, 10).trim()) m = i + 1; return m; };
  function stdNormalize() {
    const n = stuRowCount(); if (!n) { A().toast('ยังไม่มีรายชื่อนักเรียน', 'warn'); return; }
    let changed = 0; A().S.batch = true;
    const put = (r, c, v) => { const old = get(ST, r, c); const o = old === null ? '' : String(old), nv = v === null ? '' : String(v); if (o !== nv) { setv(ST, r, c, v); changed++; } };
    try {
      for (let i = 0; i < n; i++) {
        const r = 3 + i, cur = {};
        for (const c of STU_ALL()) { const v = get(ST, r, c); cur[c] = v === null || (v && v.isRange) ? '' : (typeof v === 'number' && c !== 13 ? String(v) : v); }
        const text = (c) => String(cur[c] === undefined ? '' : cur[c]).replace(/[ ​‌‍﻿]/g, ' ').replace(/\s+/g, ' ').trim();
        for (const c of STU_ALL()) {
          if (c === 13 || c === 3 || c === 5) continue;
          const raw = cur[c]; let s = typeof raw === 'number' ? String(raw) : text(c); if (s === '-' && c !== 7) s = '';
          put(r, c, s === '' ? null : s);
        }
        // คำนำหน้า: รูปแบบย่อ → เต็ม · แยกคำนำหน้า/นามสกุลที่ติดมาในช่องชื่อ
        let pre = normPrefix(text(8)), first = text(9), last = text(10);
        if (!pre) { const m = PRE_START.exec(first); if (m && first.length > m[0].length) { pre = normPrefix(m[1]); first = first.slice(m[0].length); } }
        if (!last && /\s/.test(first)) { const k = first.indexOf(' '); last = first.slice(k + 1).trim(); first = first.slice(0, k); }
        put(r, 8, pre || null); put(r, 9, first || null); put(r, 10, last || null);
        for (const c of [27, 33, 40]) put(r, c, normPrefix(text(c)) || null);
        // เลขประจำตัวประชาชน: เหลือเฉพาะตัวเลข
        const pid = String(get(ST, r, 3) === null ? '' : get(ST, r, 3)).replace(/\D/g, ''); put(r, 3, pid || null);
        // วันเกิด: ข้อความ → วันที่ (ปี ค.ศ./พ.ศ. ได้)
        const b = get(ST, r, 13); if (typeof b === 'string' && b.trim()) { const d = A().parseInput(ST, 13, b.trim()); if (typeof d === 'number') put(r, 13, d); }
        // เพศ
        const sx = parseCell('sex', text(7)); put(r, 7, sx); if (!sx) deriveSex(r);
      }
    } finally { A().S.batch = false; A().markDirty(); }
    refresh(); A().toast(changed ? 'ปรับข้อมูลแล้ว ' + changed + ' ช่อง' : 'ข้อมูลอยู่ในรูปแบบที่ถูกต้องแล้ว', 'ok');
  }
  // เรียง: เพศชายก่อนเพศหญิง แต่ละกลุ่มเรียงเลขประจำตัวนักเรียนจากน้อยไปมาก
  const sexRank = (s) => (s === 'ช' ? 0 : s === 'ญ' ? 1 : 2);
  const codeKey = (s) => { s = String(s === null || s === undefined ? '' : s).trim(); const m = /^\d+$/.test(s); return m ? [0, +s] : s ? [1, s] : [2, 0]; };
  function sortStudentList(list, sexOf, codeOf) {
    return list.map((x, i) => ({ x, i, s: sexRank(sexOf(x)), k: codeKey(codeOf(x)) })).sort((a, b) => a.s - b.s || a.k[0] - b.k[0] || (a.k[0] === 1 ? String(a.k[1]).localeCompare(String(b.k[1]), 'th', { numeric: true }) : a.k[1] - b.k[1]) || a.i - b.i).map(o => o.x);
  }
  async function stdSort() {
    const n = stuRowCount(); if (n < 2) return;
    const yes = await A().modal('เรียงลำดับนักเรียน', '<p>เรียงใหม่: <b>เพศชายก่อน</b> แล้วตามด้วยเพศหญิง แต่ละกลุ่มเรียงตาม <b>เลขประจำตัวนักเรียน</b> จากน้อยไปมาก</p><p class="hint">⚠ เลขที่ของนักเรียนจะเปลี่ยน และคะแนน/การประเมิน/เวลาเรียนที่กรอกไว้แล้วผูกกับเลขที่เดิม ควรเรียงก่อนเริ่มกรอกคะแนน</p>', [{ label: 'ยกเลิก', value: false }, { label: 'เรียงลำดับ', value: true, primary: true }]);
    if (!yes) return;
    const cols = STU_ALL(), com = [27, 28, 29, 30, 31];
    const list = Array.from({ length: n }, (_, i) => ({ r: get(ST, 3 + i, 7), c: get(ST, 3 + i, 6), v: Object.fromEntries(cols.map(c => [c, get(ST, 3 + i, c)])), k: Object.fromEntries(com.map(c => [c, get(ROSTER, 2 + i, c)])) }));
    const sorted = sortStudentList(list, o => o.r, o => o.c);
    A().S.batch = true;
    try {
      sorted.forEach((o, i) => { for (const c of cols) { const v = o.v[c] === undefined ? null : o.v[c]; if ((get(ST, 3 + i, c) ?? null) !== v) setv(ST, 3 + i, c, v); } for (const c of com) { const v = o.k[c] === undefined ? null : o.k[c]; if ((get(ROSTER, 2 + i, c) ?? null) !== v) setv(ROSTER, 2 + i, c, v); } });
    } finally { A().S.batch = false; A().markDirty(); }
    refresh(); A().toast('เรียงลำดับแล้ว ' + n + ' คน', 'ok');
  }
  // นำเข้าไฟล์ DMC (.xlsx) — แทนที่รายชื่อนักเรียนของห้องนี้
  async function importDmc(file) {
    if (!file) return;
    let res;
    try { res = await root.PPDmc.parse(file); } catch (e) { A().toast('อ่านไฟล์ไม่ได้: ' + (e.message || e), 'warn'); return; }
    if (!res.classes.length) { A().toast('ไม่พบรายชื่อนักเรียนในไฟล์', 'warn'); return; }
    const S = A().S, want = 'ป.' + S.grade;
    const def = res.classes.find(c => c.cls === want && String(c.room) === String(S.room)) || res.classes.find(c => c.cls === want) || res.classes[0];
    let pick = def.key;
    const opts = res.classes.map(c => `<option value="${esc(c.key)}"${c.key === def.key ? ' selected' : ''}>${esc(c.label)} · ${c.rows.length} คน</option>`).join('');
    const had = nStudents();
    const ok = await A().modal('นำเข้าไฟล์ DMC', `<p>เลือกชั้น/ห้องในไฟล์ที่จะนำเข้ามายัง <b>${esc(A().roomLabel(S.grade, S.room))}</b></p>
      <label class="fld"><span>ชั้น/ห้องในไฟล์</span><select id="dmc-class">${opts}</select></label>
      <p class="hint">${had ? `⚠ รายชื่อเดิม ${had} คนในห้องนี้จะถูกแทนที่ (คะแนนที่กรอกไว้ผูกกับเลขที่ จะตรงกับรายชื่อใหม่ตามลำดับ)` : 'นำเข้าเป็นรายชื่อใหม่ของห้องนี้'}</p><p class="hint">จัดเรียงให้อัตโนมัติ: เพศชายก่อน แล้วตามด้วยเพศหญิง แต่ละกลุ่มเรียงตามเลขประจำตัวนักเรียนจากน้อยไปมาก</p>`,
      [{ label: 'ยกเลิก', value: false }, { label: 'นำเข้า', value: true, primary: true }].map(b => b));
    // modal ลบ DOM เมื่อกดปุ่ม — อ่านตัวเลือกที่เลือกไว้ล่าสุดจากเหตุการณ์ change
    if (!ok) return;
    const cls = res.classes.find(c => c.key === (window.__dmcPick || pick)) || def;
    window.__dmcPick = null;
    A().S.batch = true;
    try {
      const old = Math.max(had, 3), cols = [...new Set([...Object.values(root.PPDmc.COL), 25, 52])];
      for (let i = 0; i < Math.max(old, cls.rows.length); i++) for (const c of cols) if (get(ST, 3 + i, c) !== null && get(ST, 3 + i, c) !== '') setv(ST, 3 + i, c, null);
      for (let i = 0; i < Math.max(old, cls.rows.length); i++) for (const c of [27, 28, 29, 30, 31]) if (get(ROSTER, 2 + i, c) !== null && get(ROSTER, 2 + i, c) !== '') setv(ROSTER, 2 + i, c, null);
      const C_ = root.PPDmc.COL;
      sortStudentList(cls.rows, o => o.sex || (/^(เด็กชาย|นาย)/.test(o.prefix) ? 'ช' : /^(เด็กหญิง|นางสาว|นาง)/.test(o.prefix) ? 'ญ' : ''), o => o.code).slice(0, MAX_STU).forEach((o, i) => {
        const r = 3 + i;
        for (const [k, c] of Object.entries(C_)) {
          let v = o[k]; if (v === undefined || v === '') continue;
          if (k === 'birth') { const d = A().parseInput(ST, 13, v); v = d === null ? v : d; }
          setv(ST, r, c, v);
        }
        setv(ST, r, 4, 'ประถมศึกษาปีที่ ' + S.grade); setv(ST, r, 5, S.room);
      });
      A().S.userEdits[ST] = A().S.userEdits[ST] || {};
    } finally { A().S.batch = false; A().markDirty(); }
    U.stuPrefix = undefined; refresh(); A().toast('นำเข้า ' + Math.min(MAX_STU, cls.rows.length) + ' คน จาก ' + cls.label, 'ok');
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
      <div class="tw"><table class="t sheet subj"><thead><tr><th class="no">ที่</th><th>รหัสวิชา</th><th>รายวิชา</th><th>เวลาเรียน (ชม./ปี)</th><th>กลุ่มสาระการเรียนรู้</th><th>ครูประจำวิชา</th></tr></thead><tbody>${body}</tbody></table></div>
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
  // ---------------------------------------------------------------- ข้อมูลส่วนตัวนักเรียน (ใช้ในเล่ม ปพ.6)
  // เก็บในชีท ข้อมูลนักเรียน (แถว 3+) และ ข้อมูลเพิ่มเติม (ความเห็นครู แถว 2+) — เล่ม ปพ.6 ดึงไปแสดงเอง
  // หมายเหตุ: ป้าย เชื้อชาติ/สัญชาติ ตรงกับที่เล่ม ปพ.6 พิมพ์ (ในไฟล์ต้นฉบับคอลัมน์ Q/R ป้ายสลับกัน)
  const PROF = {
    note: { title: 'ความเห็นครูประจำชั้น', cols: [['ด้านหน้าที่รับผิดชอบ ความเอาใจใส่การเรียน', ROSTER, 27, 'text', 2, 'note'], ['ด้านการใช้เวลาว่าง', ROSTER, 28, 'text', 2, 'note'], ['ด้านความสัมพันธ์กับบุคคลรอบข้าง', ROSTER, 29, 'text', 2, 'note'], ['ด้านอุปนิสัย บุคลิกภาพ', ROSTER, 30, 'text', 2, 'note'], ['ด้านสุขภาพ', ROSTER, 31, 'text', 2, 'note']] },
  };
  function vNote() {
    const n = nStudents(); U.psub = 'note';
    const head = '';
    if (!n) return head + '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่แท็บ “นักเรียน” ก่อน</div>';
    const P_ = PROF[U.psub];
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>' + P_.cols.map(c => `<th>${esc(c[0])}</th>`).join(''), body = '';
    for (let i = 0; i < n; i++) {
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>${P_.cols.map((c, j) => gcell(c[1], c[4] + i, c[2], c[3], i, j, { cls: c[5] || '' })).join('')}</tr>`;
    }
    const tools = P_.fill ? `<div class="bar"><button class="btn sm" data-act="proffill">${esc(P_.fillLabel)}</button></div>` : '';
    const hint = U.psub === 'note' ? 'ข้อความสั้น ๆ ต่อด้าน จะพิมพ์ในหน้าความคิดเห็นของเล่ม ปพ.6' : 'คลิกช่องแล้วพิมพ์ · วางข้อมูลหลายช่องจาก Excel ได้ (Ctrl+V) · ข้อมูลนี้แสดงในเล่ม ปพ.6 (รายงานประจำตัวนักเรียน)';
    return head + `<p class="hint">${hint}</p>${tools}<div class="tw"><table class="t sheet prof ${U.psub}"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  // วันที่ออกเอกสาร ปพ.6 (ใช้คำนวณอายุและวันที่ลงนามในเล่ม ปพ.6 ทั้ง 3 ฉบับ)
  const BOOKS = ['เล่ม ปพ.6 เทอม 1', 'เล่ม ปพ.6 เทอม 2', 'เล่ม ปพ.6 รวม'], BOOK_DATE = { r: 20, c: 'CX' };
  const bookDate = (term) => get(BOOKS[term === 1 ? 0 : 1], BOOK_DATE.r, BOOK_DATE.c);

  // ---------------------------------------------------------------- ปฏิทินวันเรียน · เวลาเรียน
  // ชีท "ปฏิทิน": แถว 5-10 = พ.ค.-ต.ค. (เทอม 1) · แถว 12-17 = พ.ย.-เม.ย. (เทอม 2) · คอลัมน์ B.. = วันที่ 1..31
  // ใส่อักษรย่อวันในสัปดาห์ในวันที่เรียน (จ อ พ พฤ ศ) หรือ ส / อา สำหรับเรียนชดเชย — ชีท "เวลาเรียน" ดึงวันเรียนไปเองอัตโนมัติ
  const CAL = 'ปฏิทิน', ATT = 'เวลาเรียน';
  const MONTHS = ['พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน'];
  const MON_G = [4, 5, 6, 7, 8, 9, 10, 11, 0, 1, 2, 3];   // เดือนแบบ JS (0 = ม.ค.)
  const DOW = ['อา', 'จ', 'อ', 'พ', 'พฤ', 'ศ', 'ส'];
  const calRow = (m) => (m < 6 ? 5 + m : 6 + m);
  const yearBE = () => num(SUBJ, 4, 2) || 2566;
  const monthBE = (m) => yearBE() + (m >= 8 ? 1 : 0);
  const daysIn = (m) => new Date(Date.UTC(monthBE(m) - 543, MON_G[m] + 1, 0)).getUTCDate();
  const dow = (m, d) => new Date(Date.UTC(monthBE(m) - 543, MON_G[m], d)).getUTCDay();
  const calDay = (m, d) => txt(CAL, calRow(m), 1 + d).trim();     // อักษรวันในปฏิทิน ('' = ไม่ใช่วันเรียน)
  const calCount = (m) => { let n = 0; for (let d = 1; d <= daysIn(m); d++) if (calDay(m, d)) n++; return n; };
  const calTotal = () => MONTHS.reduce((a, _, m) => a + calCount(m), 0);
  const TERM_CELLS = [['AP', 2], ['AQ', 2], ['AP', 3], ['AQ', 3]];   // วันเริ่ม/สิ้นสุดของภาคเรียน (เก็บเป็นข้อความ วว/ดด/ปปปป) นอกพื้นที่พิมพ์
  const parseTh = (t) => { const m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/.exec(String(t || '').trim()); if (!m) return null; let y = +m[3]; if (y < 2400) y += 543; return y * 10000 + (+m[2]) * 100 + (+m[1]); };
  const termDate = (i) => txt(CAL, TERM_CELLS[i][1], TERM_CELLS[i][0]);
  const thaiLong = (key) => { const y = Math.floor(key / 10000), mo = Math.floor(key / 100) % 100, d = key % 100; return d + ' ' + MONTHS[MON_G.indexOf(mo - 1)] + ' ' + y; };

  function genCalendar() {
    const t = [0, 1, 2, 3].map(i => parseTh(termDate(i)));
    if (t.some(x => x === null) || t[0] > t[1] || t[2] > t[3]) { A().toast('กรอกวันเปิดและวันปิดเรียนของทั้ง 2 ภาคเรียนให้ครบก่อน', 'warn'); return false; }
    for (let m = 0; m < 12; m++) for (let d = 1; d <= 31; d++) {
      let letter = null;
      if (d <= daysIn(m)) {
        const key = monthBE(m) * 10000 + (MON_G[m] + 1) * 100 + d, w = dow(m, d);
        if (((key >= t[0] && key <= t[1]) || (key >= t[2] && key <= t[3])) && w >= 1 && w <= 5) letter = DOW[w];
      }
      if (calDay(m, d) !== (letter || '')) setv(CAL, calRow(m), 1 + d, letter);
    }
    setv(CAL, 2, 1, thaiLong(t[0]) + ' - ' + thaiLong(t[3]));
    return true;
  }
  async function onGenCal() {
    if (calTotal() > 0) {
      const yes = await A().modal('สร้างปฏิทินใหม่', '<p>ปฏิทินที่กำหนดไว้แล้ว (รวมวันหยุดที่ปิดไว้) จะถูกแทนที่ด้วยวันจันทร์-ศุกร์ตามช่วงภาคเรียนที่ระบุ</p>', [{ label: 'ยกเลิก', value: false }, { label: 'สร้างปฏิทินใหม่', value: true, primary: true }]);
      if (!yes) return;
    }
    if (genCalendar()) { refresh(); A().toast('สร้างปฏิทินแล้ว — คลิกวันที่หยุด (วันหยุดราชการ ฯลฯ) เพื่อปิดวันเรียน', 'ok'); }
  }
  function calToggle(b) {
    const m = +b.dataset.m, d = +b.dataset.d, cur = calDay(m, d);
    setv(CAL, calRow(m), 1 + d, cur ? null : DOW[dow(m, d)]);
    b.classList.toggle('on', !cur);
    const card = b.closest('.mcard'); if (card) $('.mc-n', card).textContent = calCount(m) + ' วัน';
    const t = $('#cal-total'); if (t) t.textContent = calTotal();
  }
  // สีพื้นของช่องวันที่ในชีทปฏิทิน (ในไฟล์ต้นฉบับเป็นสีคงที่ของปี 2566) — คำนวณใหม่ตามวันในสัปดาห์จริงของปีการศึกษาที่เลือก
  function beforeSheet(name) {
    if (name !== CAL) return;
    const t = [0, 1, 2, 3].map(i => parseTh(termDate(i))), bg = {};
    const inTerm = (k) => t.every(x => x !== null) && ((k >= t[0] && k <= t[1]) || (k >= t[2] && k <= t[3]));
    for (let m = 0; m < 12; m++) for (let d = 1; d <= 31; d++) {
      const ref = L(1 + d) + calRow(m);
      if (d > daysIn(m)) { bg[ref] = '#000000'; continue; }
      const w = dow(m, d), on = !!calDay(m, d), key = monthBE(m) * 10000 + (MON_G[m] + 1) * 100 + d;
      bg[ref] = on ? (w === 0 ? '#ff0000' : w === 6 ? '#8064a2' : '#ffffff') : w === 0 ? '#ff0000' : w === 6 ? '#8064a2' : inTerm(key) ? '#ffff00' : '#bfbfbf';
    }
    A().S.view.bg[CAL] = bg;
  }
  function holidayLines() {
    const out = [];
    for (let r = 21; r <= 30; r++) { const v = txt(CAL, r, 'M'); if (v.trim()) out.push(v.trim()); }
    for (let r = 21; r <= 29; r++) { const v = txt(CAL, r, 'Z'); if (v.trim()) out.push(v.trim()); }
    return out;
  }
  function setHolidayLines(text) {
    const lines = String(text).split('\n').map(x => x.trim()).filter(Boolean).slice(0, 19);
    for (let i = 0; i < 10; i++) setv(CAL, 21 + i, 'M', lines[i] || null);
    for (let i = 0; i < 9; i++) setv(CAL, 21 + i, 'Z', lines[10 + i] || null);
  }
  function monthCard(m) {
    const first = dow(m, 1), nd = daysIn(m);
    let cells = '<span class="cdh">' + DOW.map(x => `<i>${x}</i>`).join('') + '</span><span class="cdg">';
    for (let i = 0; i < first; i++) cells += '<i class="blank"></i>';
    for (let d = 1; d <= nd; d++) { const w = dow(m, d); cells += `<button class="cd${calDay(m, d) ? ' on' : ''}${w === 0 || w === 6 ? ' we' : ''}" data-act="caltog" data-m="${m}" data-d="${d}">${d}</button>`; }
    return `<div class="card mcard"><div class="row sp"><b>${MONTHS[m]} ${monthBE(m)}</b><span class="mc-n chip">${calCount(m)} วัน</span></div>${cells}</span></div>`;
  }
  function vTime() {
    const tabs = `<div class="bar">${seg([['cal', '1. ปฏิทิน'], ['att', '2. เวลาเรียนรายห้อง'], ['sched', '3. ตารางสอน'], ['subj', '4. เวลาเรียนรายวิชา']], U.tsub || 'cal', 'tsub')}</div>`;
    const t = U.tsub || 'cal';
    return tabs + (t === 'att' ? vAttendance() : t === 'sched' ? vSchedule() : t === 'subj' ? vSubjAtt() : vCalendar());
  }
  function vCalendar() {
    const f = (i, label) => `<label class="fld"><span>${label}</span><input data-f="cal-term|${i}" value="${esc(termDate(i))}" placeholder="วว/ดด/ปปปป (พ.ศ.)"></label>`;
    const total = calTotal(), t1 = [0, 1, 2, 3, 4, 5].reduce((a, m) => a + calCount(m), 0);
    return `<div class="card"><h3>กำหนดภาคเรียน ปีการศึกษา ${yearBE()}</h3>
      <div class="grid2">${f(0, 'ภาคเรียนที่ 1 · วันเปิดเรียน')}${f(1, 'ภาคเรียนที่ 1 · วันปิดเรียน')}${f(2, 'ภาคเรียนที่ 2 · วันเปิดเรียน')}${f(3, 'ภาคเรียนที่ 2 · วันปิดเรียน')}</div>
      <div class="row" style="margin-top:10px"><button class="btn primary" data-act="calgen">สร้างปฏิทินวันเรียน (จันทร์–ศุกร์)</button><button class="btn sm" data-act="calsample">ใช้ช่วงภาคเรียนตัวอย่าง</button></div>
      <p class="hint">ระบบใส่วันจันทร์–ศุกร์ในช่วงที่กำหนดเป็นวันเรียน แล้วคลิกวันที่ด้านล่างเพื่อปิดวันหยุด (หรือเปิดวันเสาร์-อาทิตย์ที่เรียนชดเชย) ปีการศึกษาแก้ได้ที่ ข้อมูล › โรงเรียน</p></div>
      <div class="card soft row sp"><span>จำนวนวันเรียนทั้งปี <b id="cal-total">${total}</b> วัน <span class="hint">(ภาคเรียนที่ 1 = ${t1} วัน)</span></span>
        <span class="row"><button class="btn" data-act="calprint">พิมพ์ปฏิทิน</button><button class="btn primary" data-act="calnext">เสร็จแล้ว → เช็คเวลาเรียนรายห้อง</button></span></div>
      <div class="months">${MONTHS.map((_, m) => monthCard(m)).join('')}</div>
      <div class="card"><h3>หมายเหตุวันหยุด (แสดงใต้ปฏิทินที่พิมพ์)</h3>
        <textarea class="in" rows="5" data-f="cal-notes" placeholder="บรรทัดละ 1 วัน เช่น&#10;พ. 17 พฤษภาคม วันพืชมงคล&#10;ศ. 28 กรกฎาคม วันเฉลิมพระชนมพรรษา ร.10">${esc(holidayLines().join('\n'))}</textarea>
        <p class="hint">สูงสุด 19 บรรทัด</p></div>`;
  }

  // ---- เวลาเรียนรายเดือน
  // ชีท "เวลาเรียน": 12 บล็อกต่อเนื่องตามเดือน พ.ค.→เม.ย. แต่ละบล็อก = ช่องวันที่ 1..n แล้วตามด้วยสรุป มา ป ล ข
  const _blocks = new WeakMap();
  function attBlocks() {
    const wb = WB(); if (_blocks.has(wb)) return _blocks.get(wb);
    const sh = wb.sheet(ATT), out = [];
    for (let c = 3; c <= sh.maxC; c++) {
      if (String(txt(ATT, 3, c)) !== '1') continue;
      let d = 1; while (String(txt(ATT, 3, c + d)) === String(d + 1)) d++;
      out.push({ start: c, days: d, sum: c + d });
    }
    if (out.length === 12) _blocks.set(wb, out);
    return out;
  }
  const MARK = { 1: '✓', 'ข': 'ข', 'ป': 'ป', 'ล': 'ล' };
  const attKind = (v) => (v === 1 || v === '1' ? 1 : MARK[v] ? v : null);
  // ---- สรุปเวลาเรียนรายภาคเรียน (ต่อจากเดือนเมษายน): เทอม 1 · เทอม 2 · รวม
  const TERM_KEY = { T1: 1, T2: 2, S: 'sum' };
  const isTermKey = (k) => typeof k === 'string' && k in TERM_KEY;
  const termLabel = (t) => (t === 'sum' ? 'รวมทั้งปี' : 'ภาคเรียนที่ ' + t);
  const termBtns = (act, cur) => [['T1', 'สรุปเทอม 1'], ['T2', 'สรุปเทอม 2'], ['S', 'สรุปรวม']].map(([k, t]) => `<button role="tab" data-act="${act}" data-k="${k}" class="${cur === k ? 'on' : ''}">${t}</button>`).join('');
  const monthsOf = (term) => MONTHS.map((_, m) => m).filter(m => inTerm(m, term));
  const shortMon = (m) => MONTHS[m].slice(0, 3) + '.';
  // สถิติเวลาเรียนรายห้อง (หน่วย = วัน) ของนักเรียนคนที่ i ในภาคเรียน/ทั้งปี · pct คิดจากวันที่เช็คแล้ว
  function classStat(i, term) {
    const b = attBlocks(), r = { plan: 0, checked: 0, present: 0, sick: 0, leave: 0, absent: 0, pct: null, per: {} };
    for (const m of monthsOf(term)) {
      r.per[m] = 0; if (!b[m]) continue;
      for (let d = 1; d <= Math.min(daysIn(m), b[m].days); d++) {
        if (!calDay(m, d)) continue;
        r.plan++; const k = attKind(get(ATT, 5 + i, b[m].start + d - 1)); if (!k) continue;
        r.checked++; if (k === 1) { r.present++; r.per[m]++; } else if (k === 'ป') r.sick++; else if (k === 'ล') r.leave++; else if (k === 'ข') r.absent++;
      }
    }
    if (r.checked) r.pct = r.present * 100 / r.checked;
    return r;
  }
  // ตารางสรุปบนหน้าจอ (รายห้อง): คอลัมน์ = เดือนในภาคเรียน (หรือเทอม 1 / เทอม 2 เมื่อรวมทั้งปี)
  function vAttSummary(term) {
    const n = nStudents(), low = (p) => p !== null && p < PCT_MIN;
    const cols = term === 'sum' ? [[1, 'เทอม 1'], [2, 'เทอม 2']] : monthsOf(term).filter(m => calCount(m)).map(m => [m, shortMon(m)]);
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>' + cols.map(([, t]) => `<th>${esc(t)}<br><small>มา (วัน)</small></th>`).join('') + '<th>รวมมา</th><th>ป่วย</th><th>ลา</th><th>ขาด</th><th>วันที่เช็ค<br><small>จากวันเรียน</small></th><th>ร้อยละ</th>', body = '';
    for (let i = 0; i < n; i++) {
      const st = classStat(i, term), t1 = term === 'sum' ? classStat(i, 1) : null, t2 = term === 'sum' ? classStat(i, 2) : null;
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>` + cols.map(([c]) => `<td class="num">${(term === 'sum' ? (c === 1 ? t1 : t2).present : st.per[c]) || '-'}</td>`).join('')
        + `<td class="num"><b>${st.present || '-'}</b></td><td class="num">${st.sick || '-'}</td><td class="num">${st.leave || '-'}</td><td class="num">${st.absent || '-'}</td><td class="num">${st.checked}/${st.plan}</td><td class="num${low(st.pct) ? ' warnc' : ''}"><b>${fmtPct(st.pct)}</b></td></tr>`;
    }
    return `<div class="card soft row sp"><span><b>สรุปเวลาเรียน ${termLabel(term)}</b> <span class="hint">ดูอย่างเดียว คำนวณจากการเช็คชื่อรายวันของครูประจำชั้น · ร้อยละ = วันที่มา ÷ วันที่เช็คแล้ว</span></span><span class="row"><button class="btn primary" data-act="attprint">พิมพ์สรุปนี้</button><button class="btn" data-act="attprintall">พิมพ์ทุกเดือน + สรุปทั้ง 3</button></span></div>
      <div class="tw"><table class="t sheet p5t"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  function vAttendance() {
    const n = nStudents(), blocks = attBlocks();
    if (calTotal() === 0) return '<div class="card empty"><b>ยังไม่ได้สร้างปฏิทินวันเรียน</b>ต้องกำหนดวันเรียนก่อน ระบบจึงจะรู้ว่าแต่ละเดือนมีวันเรียนกี่วัน<br><br><button class="btn primary" data-act="tsub" data-k="cal">ไปสร้างปฏิทิน</button></div>';
    if (!n) return '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่ ข้อมูล › นักเรียน ก่อน</div>';
    if (blocks.length < 12) return '<div class="card empty"><b>อ่านผังเวลาเรียนไม่ได้</b>ลองเปิดชีทเวลาเรียนในเมนูเอกสาร</div>';
    let m = U.attMonth; if (!isTermKey(m) && (m === undefined || !calCount(m))) m = MONTHS.findIndex((_, i) => calCount(i) > 0); U.attMonth = m;
    const mseg = `<div class="seg" role="tablist">${MONTHS.map((nm, i) => `<button role="tab" data-act="attm" data-k="${i}" class="${i === m ? 'on' : ''}"${calCount(i) ? '' : ' disabled'}>${esc(nm.length > 7 ? nm.slice(0, 5) + '.' : nm)}</button>`).join('')}${termBtns('attm', m)}</div>`;
    if (isTermKey(m)) return `<div class="bar">${mseg}</div>` + vAttSummary(TERM_KEY[m]);
    const mb = blocks[m], days = []; for (let d = 1; d <= mb.days; d++) if (calDay(m, d)) days.push(d);
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>', body = '';
    for (const d of days) th += `<th class="dh" data-act="attday" data-d="${d}" title="กดเพื่อเติม “มา” ทั้งห้องในวันนี้">${d}<br><small>${esc(calDay(m, d))}</small></th>`;
    th += '<th class="ro">มา</th><th class="ro">ป่วย</th><th class="ro">ลา</th><th class="ro">ขาด</th>';
    for (let i = 0; i < n; i++) {
      const r = 5 + i; body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`;
      for (const d of days) body += attCell(r, mb.start + d - 1);
      for (let k = 0; k < 4; k++) body += rocell(ATT, r, mb.sum + k);
      body += '</tr>';
    }
    return `<div class="bar">${mseg}</div>
      <div class="card soft row sp"><span><b>${MONTHS[m]} ${monthBE(m)}</b> · วันเรียน ${days.length} วัน <span class="hint">แตะช่อง: ว่าง → ✓ มา → ข ขาด → ป ป่วย → ล ลา → ว่าง · กดตัวเลขวันที่หัวตารางเพื่อเติม “มา” ทั้งห้อง</span></span>
        <span class="row"><button class="btn sm" data-act="attfill">มาครบทั้งเดือน (เติมช่องว่าง)</button><button class="btn sm danger" data-act="attclear">ล้างเดือนนี้</button></span></div>
      <div class="tw"><table class="t sheet att"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>
      <div class="card row sp" style="margin-top:12px"><span class="hint">ครูประจำชั้นเช็คที่นี่ครั้งเดียว เวลาเรียนรายวิชาจะดึงไปใช้เอง</span><span class="row"><button class="btn" data-act="tsub" data-k="sched">ต่อไป: ตารางสอน →</button><button class="btn" data-act="attprint">พิมพ์เดือนนี้</button><button class="btn primary" data-act="attprintall">พิมพ์ทุกเดือน + สรุปเทอม 1/2/รวม</button></span></div>`;
  }
  function attCell(r, c) {
    const k = attKind(get(ATT, r, c));
    return `<td class="gtd"><button class="pc att${k === 1 ? ' ok' : k ? ' no' : ''}" data-act="attcyc" data-r="${r}" data-c="${c}">${k ? MARK[k] : '&nbsp;'}</button></td>`;
  }
  const ATT_NEXT = { '': 1, 1: 'ข', 'ข': 'ป', 'ป': 'ล', 'ล': null };
  function attCycle(b) {
    const r = +b.dataset.r, c = +b.dataset.c, k = attKind(get(ATT, r, c)), next = ATT_NEXT[k === null ? '' : k];
    setv(ATT, r, c, next === undefined ? null : next);
    b.className = 'pc att' + (next === 1 ? ' ok' : next ? ' no' : ''); b.innerHTML = next ? MARK[next] : '&nbsp;';
    updateRo($('#view'));
  }
  function attFill(m, only) {
    const n = nStudents(), mb = attBlocks()[m];
    for (let d = 1; d <= mb.days; d++) {
      if (!calDay(m, d) || (only && only !== d)) continue;
      for (let i = 0; i < n; i++) { const c = mb.start + d - 1; if (get(ATT, 5 + i, c) === null) setv(ATT, 5 + i, c, 1); }
    }
  }
  async function attClear(m) {
    const yes = await A().modal('ล้างเวลาเรียน', `<p>ล้างการเช็คชื่อเดือน <b>${MONTHS[m]}</b> ของนักเรียนทุกคน</p>`, [{ label: 'ยกเลิก', value: false }, { label: 'ล้าง', value: true, primary: true }]);
    if (!yes) return;
    const n = nStudents(), mb = attBlocks()[m];
    for (let d = 1; d <= mb.days; d++) for (let i = 0; i < n; i++) { const c = mb.start + d - 1; if (get(ATT, 5 + i, c) !== null) setv(ATT, 5 + i, c, null); }
    refresh();
  }

  // ---- พิมพ์เวลาเรียน: 1 หน้า A4 แนวตั้งต่อเดือน + หน้าสรุปทั้งปี (ออกแบบใหม่ ไม่ใช้ผังชีท Excel)
  const PF = "font-family:'TH Sarabun New','Sarabun',sans-serif;color:#000;";
  const pcell = (x, o) => `<td style="border:1px solid #000;padding:0 2px;text-align:center;${o || ''}">${x}</td>`;
  const sign = (label, name) => `<div>ลงชื่อ ................................................<br>(${esc(name || '................................................')})<br>${label}</div>`;
  const realName = (r) => { const v = get(INFO, r, 2); return v === null || ph(v) ? '' : String(v); };
  function attHead(title) {
    return `<div style="${PF}text-align:center;line-height:1.35;margin-bottom:6px"><div style="font-size:21px;font-weight:700">${esc(title)}</div><div style="font-size:17px">โรงเรียน${esc(A().S.school.name.replace(/^โรงเรียน/, ''))} ${esc(txt(SUBJ, 2, 2))} ปีการศึกษา ${yearBE()}</div></div>`;
  }
  function attMonthPage(m, blocks) {
    const n = nStudents(), mb = blocks[m], nd = daysIn(m), grey = (d) => (calDay(m, d) ? '' : 'background:#d9d9d9;');
    let h = '<tr><th rowspan="2" style="border:1px solid #000;width:26px">ที่</th><th rowspan="2" style="border:1px solid #000;width:190px">ชื่อ - สกุล</th>';
    for (let d = 1; d <= nd; d++) h += `<th style="border:1px solid #000;font-weight:400;width:20px;${grey(d)}">${d}</th>`;
    h += '<th colspan="4" style="border:1px solid #000">สรุป (วัน)</th></tr><tr>';
    for (let d = 1; d <= nd; d++) h += `<th style="border:1px solid #000;font-weight:400;font-size:12px;${grey(d)}">${esc(calDay(m, d))}</th>`;
    h += ['มา', 'ป', 'ล', 'ข'].map(x => `<th style="border:1px solid #000;width:28px">${x}</th>`).join('') + '</tr>';
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const r = 5 + i; rows += `<tr style="height:19px">${pcell(i + 1)}<td style="border:1px solid #000;padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>`;
      for (let d = 1; d <= nd; d++) { const k = d <= mb.days && calDay(m, d) ? attKind(get(ATT, r, mb.start + d - 1)) : null; rows += pcell(k === 1 ? '/' : k ? k : '', grey(d)); }
      for (let k = 0; k < 4; k++) { const v = get(ATT, r, mb.sum + k); rows += pcell(v === null || v === '' ? '-' : esc(String(v))); }
      rows += '</tr>';
    }
    const html = `<div style="${PF}font-size:15px">${attHead('บันทึกเวลาเรียน ประจำเดือน' + MONTHS[m] + ' ' + monthBE(m))}
      <div style="margin:0 0 4px">จำนวนวันเรียนในเดือนนี้ <b>${calCount(m)}</b> วัน &nbsp; เครื่องหมาย: / = มาเรียน · ป = ป่วย · ล = ลา · ข = ขาด · ช่องสีเทา = ไม่ใช่วันเรียน</div>
      <table style="border-collapse:collapse;font-size:14px;width:100%">${h}${rows}</table>
      <div style="display:flex;justify-content:space-around;margin-top:26px;font-size:16px;text-align:center">${sign('ครูประจำชั้น', realName(13))}</div></div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  // หน้าสรุปเวลาเรียนรายห้อง: term = 1 | 2 | 'sum'
  function attSummaryPage(term) {
    const n = nStudents(), ms = monthsOf(term).filter(m => calCount(m));
    const th = (x, w) => `<th rowspan="2" style="border:1px solid #000;width:${w}px">${x}</th>`;
    let h = `<tr>${th('ที่', 26)}${th('ชื่อ - สกุล', 180)}<th colspan="${ms.length}" style="border:1px solid #000">จำนวนวันที่มาเรียน (ใต้ชื่อเดือน = วันเรียนของเดือน)</th>${th('รวมมา', 40)}${th('ป่วย', 34)}${th('ลา', 34)}${th('ขาด', 34)}${th('วันเรียน<br>ที่เช็ค', 42)}${th('ร้อยละ', 46)}</tr><tr>`;
    h += ms.map(m => `<th style="border:1px solid #000;font-weight:400;font-size:12px;width:34px">${esc(shortMon(m))}<br>${calCount(m)}</th>`).join('') + '</tr>';
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const st = classStat(i, term);
      rows += `<tr style="height:19px">${pcell(i + 1)}<td style="border:1px solid #000;padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>${ms.map(m => pcell(st.per[m] || '-')).join('')}${[st.present, st.sick, st.leave, st.absent, st.checked].map(x => pcell(x || '-')).join('')}${pcell(fmtPct(st.pct), st.pct !== null && st.pct < PCT_MIN ? 'font-weight:700;' : '')}</tr>`;
    }
    const planned = ms.reduce((t, m) => t + calCount(m), 0);
    const html = `<div style="${PF}font-size:15px">${attHead('สรุปเวลาเรียน ' + termLabel(term) + ' ปีการศึกษา ' + yearBE())}<div style="margin:0 0 4px">จำนวนวันเรียนตามปฏิทิน <b>${planned}</b> วัน · ร้อยละ = วันที่มา ÷ วันที่เช็คแล้ว · เกณฑ์ไม่น้อยกว่าร้อยละ ${PCT_MIN}</div>
      <table style="border-collapse:collapse;font-size:14px;width:100%;margin-top:6px">${h}${rows}</table>
      <div style="display:flex;justify-content:space-around;margin-top:26px;font-size:16px;text-align:center">${sign('ครูประจำชั้น', realName(13))}</div></div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  function attPages(months) {
    if (!calTotal()) { A().toast('ยังไม่ได้สร้างปฏิทินวันเรียน', 'warn'); return []; }
    const blocks = attBlocks(); if (blocks.length < 12) return [];
    if (isTermKey(months)) return [attSummaryPage(TERM_KEY[months])];
    const list = months === 'all' ? MONTHS.map((_, i) => i).filter(i => calCount(i)) : [months];
    const pages = list.map(m => attMonthPage(m, blocks));
    if (months === 'all') pages.push(attSummaryPage(1), attSummaryPage(2), attSummaryPage('sum'));
    return pages;
  }
  // หน้าพิมพ์ที่ประกอบพิเศษ: เวลาเรียนรายวิชาต่อท้ายเอกสาร ปพ.5 / ปพ.6 ทุกครั้งที่พิมพ์ (ถ้าตั้งตารางสอนแล้ว)
  const hasSched = () => calTotal() > 0 && subjects().some(x => weekPeriods(x.slot));
  const termOfName = (n) => (/รวม$/.test(n) ? 'sum' : /เทอม ?2$|\b2$/.test(n) ? 2 : 1);
  // ---------------------------------------------------------------- บันทึกผลรายวิชา (เอกสาร › วิชาที่ N): แยกหน้า เทอม 1 · เทอม 2 · รวม
  // คอลัมน์ตัวชี้วัดดึงจากเมนู คะแนน ของวิชานั้น (ชื่อ + คะแนนเต็ม + คะแนนที่ได้) ตั้งแต่เลขที่ถึงผลการประเมิน
  const SB = 'border:1px solid #000;';
  const sv = (v) => (v === null || v === undefined || v === '' || (v && v.code) || (v && v.isRange)) ? '' : fmt2(v);
  function subjHead(slot, title, sub) {
    const ratio = [num('วิชาที่ ' + slot, 3, 'DA'), num('วิชาที่ ' + slot, 3, 'DC')];
    return `${attHead(title)}<div style="${PF}display:flex;justify-content:space-between;font-size:15px;margin:0 2px 4px"><span>รายวิชา <b>${esc(subjName(slot))}</b> รหัส ${esc(txt(SUBJ, 5 + slot, 2) || '-')} · ${esc(txt(SUBJ, 5 + slot, 4) ? 'เวลาเรียน ' + txt(SUBJ, 5 + slot, 4) + ' ชั่วโมง' : '')}</span><span>${sub || ''} สัดส่วนระหว่างเรียน : ปลายภาค = ${fmt2(ratio[0])} : ${fmt2(ratio[1])}</span></div>`;
  }
  function subjFoot(slot) {
    return `<div style="${PF}display:flex;justify-content:space-around;margin-top:16px;font-size:16px;text-align:center">${sign('ครูผู้สอน', txt(SUBJ, 5 + slot, 6))}</div>`;
  }
  function scoreTermPage(slot, term) {
    const sh = 'วิชาที่ ' + slot, T = SC[term], n = nStudents();
    let nw = 0; for (let c = T.work[0]; c <= T.work[1]; c++) if ((num(sh, SC_MAX, c) || 0) > 0) nw = c - T.work[0] + 1;
    const th = (x, o) => `<th style="${SB}font-weight:400;${o || ''}">${x}</th>`;
    let h = `<tr>${th('ที่', 'width:26px')}${th('ชื่อ - สกุล', 'width:190px')}`;
    for (let j = 0; j < nw; j++) h += th(`${esc(txt(sh, SC_NAME, T.work[0] + j) || 'ตัวชี้วัดที่ ' + (j + 1))}<br><small>(${fmt2(num(sh, SC_MAX, T.work[0] + j))})</small>`, 'min-width:44px');
    h += th(`ระหว่างเรียน<br><small>(${fmt2(num(sh, 3, 'DA'))})</small>`) + th(`สอบปลายภาค<br><small>(${fmt2(num(sh, 3, 'DC'))})</small>`) + th('รวม<br><small>(100)</small>') + th('ร้อยละ') + th('ระดับผล<br>การเรียน') + '</tr>';
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const r = IND_ROW + i;
      rows += `<tr style="height:19px">${pcell(i + 1)}<td style="${SB}padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>`;
      for (let j = 0; j < nw; j++) rows += pcell(sv(get(sh, r, T.work[0] + j)));
      rows += [T.btw, T.exam, T.total, T.pct].map(c => pcell(sv(get(sh, r, c)))).join('') + pcell(sv(get(sh, r, T.grade)), 'font-weight:700;') + '</tr>';
    }
    const html = `<div style="${PF}font-size:15px">${subjHead(slot, 'บันทึกผลการเรียนรายวิชา ' + subjName(slot), 'ภาคเรียนที่ ' + term + ' ·')}<table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center">${h}${rows}</table>${subjFoot(slot)}</div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  function scoreYearPage(slot) {
    const sh = 'วิชาที่ ' + slot, n = nStudents();
    const th = (x, o) => `<th style="${SB}font-weight:400;${o || ''}">${x}</th>`;
    const sub = ['คะแนนเก็บ', 'สอบ', 'รวม', 'ร้อยละ', 'ระดับ'];
    let h = `<tr>${th('ที่', 'width:26px', 2)}${th('ชื่อ - สกุล', 'width:190px')}${['ภาคเรียนที่ 1', 'ภาคเรียนที่ 2', 'รวมทั้งปี'].map(x => `<th colspan="5" style="${SB}">${x}</th>`).join('')}</tr><tr><th style="${SB}"></th><th style="${SB}"></th>${[0, 1, 2].map(() => sub.map(x => th(x)).join('')).join('')}</tr>`;
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const r = IND_ROW + i, g = (c) => get(sh, r, c), a = (c1, c2) => { const x = g(c1), y = g(c2); return typeof x === 'number' || typeof y === 'number' ? (x || 0) + (y || 0) : ''; };
      rows += `<tr style="height:19px">${pcell(i + 1)}<td style="${SB}padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>`
        + [g(C('EI')), g(C('EJ')), g(C('EK')), g(C('EL'))].map(v => pcell(sv(v))).join('') + pcell(sv(g(SC[1].grade)), 'font-weight:700;')
        + [g(C('EM')), g(C('EN')), g(C('EO')), g(C('EP'))].map(v => pcell(sv(v))).join('') + pcell(sv(g(SC[2].grade)), 'font-weight:700;')
        + [a(C('EI'), C('EM')), a(C('EJ'), C('EN')), g(C('EQ')), g(C('ER'))].map(v => pcell(sv(v))).join('') + pcell(sv(g(C('ES'))), 'font-weight:700;') + '</tr>';
    }
    const html = `<div style="${PF}font-size:15px">${subjHead(slot, 'สรุปผลการเรียนรายวิชา ' + subjName(slot) + ' รวมทั้งปี', '')}<table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center">${h}${rows}</table>${subjFoot(slot)}</div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  const subjDocPages = (slot) => [scoreTermPage(slot, 1), scoreTermPage(slot, 2), scoreYearPage(slot)];

  // ปก ปพ.5 รายวิชา: หนึ่งหน้าต่อรายวิชา (หน้าแสดงผลและหน้าพิมพ์ใช้ชุดเดียวกัน)
  function coverSubjPages(name) {
    const list = subjects(); if (!list.length) return null;
    const prev = get(name, 1, C('EM')), out = [];
    try { for (const s of list) { WB().setValue(name, 'EM1', s.slot); out.push(...A().pagesFor(name)); } } finally { WB().setValue(name, 'EM1', prev); }
    return out;
  }
  // เอกสารฟอร์มต้นฉบับ (ปก ปพ.5 · ปพ.6 · เล่ม ปพ.6 · คำชี้แจง · เอกสารแนบท้าย): ดูตัวอย่างแบบเดียวกับหน้าพิมพ์ (ไม่เปิดเป็นชีท Excel)
  const FORM_DOC = /^(ปก ปพ\.5 |ปพ\.6 |เล่ม ปพ\.6|คำชี้แจง ปพ\.5|เอกสารแนบท้าย)/;
  function formDocPages(name) {
    if (/^ปพ\.6 /.test(name) && nStudents()) WB().setValue(name, 'K8', Math.min(U.p6no, nStudents()));
    if (/^เล่ม ปพ\.6/.test(name) && nStudents()) WB().setValue(name, 'CX4', Math.min(U.p6no, nStudents()));
    return customPages(name) || A().pagesFor(name);
  }
  function customPages(name) {
    const ms = /^วิชาที่ (\d+)$/.exec(name); if (ms) return subjDocPages(+ms[1]);
    if (ASSESS_DOCS[name]) return assessDocPages(name);
    if (name === 'ชื่อนักเรียน') return [rosterPage()];
    if (name === ATT) return attPages('all');
    if (/^ปก ปพ\.5 รายวิชา/.test(name)) return coverSubjPages(name);
    return null;
  }

  // ---------------------------------------------------------------- เวลาเรียนรายวิชา (แต่ละวิชาสอนคนละวัน/คนละจำนวนคาบ)
  // ไม่ต้องเช็คชื่อซ้ำ: ดึงจากเวลาเรียนรายห้อง (ครูประจำชั้นเช็คทุกวัน) — ถ้ามาเรียนวันนั้น ถือว่ามาทุกวิชาที่สอนในวันนั้น
  // ครูประจำวิชาแก้เฉพาะวิชาของตนได้ เมื่อนักเรียนเรียนไม่เต็มวัน (เช่น มาสาย/ออกก่อน/ไปกิจกรรม) — แก้ทับเวลาเรียนรายห้องเฉพาะวิชานั้น
  // ข้อมูลที่เก็บเพิ่ม: doc.att[ลำดับวิชา] = { p: [คาบต่อวัน อา..ส], off: [วันที่ไม่มีสอน 'ปปปปดดวว'], ov: { 'ปปปปดดวว': 'สตริงต่อนักเรียน' } }
  // ov: ตัวอักษรตำแหน่ง i = นักเรียนคนที่ i+1  - = ตามเวลาเรียนรายห้อง · 1 มา · ข ขาด · ป ป่วย · ล ลา
  const EMPTY_ATT = () => ({ p: [0, 0, 0, 0, 0, 0, 0], off: [], ov: {} });
  const attAll = () => (A().S.att = A().S.att || {});
  const getAtt = (slot) => attAll()[slot] || EMPTY_ATT();
  const subjAtt = (slot) => (attAll()[slot] = attAll()[slot] || EMPTY_ATT());
  const dkey = (m, d) => String(monthBE(m)) + String(MON_G[m] + 1).padStart(2, '0') + String(d).padStart(2, '0');
  const inTerm = (m, term) => term === 'sum' || term === undefined || (term === 1 ? m < 6 : m >= 6);
  const weekPeriods = (slot) => getAtt(slot).p.reduce((a, b) => a + (b || 0), 0);
  // วันที่มีสอนของวิชา (ปฏิทินกลาง × ตารางสอน) — includeOff = แสดงวันที่ปิดไว้ด้วย
  function sessionsOf(slot, term, includeOff) {
    const a = getAtt(slot), out = [];
    for (let m = 0; m < 12; m++) {
      if (!inTerm(m, term)) continue;
      for (let d = 1; d <= daysIn(m); d++) {
        if (!calDay(m, d)) continue;
        const per = a.p[dow(m, d)] || 0; if (!per) continue;
        const k = dkey(m, d), off = a.off.includes(k);
        if (off && !includeOff) continue;
        out.push({ m, d, k, per, off });
      }
    }
    return out;
  }
  // สถานะของนักเรียนคนที่ i ในวันที่ s จากเวลาเรียนรายห้อง: '1' มา · 'ป' 'ล' 'ข' · null ยังไม่เช็ค
  // จำนวนคาบที่นักเรียนเรียนในวันนั้น: { pres: คาบที่มา, miss: คาบที่ไม่ได้เรียน, reason: 'ข' | 'ป' | 'ล' } · null = ยังไม่เช็ค
  // ครูประจำวิชาแก้ทับได้ด้วยเลข 1-9 (= เรียนกี่คาบ เช่น วันที่สอน 2 คาบ มาแค่ 1 คาบ) หรือ ข ป ล (ไม่ได้เรียนเลยวันนั้น)
  function dayAmt(s, i, slot) {
    const o = slot ? ovOf(slot, s.k, i) : null;
    if (o) {
      if (/^[1-9]$/.test(o)) { const n = Math.min(+o, s.per); return { pres: n, miss: s.per - n, reason: 'ข' }; }
      return { pres: 0, miss: s.per, reason: o };
    }
    const h = homeCh(s, i); if (!h) return null;
    return h === '1' ? { pres: s.per, miss: 0, reason: 'ข' } : { pres: 0, miss: s.per, reason: h };
  }
  // ข้อความในช่อง: เลขคาบที่เรียน (มาไม่ครบแสดงเลขที่ได้เรียน) หรือ ข/ป/ล เมื่อไม่ได้เรียนเลย
  const amtLabel = (a) => (!a ? '' : a.pres > 0 ? String(a.pres) : a.reason);
  const amtCls = (a) => (!a ? '' : a.miss === 0 ? ' ok' : a.pres > 0 ? ' part' : ' no');
  const homeCh = (s, i) => {
    const b = attBlocks()[s.m]; if (!b || s.d > b.days) return null;
    const k = attKind(get(ATT, 5 + i, b.start + s.d - 1));
    return k === 1 ? '1' : k;
  };
  const ovOf = (slot, k, i) => { const o = ((getAtt(slot).ov || {})[k] || '')[i]; return o && o !== '-' ? o : null; };
  function setOv(slot, k, i, ch) {
    const a = subjAtt(slot); a.ov = a.ov || {};
    const cur = (a.ov[k] || '').padEnd(MAX_STU, '-').split(''); cur[i] = ch || '-'; const str = cur.join('');
    if (/^-+$/.test(str)) delete a.ov[k]; else a.ov[k] = str;
    A().markDirty();
  }
  // สรุปเวลาเรียนของนักเรียนคนที่ i (หน่วย = คาบ) · pct คิดจากคาบที่ครูประจำชั้นเช็คแล้ว
  function attStat(slot, i, term) {
    const r = { plan: 0, checked: 0, present: 0, sick: 0, leave: 0, absent: 0, pct: null };
    for (const s of sessionsOf(slot, term)) {
      r.plan += s.per; const a = dayAmt(s, i, slot);
      if (!a) continue;
      r.checked += s.per; r.present += a.pres;
      if (a.miss) { if (a.reason === 'ป') r.sick += a.miss; else if (a.reason === 'ล') r.leave += a.miss; else r.absent += a.miss; }
    }
    if (r.checked) r.pct = r.present * 100 / r.checked;
    return r;
  }
  const PCT_MIN = 80;
  const fmtPct = (p) => (p === null ? '-' : p.toFixed(2));

  // ---- ตารางสอน: คาบต่อวันของแต่ละวิชา
  function vSchedule() {
    const list = subjects();
    if (!list.length) return '<div class="card empty"><b>ยังไม่มีรายวิชา</b>เพิ่มที่ ข้อมูล › รายวิชา ก่อน</div>';
    const head = DOW.map((x, w) => `<th${w === 0 || w === 6 ? ' class="we"' : ''}>${x}</th>`);
    const order = [1, 2, 3, 4, 5, 6, 0];
    let body = '';
    for (const s of list) {
      const a = getAtt(s.slot), wk = weekPeriods(s.slot), yr = sessionsOf(s.slot, 'sum').reduce((t, x) => t + x.per, 0), hrs = num(SUBJ, s.row, 4);
      body += `<tr><td class="stk" style="min-width:12rem">${esc(s.name)}</td>${order.map(w => `<td class="gtd"><input class="gc" inputmode="numeric" data-attp="${s.slot}|${w}" value="${a.p[w] || ''}" placeholder="-"></td>`).join('')}<td class="num">${wk || '-'}</td><td class="num">${yr || '-'}</td><td class="num">${hrs === null ? '-' : fmt2(hrs)}</td></tr>`;
    }
    return `<div class="card soft"><b>ตารางสอนรายวิชา</b><p class="hint" style="margin:.3rem 0 0">กรอก <b>จำนวนคาบ</b> ที่วิชานั้นสอนในแต่ละวัน (เว้นว่างหรือ 0 = ไม่สอน) ระบบดึงวันที่สอนจริงจากปฏิทินวันเรียน (ตัดวันหยุดให้เอง) แล้วดึงเวลาเรียนรายห้องของครูประจำชั้นมาคิดเป็นเวลาเรียนรายวิชา (มาในวันนั้น = มาทุกวิชาที่สอนวันนั้น) · ถ้าบางวันไม่มีสอน (เช่น ครูติดราชการ) ปิดวันนั้นได้ที่หน้าเวลาเรียนรายวิชา</p></div>
      ${calTotal() ? '' : '<div class="card"><b>ยังไม่ได้สร้างปฏิทินวันเรียน</b> — ตัวเลข “ทั้งปี” จะเป็น 0 จนกว่าจะสร้างปฏิทิน <button class="btn sm" data-act="tsub" data-k="cal">ไปสร้างปฏิทิน</button></div>'}
      <div class="tw"><table class="t sheet"><thead><tr><th class="stk">รายวิชา</th>${order.map(w => head[w]).join('')}<th>รวม/สัปดาห์<br><small>คาบ</small></th><th>ทั้งปีตามปฏิทิน<br><small>คาบ</small></th><th>เวลาเรียนที่กำหนด<br><small>ชม./ปี</small></th></tr></thead><tbody>${body}</tbody></table></div>
      <div class="card row sp" style="margin-top:12px"><span class="hint">ตั้งค่าเสร็จแล้วดูเวลาเรียนรายวิชาได้เลย (ดึงจากเวลาเรียนรายห้องอัตโนมัติ)</span><button class="btn primary" data-act="tsub" data-k="subj">ดูเวลาเรียนรายวิชา →</button></div>`;
  }
  function schedChange(inp) {
    const [slot, w] = inp.dataset.attp.split('|').map(Number);
    const v = String(inp.value).trim() === '' ? 0 : Number(inp.value);
    if (isNaN(v) || v < 0 || v > 9 || !Number.isInteger(v)) { A().toast('กรอกจำนวนคาบเป็นเลข 0-9', 'warn'); inp.value = getAtt(slot).p[w] || ''; return; }
    if ((getAtt(slot).p[w] || 0) === v) return;
    subjAtt(slot).p[w] = v; A().markDirty(); refresh();
  }

  // ---- ดูเวลาเรียนรายวิชา (ดึงจากเวลาเรียนรายห้อง)
  function vSubjAtt() {
    const list = subjects(), n = nStudents();
    if (!list.length) return '<div class="card empty"><b>ยังไม่มีรายวิชา</b>เพิ่มที่ ข้อมูล › รายวิชา ก่อน</div>';
    if (!calTotal()) return '<div class="card empty"><b>ยังไม่ได้สร้างปฏิทินวันเรียน</b><br><br><button class="btn primary" data-act="tsub" data-k="cal">ไปสร้างปฏิทิน</button></div>';
    if (!n) return '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่ ข้อมูล › นักเรียน ก่อน</div>';
    if (!list.some(s => s.slot === U.asub)) U.asub = list[0].slot;
    const idx = list.findIndex(s => s.slot === U.asub), slot = U.asub;
    const pick = `<div class="pick"><button class="btn" data-act="asub-step" data-k="-1" ${idx <= 0 ? 'disabled' : ''}>‹</button><select data-f="asub" aria-label="เลือกวิชา">${list.map(s => `<option value="${s.slot}"${s.slot === slot ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}</select><button class="btn" data-act="asub-step" data-k="1" ${idx >= list.length - 1 ? 'disabled' : ''}>›</button></div>`;
    const info = '<div class="card soft"><b>ดึงจากเวลาเรียนรายห้องอัตโนมัติ</b><p class="hint" style="margin:.3rem 0 0">ครูประจำชั้นเช็คชื่อทุกวันที่ “เวลาเรียนรายห้อง” — ถ้านักเรียนมาในวันนั้น ถือว่ามาทุกวิชาที่สอนในวันนั้น · <b>ครูประจำวิชาแก้เฉพาะวิชาของตนได้</b> ถ้านักเรียนเรียนไม่เต็มวัน (เช่น ออกก่อนคาบนี้) โดยแตะช่อง: ตามรายห้อง → ข → ป → ล → 1 → 2 (เลข = จำนวนคาบที่ได้เรียนวันนั้น) → กลับตามรายห้อง ช่องที่แก้จะมีเส้นประล้อมรอบ · ตัวเลขในช่อง = จำนวนคาบที่เรียน (วิชาที่สอน 2 คาบต่อวัน มาครบแสดง 2)</p><div class="row" style="margin-top:8px"><button class="btn sm" data-act="tsub" data-k="att">ไปเช็คชื่อรายห้อง</button></div></div>';
    if (!weekPeriods(slot)) return pick + info + `<div class="card empty"><b>วิชานี้ยังไม่ได้ตั้งตารางสอน</b>กรอกจำนวนคาบต่อวันก่อน<br><br><button class="btn primary" data-act="tsub" data-k="sched">ไปตั้งตารางสอน</button></div>`;
    const all = sessionsOf(slot, 'sum', true), months = MONTHS.map((_, m) => m).filter(m => all.some(s => s.m === m));
    let m = U.amonth; if (!isTermKey(m) && (m === undefined || !months.includes(m))) m = months[0]; U.amonth = m;
    const mseg = `<div class="seg" role="tablist">${months.map(i => `<button role="tab" data-act="amonth" data-k="${i}" class="${i === m ? 'on' : ''}">${esc(MONTHS[i].length > 7 ? MONTHS[i].slice(0, 5) + '.' : MONTHS[i])}</button>`).join('')}${termBtns('amonth', m)}</div>`;
    if (isTermKey(m)) return pick + info + `<div class="bar">${mseg}</div>` + vSubjSummary(slot, TERM_KEY[m], list[idx].name);
    const days = all.filter(s => s.m === m);
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>';
    for (const s of days) th += `<th class="dh${s.off ? ' off' : ''}" data-act="aoff" data-k="${s.k}" title="${s.off ? 'วันนี้ปิดไว้ (ไม่มีสอน) — กดเพื่อเปิด' : 'กดเพื่อปิดวันนี้ (ไม่มีสอนวิชานี้)'}">${s.d}<br><small>${esc(DOW[dow(m, s.d)])}${s.per > 1 ? ' ' + s.per + 'ค' : ''}</small></th>`;
    th += '<th class="ro">มา</th><th class="ro">ป</th><th class="ro">ล</th><th class="ro">ข</th><th class="ro">% สะสม</th>';
    let body = '';
    for (let i = 0; i < n; i++) {
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`;
      const mo = { 1: 0, 'ป': 0, 'ล': 0, 'ข': 0 };
      for (const s of days) {
        if (s.off) { body += '<td class="offc"></td>'; continue; }
        const a = dayAmt(s, i, slot), ov = ovOf(slot, s.k, i);
        if (a) { mo[1] += a.pres; if (a.miss) mo[a.reason] += a.miss; }
        body += `<td class="gtd"><button class="pc att${amtCls(a)}${ov ? ' ovr' : ''}" data-act="aov" data-k="${s.k}" data-i="${i}" data-per="${s.per}" title="${ov ? 'แก้เฉพาะวิชานี้ (กดเพื่อเปลี่ยน)' : 'ตามเวลาเรียนรายห้อง — กดเพื่อแก้เฉพาะวิชานี้'}">${amtLabel(a) || '&nbsp;'}</button></td>`;
      }
      const st = attStat(slot, i, 'sum'), low = st.pct !== null && st.pct < PCT_MIN, c = (x) => `<td class="ro">${x || '-'}</td>`;
      body += c(mo[1]) + c(mo['ป']) + c(mo['ล']) + c(mo['ข']) + `<td class="ro${low ? ' gbad' : ''}">${fmtPct(st.pct)}</td></tr>`;
    }
    const dayN = days.filter(s => !s.off).length;
    return pick + info + `<div class="bar">${mseg}</div>
      <div class="card soft"><b>${esc(list[idx].name)}</b> · ${MONTHS[m]} ${monthBE(m)} · มีสอน ${dayN} วัน <span class="hint">· กดวันที่ที่หัวตารางเพื่อปิดวันที่ไม่มีสอนวิชานี้ (เช่น ครูติดราชการ) วันที่ปิดจะไม่นับคาบ</span> <button class="btn sm" data-act="aovclear" style="margin-left:8px">ล้างที่แก้เฉพาะวิชา (เดือนนี้)</button></div>
      <div class="tw"><table class="t sheet att"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>
      <div class="card row sp" style="margin-top:12px"><span class="hint">พิมพ์เป็นแบบฟอร์มรายเดือนของวิชานี้ และหน้าสรุปทั้งปี · เกณฑ์เวลาเรียนไม่น้อยกว่าร้อยละ ${PCT_MIN}</span><span class="row"><button class="btn" data-act="aprint">พิมพ์เดือนนี้</button><button class="btn primary" data-act="aprintall">พิมพ์ทุกเดือน + สรุปเทอม 1/2/รวม</button></span></div>`;
  }
  // ตารางสรุปรายวิชาบนหน้าจอ: คอลัมน์ = เดือนในภาคเรียน (หรือเทอม 1 / เทอม 2 เมื่อรวมทั้งปี) หน่วย = คาบ
  function vSubjSummary(slot, term, name) {
    const n = nStudents(), low = (p) => p !== null && p < PCT_MIN;
    const cols = term === 'sum' ? [[1, 'เทอม 1'], [2, 'เทอม 2']] : monthsOf(term).filter(m => sessionsOf(slot, 'sum').some(s => s.m === m)).map(m => [m, shortMon(m)]);
    const presOf = (i, c) => term === 'sum' ? attStat(slot, i, c).present : sessionsOf(slot, 'sum').filter(s => s.m === c).reduce((t, s) => { const a = dayAmt(s, i, slot); return t + (a ? a.pres : 0); }, 0);
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>' + cols.map(([, t]) => `<th>${esc(t)}<br><small>มา (คาบ)</small></th>`).join('') + '<th>รวมมา</th><th>ป่วย</th><th>ลา</th><th>ขาด</th><th>คาบที่เช็ค<br><small>จากตามตาราง</small></th><th>ร้อยละ</th>', body = '';
    for (let i = 0; i < n; i++) {
      const st = attStat(slot, i, term);
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>` + cols.map(([c]) => `<td class="num">${presOf(i, c) || '-'}</td>`).join('')
        + `<td class="num"><b>${st.present || '-'}</b></td><td class="num">${st.sick || '-'}</td><td class="num">${st.leave || '-'}</td><td class="num">${st.absent || '-'}</td><td class="num">${st.checked}/${st.plan}</td><td class="num${low(st.pct) ? ' warnc' : ''}"><b>${fmtPct(st.pct)}</b></td></tr>`;
    }
    return `<div class="card soft row sp"><span><b>${esc(name)} · สรุป ${termLabel(term)}</b> <span class="hint">ดูอย่างเดียว · หน่วยเป็นคาบ · ร้อยละ = คาบที่มา ÷ คาบที่เช็คแล้ว</span></span><span class="row"><button class="btn primary" data-act="aprint">พิมพ์สรุปนี้</button><button class="btn" data-act="aprintall">พิมพ์ทุกเดือน + สรุปทั้ง 3</button></span></div>
      <div class="tw"><table class="t sheet p5t"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  function subjOff(slot, k) {
    const a = subjAtt(slot), j = a.off.indexOf(k); if (j >= 0) a.off.splice(j, 1); else a.off.push(k);
    A().markDirty(); refresh();
  }

  // ---- หน้าพิมพ์ (ออกแบบเอง ไม่ใช้ผังชีท Excel)
  const subjName = (slot) => txt(SUBJ, 5 + slot, 3);
  function sMonthPage(slot, m) {
    const n = nStudents(), days = sessionsOf(slot, 'sum', true).filter(s => s.m === m);
    let h = '<tr><th rowspan="2" style="border:1px solid #000;width:26px">ที่</th><th rowspan="2" style="border:1px solid #000;width:190px">ชื่อ - สกุล</th>';
    for (const s of days) h += `<th style="border:1px solid #000;font-weight:400;width:22px;${s.off ? 'background:#d9d9d9' : ''}">${s.d}</th>`;
    h += '<th colspan="5" style="border:1px solid #000">สรุป (คาบ)</th></tr><tr>';
    for (const s of days) h += `<th style="border:1px solid #000;font-weight:400;font-size:12px;${s.off ? 'background:#d9d9d9' : ''}">${esc(DOW[dow(m, s.d)])}</th>`;
    h += ['มา', 'ป', 'ล', 'ข', '%สะสม'].map(x => `<th style="border:1px solid #000;width:${x === '%สะสม' ? 42 : 26}px">${x}</th>`).join('') + '</tr>';
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      rows += `<tr style="height:19px">${pcell(i + 1)}<td style="border:1px solid #000;padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>`;
      const mo = { 1: 0, 'ป': 0, 'ล': 0, 'ข': 0 };
      for (const s of days) { const a = s.off ? null : dayAmt(s, i, slot); if (a) { mo[1] += a.pres; if (a.miss) mo[a.reason] += a.miss; } rows += pcell(amtLabel(a), s.off ? 'background:#d9d9d9;' : ''); }
      const st = attStat(slot, i, 'sum');
      rows += [1, 'ป', 'ล', 'ข'].map(x => pcell(mo[x] || '-')).join('') + pcell(st.pct === null ? '-' : st.pct.toFixed(1)) + '</tr>';
    }
    const planned = days.filter(s => !s.off).reduce((t, s) => t + s.per, 0);
    const html = `<div style="${PF}font-size:15px">${attHead('บันทึกเวลาเรียนรายวิชา ' + subjName(slot) + ' ประจำเดือน' + MONTHS[m] + ' ' + monthBE(m))}
      <div style="margin:0 0 4px">วันที่มีสอนในเดือนนี้ <b>${days.filter(s => !s.off).length}</b> วัน (${planned} คาบ) &nbsp; ครูผู้สอน ${esc(txt(SUBJ, 5 + slot, 6) || '-')} &nbsp; ตัวเลข = จำนวนคาบที่เรียนในวัน · ข ขาด · ป ป่วย · ล ลา (ไม่ได้เรียนเลย) · ช่องสีเทา = ไม่มีสอน</div>
      <table style="border-collapse:collapse;font-size:14px;width:100%">${h}${rows}</table>
      <div style="display:flex;justify-content:space-around;margin-top:26px;font-size:16px;text-align:center">${sign('ครูผู้สอน', txt(SUBJ, 5 + slot, 6))}</div></div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  function sYearPage(slot, term) {
    term = term || 'sum';
    const n = nStudents(); const ms = MONTHS.map((_, m) => m).filter(m => inTerm(m, term) && sessionsOf(slot, 'sum').some(s => s.m === m));
    const perM = (m) => sessionsOf(slot, 'sum').filter(s => s.m === m).reduce((t, s) => t + s.per, 0);
    let h = `<tr><th rowspan="2" style="border:1px solid #000;width:26px">ที่</th><th rowspan="2" style="border:1px solid #000;width:180px">ชื่อ - สกุล</th><th colspan="${ms.length}" style="border:1px solid #000">คาบที่มาเรียน (ใต้ชื่อเดือน = คาบเรียนของเดือน)</th>${['มา', 'ป', 'ล', 'ข', 'คาบที่เช็ค', '%'].map(x => `<th rowspan="2" style="border:1px solid #000;width:36px">${x}</th>`).join('')}</tr><tr>`;
    h += ms.map(m => `<th style="border:1px solid #000;font-weight:400;font-size:12px;width:34px">${esc(MONTHS[m].slice(0, 3))}.<br>${perM(m)}</th>`).join('') + '</tr>';
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const st = attStat(slot, i, term); rows += `<tr style="height:19px">${pcell(i + 1)}<td style="border:1px solid #000;padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>`;
      for (const m of ms) { let x = 0; for (const s of sessionsOf(slot, 'sum').filter(q => q.m === m)) { const a = dayAmt(s, i, slot); if (a) x += a.pres; } rows += pcell(x || '-'); }
      rows += [st.present, st.sick, st.leave, st.absent, st.checked].map(x => pcell(x || '-')).join('') + pcell(fmtPct(st.pct), st.pct !== null && st.pct < PCT_MIN ? 'font-weight:700;' : '') + '</tr>';
    }
    const html = `<div style="${PF}font-size:15px">${attHead('สรุปเวลาเรียนรายวิชา ' + subjName(slot) + ' ' + termLabel(term) + ' ปีการศึกษา ' + yearBE())}<table style="border-collapse:collapse;font-size:14px;width:100%;margin-top:6px">${h}${rows}</table>
      <div style="margin-top:6px;font-size:14px">เกณฑ์เวลาเรียนไม่น้อยกว่าร้อยละ ${PCT_MIN} ของเวลาเรียนรายวิชา</div>
      <div style="display:flex;justify-content:space-around;margin-top:22px;font-size:16px;text-align:center">${sign('ครูผู้สอน', txt(SUBJ, 5 + slot, 6))}</div></div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  function sPages(slot, month) {
    if (!weekPeriods(slot) || !calTotal()) { A().toast('ยังไม่ได้ตั้งตารางสอนหรือปฏิทิน', 'warn'); return []; }
    const ms = MONTHS.map((_, m) => m).filter(m => sessionsOf(slot, 'sum', true).some(s => s.m === m));
    if (isTermKey(month)) return [sYearPage(slot, TERM_KEY[month])];
    if (month !== 'all') return [sMonthPage(slot, month)];
    return [...ms.map(m => sMonthPage(slot, m)), sYearPage(slot, 1), sYearPage(slot, 2), sYearPage(slot, 'sum')];
  }
  // ปพ.5 รายชั้น: เวลาเรียนร้อยละของนักเรียนทุกคนทุกวิชา
  function classAttPage(term) {
    const n = nStudents(), list = subjects().filter(s => weekPeriods(s.slot));
    const tl = term === 'sum' ? 'ทั้งปี' : 'ภาคเรียนที่ ' + term;
    const w = Math.max(34, Math.min(60, Math.floor(430 / Math.max(1, list.length))));
    let h = `<tr><th style="border:1px solid #000;width:26px">ที่</th><th style="border:1px solid #000;width:180px">ชื่อ - สกุล</th>${list.map(s => `<th style="border:1px solid #000;font-weight:400;font-size:12px;width:${w}px">${esc(subjName(s.slot))}</th>`).join('')}<th style="border:1px solid #000;width:46px">เฉลี่ย</th></tr>`;
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      let pt = 0, ct = 0; rows += `<tr style="height:19px">${pcell(i + 1)}<td style="border:1px solid #000;padding:0 4px;white-space:nowrap;text-align:left">${esc(stuName(i))}</td>`;
      for (const s of list) { const st = attStat(s.slot, i, term); pt += st.present; ct += st.checked; rows += pcell(fmtPct(st.pct), st.pct !== null && st.pct < PCT_MIN ? 'font-weight:700;text-decoration:underline;' : ''); }
      rows += pcell(ct ? (pt * 100 / ct).toFixed(2) : '-') + '</tr>';
    }
    const html = `<div style="${PF}font-size:15px">${attHead('สรุปเวลาเรียนรายวิชา (ร้อยละ) ' + tl)}<table style="border-collapse:collapse;font-size:14px;width:100%;margin-top:6px">${h}${rows}</table>
      <div style="margin-top:6px;font-size:14px">ตัวเลขคือร้อยละของคาบที่มาเรียนต่อคาบที่เช็คแล้ว · ตัวหนาขีดเส้นใต้ = ต่ำกว่าเกณฑ์ร้อยละ ${PCT_MIN}</div>
      <div style="display:flex;justify-content:space-around;margin-top:22px;font-size:16px;text-align:center">${sign('ครูประจำชั้น', realName(13))}</div></div>`;
    return { html, zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  }
  // ปพ.6 รายบุคคล: เวลาเรียนแยกรายวิชา
  function studentAttPage(i, term) {
    const list = subjects().filter(s => weekPeriods(s.slot)), tl = term === 'sum' ? 'ทั้งปี' : 'ภาคเรียนที่ ' + term;
    let rows = '';
    for (const s of list) {
      const st = attStat(s.slot, i, term), ok = st.pct === null ? '' : st.pct >= PCT_MIN ? 'ผ่านเกณฑ์' : 'ต่ำกว่าเกณฑ์';
      rows += `<tr style="height:22px"><td style="border:1px solid #000;padding:0 6px;text-align:left">${esc(txt(SUBJ, s.row, 2))}</td><td style="border:1px solid #000;padding:0 6px;text-align:left">${esc(subjName(s.slot))}</td>${[st.plan, st.checked, st.present, st.sick, st.leave, st.absent].map(x => pcell(x || '-')).join('')}${pcell(fmtPct(st.pct))}${pcell(ok)}</tr>`;
    }
    const html = `<div style="${PF}font-size:16px">${attHead('เวลาเรียนรายวิชา ' + tl)}<div style="margin:6px 0">ชื่อ-สกุล <b>${esc(stuName(i))}</b> &nbsp; เลขที่ ${i + 1}</div>
      <table style="border-collapse:collapse;font-size:15px;width:100%"><tr>${['รหัสวิชา', 'รายวิชา', 'คาบทั้งหมดตามตาราง', 'คาบที่เช็คแล้ว', 'มา', 'ป่วย', 'ลา', 'ขาด', 'ร้อยละ', 'ผล'].map(x => `<th style="border:1px solid #000;font-weight:700">${x}</th>`).join('')}</tr>${rows || '<tr><td colspan="10" style="border:1px solid #000">ยังไม่มีตารางสอน</td></tr>'}</table>
      <div style="margin-top:6px;font-size:14px">หน่วยเป็นคาบเรียน · เกณฑ์เวลาเรียนไม่น้อยกว่าร้อยละ ${PCT_MIN} ของเวลาเรียนรายวิชา</div>
      <div style="display:flex;justify-content:space-around;margin-top:26px;font-size:16px;text-align:center">${sign('ครูประจำชั้น', realName(13))}</div></div>`;
    return { html, zoom: 1, land: false, m: [0.6, 0.6, 0.6, 0.6], fit: true, fitW: true };
  }
  // ---- ตารางสรุปบนหน้าเว็บ (ปพ.5 / ปพ.6)
  function p5AttBlock() {
    const n = nStudents(), list = subjects().filter(s => weekPeriods(s.slot)), term = U.term;
    const seg2 = `<div class="bar">${seg([['class', 'รายชั้น (ทุกวิชา)'], ['subj', 'รายวิชา']], U.p5att || 'class', 'p5att')}</div>`;
    if (!list.length) return `<div class="card"><h3>เวลาเรียนรายวิชา</h3><p class="hint">ยังไม่ได้ตั้งตารางสอน — ไปที่ ข้อมูล › ปฏิทิน · เวลาเรียน › ตารางสอน</p><button class="btn" data-act="go" data-tab="data" data-sub="time">ไปตั้งตารางสอน</button></div>`;
    if ((U.p5att || 'class') === 'class') {
      let body = '';
      for (let i = 0; i < n; i++) {
        let pt = 0, ct = 0; body += `<tr><td class="stk">${i + 1}. ${esc(stuName(i))}</td>`;
        for (const s of list) { const st = attStat(s.slot, i, term); pt += st.present; ct += st.checked; body += `<td class="num${st.pct !== null && st.pct < PCT_MIN ? ' warnc' : ''}">${fmtPct(st.pct)}</td>`; }
        body += `<td class="num"><b>${ct ? (pt * 100 / ct).toFixed(2) : '-'}</b></td></tr>`;
      }
      return `<div class="card"><div class="row sp"><h3>เวลาเรียนรายวิชา — รายชั้น (ร้อยละ)</h3><button class="btn" data-act="p5attprint" data-k="class">พิมพ์รายชั้น</button></div>${seg2}
        <div class="tw p5tw"><table class="t p5t"><thead><tr><th class="stk">นักเรียน</th>${list.map(s => `<th>${esc(s.name)}</th>`).join('')}<th>เฉลี่ย</th></tr></thead><tbody>${body}</tbody></table></div><p class="hint">สีส้ม = ต่ำกว่าเกณฑ์ร้อยละ ${PCT_MIN}</p></div>`;
    }
    if (!list.some(s => s.slot === U.p5subj)) U.p5subj = list[0].slot;
    const slot = U.p5subj; let body = '';
    for (let i = 0; i < n; i++) { const st = attStat(slot, i, term); body += `<tr><td class="stk">${i + 1}. ${esc(stuName(i))}</td>${[st.plan, st.checked, st.present, st.sick, st.leave, st.absent].map(x => `<td class="num">${x || '-'}</td>`).join('')}<td class="num${st.pct !== null && st.pct < PCT_MIN ? ' warnc' : ''}"><b>${fmtPct(st.pct)}</b></td></tr>`; }
    return `<div class="card"><div class="row sp"><h3>เวลาเรียนรายวิชา — แยกรายวิชา</h3><button class="btn" data-act="p5attprint" data-k="subj">พิมพ์วิชานี้ (ทุกเดือน + สรุป)</button></div>${seg2}
      <div class="pick" style="margin-bottom:8px"><select data-f="p5subj" aria-label="เลือกวิชา">${list.map(s => `<option value="${s.slot}"${s.slot === slot ? ' selected' : ''}>${esc(s.name)}</option>`).join('')}</select></div>
      <div class="tw p5tw"><table class="t p5t"><thead><tr><th class="stk">นักเรียน</th><th>คาบตามตาราง</th><th>เช็คแล้ว</th><th>มา</th><th>ป่วย</th><th>ลา</th><th>ขาด</th><th>ร้อยละ</th></tr></thead><tbody>${body}</tbody></table></div></div>`;
  }
  function p6AttBlock(i) {
    const list = subjects().filter(s => weekPeriods(s.slot)), term = U.p6term;
    if (!list.length) return '';
    const rows = list.map(s => { const st = attStat(s.slot, i, term); return `<tr><td class="stk" style="max-width:none">${esc(s.name)}</td>${[st.checked, st.present, st.sick, st.leave, st.absent].map(x => `<td class="num">${x || '-'}</td>`).join('')}<td class="num${st.pct !== null && st.pct < PCT_MIN ? ' warnc' : ''}"><b>${fmtPct(st.pct)}</b></td></tr>`; }).join('');
    return `<div class="card" style="margin-top:12px"><div class="row sp"><h3>เวลาเรียนรายวิชา</h3><button class="btn" data-act="p6attprint">พิมพ์เวลาเรียนคนนี้</button></div><div class="tw p5tw"><table class="t p5t"><thead><tr><th class="stk">รายวิชา</th><th>คาบที่เช็คแล้ว</th><th>มา</th><th>ป่วย</th><th>ลา</th><th>ขาด</th><th>ร้อยละ</th></tr></thead><tbody>${rows}</tbody></table></div></div>`;
  }


  function vBackup() {
    return `<div class="card"><h3>สำรองข้อมูล</h3><p class="hint">ดาวน์โหลดข้อมูลทุกห้องของโรงเรียนเป็นไฟล์ .json เก็บไว้ในเครื่อง และนำเข้ากลับได้เมื่อต้องการ</p>
      <div class="row"><button class="btn primary" data-act="backup">สำรองข้อมูล (ดาวน์โหลด)</button><button class="btn" data-act="restore">นำเข้าไฟล์สำรอง</button><button class="btn" data-act="rooms">จัดการห้องเรียน</button></div></div>`;
  }

  // ---------------------------------------------------------------- คะแนนรายวิชา
  // ผังชีท "วิชาที่ N": ตัวชี้วัด = ช่องคะแนนระหว่างเรียน (แถว 5 = ชื่อ แถว 6 = คะแนนเต็ม) · สอบปลายภาค · ผลลัพธ์
  const SC = {
    1: { ind: [C('C'), C('AK')], work: [C('CA'), C('CY')], exam: C('DA'), btw: C('CZ'), total: C('DB'), pct: C('DC'), grade: C('DD') },
    2: { ind: [C('AO'), C('BW')], work: [C('DE'), C('EB')], exam: C('EE'), btw: C('ED'), total: C('EF'), pct: C('EG'), grade: C('EH') },
  };
  const IND_ROW = 7, SC_NAME = 5, SC_MAX = 6;
  function scCfg() {
    const T = SC[U.term]; if (!T) return null;
    const sh = 'วิชาที่ ' + U.subj;
    let nw = 0;   // จำนวนตัวชี้วัด (ช่องคะแนนระหว่างเรียน) ที่กำหนดคะแนนเต็มแล้ว
    for (let c = T.work[0]; c <= T.work[1]; c++) if ((num(sh, SC_MAX, c) || 0) > 0) nw = c - T.work[0] + 1;
    return { sh, T, nw };
  }
  function workAdd(d) {
    const k = scCfg(); if (!k) return; const max = k.T.work[1] - k.T.work[0] + 1;
    if (d > 0 && k.nw < max) { const c = k.T.work[0] + k.nw; setv(k.sh, SC_NAME, c, 'ตัวชี้วัดที่ ' + (k.nw + 1)); setv(k.sh, SC_MAX, c, 10); }
    else if (d < 0 && k.nw > 0) { const c = k.T.work[0] + k.nw - 1; setv(k.sh, SC_NAME, c, null); setv(k.sh, SC_MAX, c, null); }
    refresh();
  }
  async function clearScores() {
    const k = scCfg(); if (!k) return;
    const yes = await A().modal('ล้างคะแนนภาคเรียนนี้', `<p>ล้างคะแนนตัวชี้วัดและคะแนนสอบของ <b>${esc(txt(SUBJ, 5 + U.subj, 3))}</b> ภาคเรียนที่ ${U.term} ของนักเรียนทุกคน ย้อนกลับไม่ได้</p>`, [{ label: 'ยกเลิก', value: false }, { label: 'ล้างคะแนน', value: true, primary: true }]);
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
    const indIn = Array.from({ length: k.nw }, (_, j) => `<div class="indset"><b>ตัวชี้วัดที่ ${j + 1}</b>${field('ชื่อ', `<input ${bindF(sh, SC_NAME, T.work[0] + j)} value="${esc(txt(sh, SC_NAME, T.work[0] + j))}">`)}${field('คะแนนเต็ม', `<input inputmode="numeric" ${bindF(sh, SC_MAX, T.work[0] + j, 'num')} value="${esc(txt(sh, SC_MAX, T.work[0] + j))}">`)}</div>`).join('');
    const none = !k.nw;
    const cfg = `<details class="card soft more" ${none ? 'open' : ''}><summary><b>ตั้งค่าตัวชี้วัดและคะแนนเต็ม</b> <span class="hint">รายการตัวชี้วัดแยกตามภาคเรียน · สัดส่วนคะแนนใช้ร่วมกันทั้ง 2 ภาคเรียนของวิชานี้</span></summary>
      <div class="maxbox" style="margin-top:8px">
        <label class="fld"><span>สัดส่วน ระหว่างเรียน</span><input inputmode="numeric" ${bindF(sh, 3, 'DA', 'num')} value="${esc(txt(sh, 3, 'DA'))}"></label>
        <label class="fld"><span>สัดส่วน ปลายภาค (= คะแนนสอบเต็ม)</span><input inputmode="numeric" ${bindF(sh, 3, 'DC', 'num')} value="${esc(txt(sh, 3, 'DC'))}"></label>
      </div>
      <p class="hint">ระหว่างเรียน : ปลายภาค = ${fmt2(ratio[0])} : ${fmt2(ratio[1])} (รวม ${fmt2((ratio[0] || 0) + (ratio[1] || 0))} คะแนน) คะแนนตัวชี้วัดทุกตัวจะถูกปรับให้เต็มตามสัดส่วนระหว่างเรียนอัตโนมัติ</p>
      <h4 style="margin:.6rem 0 .3rem">ตัวชี้วัด / คะแนนระหว่างเรียน เทอม ${U.term}</h4><div class="maxbox">${indIn}</div>
      <div class="row" style="margin:6px 0"><button class="btn sm" data-act="work-add">＋ เพิ่มตัวชี้วัด</button><button class="btn sm" data-act="work-del" ${k.nw ? '' : 'disabled'}>－ ลดตัวชี้วัด</button></div></details>`;
    if (none) return cfg + '<div class="card empty"><b>ยังไม่ได้กำหนดตัวชี้วัด</b>กดปุ่ม “เพิ่มตัวชี้วัด” ด้านบน หรือ<br><br><button class="btn primary" data-act="work-add">เพิ่มตัวชี้วัดแรก</button></div>';
    let th = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>', body = '';
    for (let j = 0; j < k.nw; j++) th += `<th>${esc(txt(sh, SC_NAME, T.work[0] + j) || 'ตัวชี้วัด')}<br><small>(${fmt2(num(sh, SC_MAX, T.work[0] + j))})</small></th>`;
    th += `<th>สอบ<br><small>(${esc(txt(sh, 3, 'DC'))})</small></th><th class="ro">ระหว่างเรียน</th><th class="ro">รวม</th><th class="ro">ร้อยละ</th><th class="ro">ระดับ</th>`;
    for (let i = 0; i < n; i++) {
      const r = IND_ROW + i; let gc = 0;
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`;
      for (let j = 0; j < k.nw; j++) body += gcell(sh, r, T.work[0] + j, 'num', i, gc++);
      body += gcell(sh, r, T.exam, 'num', i, gc++) + rocell(sh, r, T.btw) + rocell(sh, r, T.total) + rocell(sh, r, T.pct) + rocell(sh, r, T.grade, 'grade') + '</tr>';
    }
    return cfg + `<p class="hint">คลิกช่องแล้วพิมพ์คะแนน · ลูกศร/Enter เลื่อนช่อง · วางคะแนนหลายช่องจาก Excel ได้ (Ctrl+V)</p>
      <div class="tw"><table class="t sheet"><thead><tr>${th}</tr></thead><tbody>${body}</tbody></table></div>
      <p style="margin-top:10px"><button class="btn danger sm" data-act="clearall">ล้างคะแนนภาคเรียนนี้</button> <button class="btn sm" data-act="docx" data-name="${esc(sh)}">เปิดชีทเต็มรูปแบบ (Excel)</button></p>`;
  }
  function vScoreYear(n) {
    // คอลัมน์ผลรายภาคของชีท: EI-EL = เทอม 1 (ระหว่างภาค ปลายภาค รวม ร้อยละ) · EM-EP = เทอม 2 · EQ ER ES = รวมรายปี ร้อยละ ระดับ
    const sh = 'วิชาที่ ' + U.subj; let body = '';
    const cell = (r, c, kind) => rocell(sh, r, c, kind);
    for (let i = 0; i < n; i++) {
      const r = IND_ROW + i;
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`
        + cell(r, 'EI') + cell(r, 'EJ') + cell(r, 'EK') + cell(r, 'EL') + cell(r, SC[1].grade, 'grade')
        + cell(r, 'EM') + cell(r, 'EN') + cell(r, 'EO') + cell(r, 'EP') + cell(r, SC[2].grade, 'grade')
        + `<td class="ro" data-sum="${r}|EI|EM"></td><td class="ro" data-sum="${r}|EJ|EN"></td>` + cell(r, 'EQ') + cell(r, 'ER') + cell(r, 'ES', 'grade') + '</tr>';
    }
    return `<p class="hint">สรุปผลทั้งปี คำนวณจากคะแนนภาคเรียนที่ 1 และ 2 (ดูอย่างเดียว) · คะแนนเก็บ = คะแนนตัวชี้วัด/ระหว่างเรียน</p>
      <div class="tw"><table class="t sheet"><thead>
        <tr><th class="no" rowspan="2">ที่</th><th class="stk" rowspan="2">ชื่อ - สกุล</th><th colspan="5">ภาคเรียนที่ 1</th><th colspan="5">ภาคเรียนที่ 2</th><th colspan="5">รวมทั้งปี</th></tr>
        <tr><th>คะแนนเก็บ</th><th>คะแนนสอบ</th><th>รวม</th><th>ร้อยละ</th><th>ระดับ</th><th>คะแนนเก็บ</th><th>คะแนนสอบ</th><th>รวม</th><th>ร้อยละ</th><th>ระดับ</th><th>คะแนนเก็บ</th><th>คะแนนสอบ</th><th>รวม</th><th>ร้อยละ</th><th>ระดับผลการเรียน</th></tr>
      </thead><tbody>${body}</tbody></table></div>`;
  }

  // ---------------------------------------------------------------- ประเมิน (คุณลักษณะ · ค่านิยม · กิจกรรม)
  const LV_TRAIT = ['ไม่ผ่าน', 'ผ่าน', 'ดี', 'ดีเยี่ยม'], LV_COMP = ['ปรับปรุง', 'พอใช้', 'ดี', 'ดีเยี่ยม'];
  const AS = {
    trait: { title: 'คุณลักษณะอันพึงประสงค์', base: 'คุณลักษณะ', lv: [LV_TRAIT, LV_TRAIT] },
    value: { title: 'ค่านิยม · สมรรถนะ', base: 'ค่านิยม', lv: [LV_TRAIT, LV_COMP] },
    act: { title: 'กิจกรรมพัฒนาผู้เรียน', base: 'กิจกรรม' },
  };
  // หัวข้อของ คุณลักษณะ/ค่านิยม ปรับได้ (เพิ่ม/ลด/ตั้งชื่อ) — ตำแหน่งคอลัมน์คำนวณจากจำนวนหัวข้อของห้องนี้
  const itemCfg = (kind) => A().S.items[kind];
  const items = (kind, g) => root.PPAssess.layout(kind, itemCfg(kind).n).items[g];
  const groupName = (kind, g) => itemCfg(kind).g[g];
  const assessGroups = (kind) => [0, 1].map(g => { const cols = items(kind, g); return { name: groupName(kind, g), cols: [cols[0], cols[cols.length - 1]], lv: AS[kind].lv[g], hr: 3, g }; });
  const assessSheet = () => AS[U.akind].base + ' ' + (U.aterm === 'sum' ? 'รวม' : U.aterm);
  function vAssess() {
    const n = nStudents();
    const head = `<div class="bar">${seg([['trait', 'คุณลักษณะ'], ['value', 'ค่านิยม · สมรรถนะ'], ['act', 'กิจกรรมพัฒนาผู้เรียน']], U.akind, 'akind')}${seg([[1, 'เทอม 1'], [2, 'เทอม 2'], ['sum', 'ทั้งปี']], U.aterm, 'aterm')}<span class="grow"></span>${U.akind === 'act' ? '' : `<button class="btn${U.aedit ? ' on' : ''}" data-act="aedit">${U.aedit ? 'ปิดการปรับหัวข้อ' : 'ปรับหัวข้อ'}</button>`}<button class="btn primary" data-act="print-doc" data-name="${esc(assessSheet())}">พิมพ์${esc(AS[U.akind].title)}</button></div>`;
    if (!n) return head + '<div class="card empty"><b>ยังไม่มีรายชื่อนักเรียน</b>กรอกที่แท็บ “ข้อมูล › นักเรียน” ก่อน</div>';
    const edit = U.aterm !== 'sum', sh = assessSheet();
    if (U.akind === 'act') return head + vActivities(n, sh, edit);
    const K = { groups: assessGroups(U.akind) }; let h1 = '<th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th>', h2 = '<th class="no"></th><th class="stk"></th>', cols = [];
    for (const g of K.groups) {
      h1 += `<th colspan="${g.cols[1] - g.cols[0] + 1}">${esc(g.name)}</th>`;
      for (let c = g.cols[0]; c <= g.cols[1]; c++) { h2 += `<th data-ah="${g.g}|${c - g.cols[0]}">${esc(txt(sh, g.hr, c) || String(c - g.cols[0] + 1))}</th>`; cols.push({ c, lv: g.lv }); }
    }
    let body = '';
    for (let i = 0; i < n; i++) {
      body += `<tr><td class="no">${i + 1}</td><td class="stk">${esc(stuName(i))}</td>`;
      for (const x of cols) body += lvCell(sh, 5 + i, x.c, x.lv, edit);
      body += '</tr>';
    }
    const tools = edit ? `<div class="bar"><span class="hint">ตั้งทั้งห้อง:</span><button class="btn sm" data-act="setall" data-k="1">ผ่าน</button><button class="btn sm" data-act="setall" data-k="2">ดี</button><button class="btn sm" data-act="setall" data-k="3">ดีเยี่ยม</button><button class="btn sm danger" data-act="setall" data-k="-1">ล้างทั้งหมด</button></div>` : '';
    return head + (U.aedit ? itemEditor(U.akind) : '') + `<p class="hint">แตะช่องเพื่อเปลี่ยนระดับ (ว่าง → ผ่าน → ดี → ดีเยี่ยม → ไม่ผ่าน → ว่าง)${edit ? '' : ' · ทั้งปีคำนวณจากเทอม 1 และ 2 ดูอย่างเดียว'}</p>${tools}<div class="tw"><table class="t sheet"><thead><tr>${h1}</tr><tr>${h2}</tr></thead><tbody>${body}</tbody></table></div>`;
  }
  // แผงปรับหัวข้อ: เปลี่ยนชื่อกลุ่ม/หัวข้อ เพิ่ม ลบ — ใช้กับทั้งเทอม 1 เทอม 2 ทั้งปี และหน้าพิมพ์ ปพ.5/ปพ.6
  function itemEditor(kind) {
    const c = itemCfg(kind), K = root.PPAssess.KINDS[kind];
    const grp = [0, 1].map(g => `<div class="agrp"><label class="fld"><span>ชื่อกลุ่ม</span><input class="in" data-agrp="${g}" value="${esc(c.g[g])}"></label>
      <div class="aitems">${c.names[g].map((nm, i) => `<span class="aitem"><span class="no">${i + 1}</span><input class="in" data-aitem="${g}|${i}" value="${esc(nm)}" aria-label="ชื่อหัวข้อที่ ${i + 1}"><button class="x" data-act="adel" data-g="${g}" data-i="${i}" title="ลบหัวข้อนี้"${c.n[g] <= 1 ? ' disabled' : ''}>✕</button></span>`).join('')}
        <button class="btn sm" data-act="aadd" data-g="${g}"${c.n[g] >= K.max[g] ? ' disabled' : ''}>+ เพิ่มหัวข้อ</button></div></div>`).join('');
    return `<div class="card aedit"><h3>ปรับหัวข้อ · ${esc(AS[kind].title)}</h3>
      <p class="hint">ตั้งชื่อกลุ่มและหัวข้อ เพิ่มหรือลบได้ — มีผลทั้งเทอม 1 เทอม 2 ทั้งปี และเอกสาร ปพ.5 / ปพ.6 · ลบหัวข้อจะลบระดับที่กรอกไว้ของหัวข้อนั้นด้วย · ชื่อสั้น ๆ จะพอดีหัวตารางพิมพ์</p>${grp}</div>`;
  }
  function itemChange(x) {
    const kind = U.akind;
    if (x.dataset.agrp !== undefined) { const g = +x.dataset.agrp, v = x.value.trim() || root.PPAssess.KINDS[kind].origG[g]; A().editItems(kind, c => { c.g[g] = v; }); refresh(); return; }
    const [g, i] = x.dataset.aitem.split('|').map(Number), v = x.value.trim() || String(i + 1);
    A().renameItem(kind, g, i, v); x.value = v;
    const th = $(`th[data-ah="${g}|${i}"]`); if (th) th.textContent = v;
  }
  function itemAdd(g) { A().editItems(U.akind, c => { c.n[g]++; c.names[g].push(String(c.n[g])); }); refresh(); }
  function itemDel(g, i) {
    const nm = itemCfg(U.akind).names[g][i];
    if (!confirm('ลบหัวข้อ “' + nm + '” ?\nระดับที่กรอกไว้ของหัวข้อนี้จะถูกลบด้วย')) return;
    A().editItems(U.akind, c => { c.n[g]--; c.names[g].splice(i, 1); }, { g, i }); refresh();
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
    for (const g of assessGroups(U.akind)) for (let i = 0; i < n; i++) for (let c = g.cols[0]; c <= g.cols[1]; c++) setv(sh, 5 + i, c, level === -1 ? null : level);
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
    return `<p class="hint">${edit ? 'แตะช่องเพื่อสลับ ว่าง → ผ่าน → ไม่ผ่าน → ว่าง' : 'ทั้งปีคำนวณจากเทอม 1 และ 2 · ผ่านเมื่อผ่านทั้งสองภาคเรียน · ดูอย่างเดียว'}</p>${tools}<div class="tw"><table class="t sheet"><thead><tr><th class="no">ที่</th><th class="stk">ชื่อ - สกุล</th><th>${esc(txt(sh, 2, 3) || 'แนะแนว')}</th><th>${esc(txt(sh, 2, 5) || 'ลูกเสือ-เนตรนารี')}</th><th>ชื่อชุมนุม</th><th>${esc(txt(sh, 2, 7) || 'ชุมนุม')}</th><th>${esc((txt(sh, 2, 10) || 'กิจกรรมเพื่อสังคมฯ').replace(/\n/g, ' '))}</th></tr></thead><tbody>${body}</tbody></table></div>`;
  }

  // ---------------------------------------------------------------- เอกสาร: คุณลักษณะ · ค่านิยม · กิจกรรม · น้ำหนัก-ส่วนสูง (ครูประจำชั้นเซ็น)
  // doc ชื่อเสมือน 'ประเมิน:trait|value|act|body' → 3 หน้า เทอม 1 · เทอม 2 · รวม (น้ำหนัก-ส่วนสูงหน้าเดียว)
  const ASSESS_DOCS = { 'ประเมิน:trait': 'คุณลักษณะอันพึงประสงค์', 'ประเมิน:value': 'ค่านิยม · สมรรถนะ', 'ประเมิน:act': 'กิจกรรมพัฒนาผู้เรียน', 'ประเมิน:body': 'น้ำหนัก-ส่วนสูง', 'สรุป:score': 'คะแนนรายชั้น', 'สรุป:pct': 'คะแนนร้อยละ', 'สรุป:grade': 'เกรดรายชั้น', 'สรุป:dist': 'สรุปผลเกรดรายวิชา (จำนวน/ร้อยละ)' };
  const homeSign = () => `<div style="${PF}display:flex;justify-content:space-around;margin-top:16px;font-size:16px;text-align:center">${sign('ครูประจำชั้น', realName(13))}</div>`;
  const PG = { zoom: 1, land: false, m: [0.5, 0.5, 0.5, 0.5], fit: true, fitW: true };
  const tdN = (x, o) => `<td style="${SB}padding:0 4px;white-space:nowrap;text-align:left;${o || ''}">${x}</td>`;
  const thS = (x, o) => `<th style="${SB}font-weight:400;${o || ''}">${x}</th>`;
  function levelOf(sh, r, cols, labels) { for (let k = 0; k < cols.length; k++) if (txt(sh, r, cols[k]).trim() === '/') return k; return -1; }
  function assessPage(kind, term) {
    const n = nStudents(), sh = AS[kind].base + ' ' + (term === 'sum' ? 'รวม' : term), cfg = itemCfg(kind), L = root.PPAssess.layout(kind, cfg.n);
    const labels = AS[kind].lv, short = ['0', '1', '2', '3'];
    let h1 = `<tr>${thS('ที่', 'width:26px')}${thS('ชื่อ - สกุล', 'width:190px')}`, h2 = '';
    for (let g = 0; g < 2; g++) h1 += `<th colspan="${cfg.n[g] + 1}" style="${SB}">${esc(cfg.g[g])}</th>`;
    h1 += '</tr><tr><th style="' + SB + '"></th><th style="' + SB + '"></th>';
    for (let g = 0; g < 2; g++) { h2 += L.items[g].map((c, i) => thS(esc(txt(sh, 3, c) || String(i + 1)), 'font-size:12px;min-width:20px;padding:0 1px')).join('') + thS('ผลการ<br>ประเมิน', 'font-size:12px;padding:0 3px'); }
    let rows = ''; const cnt = [[0, 0, 0, 0], [0, 0, 0, 0]];
    for (let i = 0; i < Math.max(n, 1); i++) {
      const r = 5 + i; rows += `<tr style="height:19px">${pcell(i + 1)}${tdN(esc(stuName(i)))}`;
      for (let g = 0; g < 2; g++) {
        rows += L.items[g].map(c => { const v = get(sh, r, c); return pcell(typeof v === 'number' ? v : (v ? esc(String(v)) : ''), 'font-size:13px'); }).join('');
        const lv = levelOf(sh, r, L.res[g]); if (lv >= 0 && n) cnt[g][lv]++;
        rows += pcell(lv >= 0 ? labels[g][lv] : '', 'font-size:12px;font-weight:700;white-space:nowrap');
      }
      rows += '</tr>';
    }
    const sum = [0, 1].map(g => `<td style="${SB}padding:2px 6px;vertical-align:top"><b>${esc(cfg.g[g])}</b><br>${labels[g].map((l, k) => `${l} ${cnt[g][k]} คน`).join(' · ')}</td>`).join('');
    const html = `<div style="${PF}font-size:15px">${attHead('สรุปการประเมิน' + AS[kind].title + ' ' + (term === 'sum' ? 'รวมทั้งปี' : 'ภาคเรียนที่ ' + term))}
      <table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center">${h1}${h2}</tr>${rows}</table>
      <div style="margin:6px 0 4px;font-size:14px">เกณฑ์: 3 = ${labels[0][3]} · 2 = ${labels[0][2]} · 1 = ${labels[0][1]} · 0 = ${labels[0][0]}${kind === 'value' ? ` (สมรรถนะ: 3 = ${labels[1][3]} · 2 = ${labels[1][2]} · 1 = ${labels[1][1]} · 0 = ${labels[1][0]})` : ''}</div>
      <table style="border-collapse:collapse;font-size:14px;width:100%"><tr>${sum}</tr></table>${homeSign()}</div>`;
    return { ...PG, html };
  }
  function actPage(term) {
    const n = nStudents(), sh = 'กิจกรรม ' + (term === 'sum' ? 'รวม' : term);
    const heads = [txt(sh, 2, 3) || 'แนะแนว', txt(sh, 2, 5) || 'ลูกเสือ-เนตรนารี', 'ชื่อชุมนุม', txt(sh, 2, 7) || 'ชุมนุม', (txt(sh, 2, 10) || 'กิจกรรมเพื่อสังคมและสาธารณประโยชน์').replace(/\n/g, '')];
    const cnt = [0, 1, 2, 3].map(() => [0, 0]); let rows = '';
    const res = (s) => s === 'pass' ? 'ผ่าน' : s === 'fail' ? '<b>ไม่ผ่าน</b>' : '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const r = 5 + i, st = ACT.map(pr => actState(sh, r, pr));
      st.forEach((s, k) => { if (n && s === 'pass') cnt[k][0]++; else if (n && s === 'fail') cnt[k][1]++; });
      rows += `<tr style="height:19px">${pcell(i + 1)}${tdN(esc(stuName(i)))}${pcell(res(st[0]))}${pcell(res(st[1]))}${pcell(esc(txt(sh, r, 7)))}${pcell(res(st[2]))}${pcell(res(st[3]))}</tr>`;
    }
    const names = ['แนะแนว', 'ลูกเสือ-เนตรนารี', 'ชุมนุม', 'กิจกรรมเพื่อสังคมฯ'];
    const html = `<div style="${PF}font-size:15px">${attHead('สรุปการประเมินกิจกรรมพัฒนาผู้เรียน ' + (term === 'sum' ? 'รวมทั้งปี' : 'ภาคเรียนที่ ' + term))}
      <table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center"><tr>${thS('ที่', 'width:26px')}${thS('ชื่อ - สกุล', 'width:190px')}${heads.map(x => thS(esc(x))).join('')}</tr>${rows}</table>
      <table style="border-collapse:collapse;font-size:14px;width:100%;margin-top:8px"><tr>${names.map((x, k) => `<td style="${SB}padding:2px 6px"><b>${x}</b><br>ผ่าน ${cnt[k][0]} · ไม่ผ่าน ${cnt[k][1]}</td>`).join('')}</tr></table>${term === 'sum' ? '<div style="margin-top:4px;font-size:14px">ผ่านทั้งปี = ผ่านทั้ง 2 ภาคเรียน</div>' : ''}${homeSign()}</div>`;
    return { ...PG, html };
  }
  function bodyPage() {
    const n = nStudents(), BS = 'น้ำหนัก-ส่วนสูง';
    const dts = [3, 5, 7, 9].map((c, k) => `<th colspan="2" style="${SB}font-weight:400">ครั้งที่ ${k + 1}<br><small>${esc(rawOf(get(BS, 6, c), 'date') || '-')}</small></th>`).join('');
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) rows += `<tr style="height:19px">${pcell(i + 1)}${tdN(esc(stuName(i)))}${[3, 4, 5, 6, 7, 8, 9, 10].map(c => pcell(sv(get(BS, 7 + i, c)))).join('')}</tr>`;
    const html = `<div style="${PF}font-size:15px">${attHead('บันทึกน้ำหนัก-ส่วนสูงของนักเรียน')}<table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center"><tr>${thS('ที่', 'width:26px')}${thS('ชื่อ - สกุล', 'width:190px')}${dts}</tr><tr><th style="${SB}"></th><th style="${SB}"></th>${thS('น้ำหนัก<br><small>(กก.)</small>').repeat(4).replace(/<\/th><th/g, '</th><th')}</tr>${rows}</table>${homeSign()}</div>`;
    return { ...PG, html };
  }
  // รายชื่อนักเรียน (เอกสาร › สรุปผลรายชั้น › ชื่อนักเรียน): ดึงจาก ข้อมูล › นักเรียน
  const TH_MON = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  // เลขวันที่ Excel (ปี พ.ศ. อยู่ในช่องปี) → '10 มีนาคม 2569'
  const thLong = (v) => { if (typeof v !== 'number') return v ? String(v) : ''; const d = new Date(Date.UTC(1899, 11, 30) + v * 864e5); return d.getUTCDate() + ' ' + TH_MON[d.getUTCMonth()] + ' ' + d.getUTCFullYear(); };
  function rosterPage() {
    const n = nStudents(); let m = 0, f = 0, rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const sex = txt(ST, 3 + i, 7).trim(); if (n) { if (sex === 'ช') m++; else if (sex === 'ญ') f++; }
      rows += `<tr style="height:21px">${pcell(i + 1)}${pcell(esc(sv(get(ST, 3 + i, 6))))}${tdN(esc(stuName(i)))}${pcell(esc(thLong(get(ST, 3 + i, 13))), 'white-space:nowrap')}${pcell('')}</tr>`;
    }
    const html = `<div style="${PF}font-size:16px">${attHead('รายชื่อนักเรียน')}
      <div style="display:flex;justify-content:space-between;margin:0 2px 4px;font-size:15px"><span>ครูประจำชั้น ${esc(realName(13) || '-')}</span><span>จำนวนนักเรียน <b>${n}</b> คน${m + f ? ` (ชาย ${m} · หญิง ${f})` : ''}</span></div>
      <table style="border-collapse:collapse;font-size:15px;width:100%;text-align:center"><tr>${thS('เลขที่', 'width:46px')}${thS('เลขประจำตัว<br>นักเรียน', 'width:90px')}${thS('ชื่อ - สกุล')}${thS('วัน/เดือน/ปีเกิด', 'width:150px')}${thS('หมายเหตุ', 'width:110px')}</tr>${rows}</table>${homeSign()}</div>`;
    return { ...PG, html };
  }
  // ---- สรุปผลรายชั้น (ไม่ใช้สี) อ่านค่าจากชีท คะแนนรายชั้น / คะแนนร้อยละ / เกรดรายชั้น / เกรดร้อยละ ซึ่งคำนวณไว้แล้ว
  const SUMDOC = { score: ['คะแนนรายชั้น', 'คะแนนผลการเรียน', true], pct: ['คะแนนร้อยละ', 'คะแนนผลการเรียนร้อยละ', true], grade: ['เกรดรายชั้น', 'ระดับผลการเรียน', false] };
  const subjCols = (sh) => { const out = []; for (let c = 3; c <= 22; c++) { const v = get(sh, 4, c); if (v !== null && v !== '' && !(v && v.code)) out.push(c); } return out; };
  const rotTh = (x) => `<th style="${SB}height:118px;padding:2px 0"><div style="writing-mode:vertical-rl;transform:rotate(180deg);margin:0 auto;white-space:nowrap;font-weight:400;font-size:13px">${esc(x)}</div></th>`;
  function sumClassPage(kind, term) {
    const [pre, title, hasFoot] = SUMDOC[kind], sh = pre + ' ' + (term === 'sum' ? 'รวม' : term), n = nStudents(), cs = subjCols(sh);
    const tl = term === 'sum' ? 'รวมทั้งปี' : 'ภาคเรียนที่ ' + term, isG = kind === 'grade';
    let h = `<tr>${thS('ที่', 'width:26px')}${thS('ชื่อ - สกุล', 'width:190px')}${cs.map(c => rotTh(txt(sh, 4, c))).join('')}${isG ? thS('ผลการเรียน<br>เฉลี่ย') + thS('ลำดับที่') : thS('รวม') + thS('ร้อยละ') + thS('ลำดับที่')}</tr>`;
    if (!isG) h += `<tr style="height:19px"><td style="${SB}"></td>${tdN('คะแนนเต็ม')}${cs.map(c => pcell(sv(get(sh, 5, c)))).join('')}${pcell(sv(get(sh, 5, 23)))}${pcell(sv(get(sh, 5, 24)))}${pcell('')}</tr>`;
    let rows = '';
    for (let i = 0; i < Math.max(n, 1); i++) {
      const r = 6 + i;
      rows += `<tr style="height:19px">${pcell(i + 1)}${tdN(esc(stuName(i)))}${cs.map(c => pcell(sv(get(sh, r, c)))).join('')}${isG ? pcell(sv(get(sh, r, 23))) + pcell(sv(get(sh, r, 24))) : pcell(sv(get(sh, r, 23))) + pcell(sv(get(sh, r, 24))) + pcell(sv(get(sh, r, 25)))}</tr>`;
    }
    let foot = '';
    if (hasFoot) for (const [r, lab] of [[51, 'รวม'], [52, 'เฉลี่ย'], [53, 'เฉลี่ย %']]) foot += `<tr style="height:19px;font-weight:700"><td style="${SB}"></td>${tdN(lab)}${cs.map(c => pcell(sv(get(sh, r, c)))).join('')}${pcell('')}${pcell('')}${pcell('')}</tr>`;
    const html = `<div style="${PF}font-size:15px">${attHead('ผลการพัฒนาคุณภาพผู้เรียน · ' + title + ' ' + tl)}<table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center">${h}${rows}${foot}</table>${homeSign()}</div>`;
    return { ...PG, land: false, html };
  }
  function sumDistPage(term) {
    const sh = 'เกรดร้อยละ ' + (term === 'sum' ? 'รวม' : term), GR = ['มส', 'ร', '0', '1', '1.5', '2', '2.5', '3', '3.5', '4'];
    let h = `<tr>${thS('รายวิชา', 'width:190px').replace('<th ', '<th rowspan="3" ')}<th colspan="3" rowspan="2" style="${SB}font-weight:400">จำนวนนักเรียนเข้าสอบ</th><th colspan="20" style="${SB}font-weight:400">ระดับผลการเรียน</th></tr><tr>${GR.map(g => `<th colspan="2" style="${SB}font-weight:400">${g}</th>`).join('')}</tr><tr>${['ชาย', 'หญิง', 'รวม'].map(x => thS(x)).join('')}${GR.map(() => thS('จำนวน', 'font-size:12px') + thS('%', 'font-size:12px')).join('')}</tr>`;
    let rows = '';
    for (let r = 8; r <= 27; r++) {
      const name = txt(sh, r, 1); if (!name) continue;
      rows += `<tr style="height:21px">${tdN(esc(name))}${[2, 3, 4].map(c => pcell(sv(get(sh, r, c)))).join('')}${Array.from({ length: 20 }, (_, k) => { const v = get(sh, r, 5 + k); return pcell(k % 2 && typeof v === 'number' ? (Math.round(v * 100) / 100).toFixed(2) : sv(v), 'font-size:13px'); }).join('')}</tr>`;
    }
    const html = `<div style="${PF}font-size:15px">${attHead('สรุปผลการประเมินผลสัมฤทธิ์ทางการเรียน ' + (term === 'sum' ? 'รวมทั้งปี' : 'ภาคเรียนที่ ' + term))}<table style="border-collapse:collapse;font-size:14px;width:100%;text-align:center">${h}${rows || `<tr><td colspan="25" style="${SB}padding:8px">ยังไม่มีรายวิชา</td></tr>`}</table>${homeSign()}</div>`;
    return { ...PG, land: false, html };
  }
  function assessDocPages(name) {
    if (name.startsWith('สรุป:')) { const k = name.split(':')[1]; return [1, 2, 'sum'].map(tm => k === 'dist' ? sumDistPage(tm) : sumClassPage(k, tm)); }
    const k = name.slice(name.indexOf(':') + 1);
    if (k === 'body') return [bodyPage()];
    return [1, 2, 'sum'].map(t => k === 'act' ? actPage(t) : assessPage(k, t));
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
      <div style="margin-top:12px">${p5AttBlock()}</div>
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
    const ass = [[39, groupName('trait', 0)], [40, groupName('trait', 1)], [41, groupName('value', 0)], [42, groupName('value', 1)]].map(([r, t]) => `<div class="row sp"><span>${t}</span>${chip(get(sh, r, 4))}</div>`).join('');
    return head + pick + `<div class="card" style="margin-top:12px"><div class="rhead"><div><h3>${esc(stuName(U.p6no - 1))}</h3><span class="hint">เลขที่ ${U.p6no} · รหัส ${esc(txt(ST, 3 + U.p6no - 1, 6))} · ${esc(txt(SUBJ, 2, 2))}</span></div>
        <div class="row"><button class="btn" data-act="doc" data-name="${esc(sh)}">ดูเอกสาร</button><button class="btn primary" data-act="p6print">พิมพ์คนนี้</button><button class="btn" data-act="p6all">พิมพ์ทั้งห้อง</button></div></div>
        <div class="row" style="margin-top:8px"><span class="hint">เล่ม ปพ.6 (รายงานประจำตัวนักเรียน ฉบับเต็ม):</span><button class="btn" data-act="p6book">พิมพ์เล่มคนนี้</button><button class="btn" data-act="p6bookall">พิมพ์เล่มทั้งห้อง</button></div></div>
      <div class="tw p5tw"><table class="t p5t"><thead><tr><th>รหัส</th><th class="stk">รายวิชา</th><th>เวลาเรียน</th><th>คะแนนเต็ม</th><th>เฉลี่ยชั้น</th><th>คะแนนที่ได้</th><th>ร้อยละ</th><th>ระดับ</th></tr></thead><tbody>${rows || '<tr><td colspan="8" class="hint">ยังไม่มีรายวิชา</td></tr>'}</tbody>
        <tfoot><tr><td colspan="3" class="stk">รวม · ร้อยละ ${esc(fmt2(get(sh, 30, 4)))}</td><td class="num">${esc(fmt2(get(sh, 29, 5)))}</td><td class="num">${esc(fmt2(get(sh, 29, 6)))}</td><td class="num">${esc(fmt2(get(sh, 29, 7)))}</td><td class="num">${esc(fmt2(get(sh, 29, 8)))}</td><td></td></tr></tfoot></table></div>
      <div class="grid2" style="margin-top:12px"><div class="card"><h3>ผลการเรียนเฉลี่ย</h3><div class="row sp"><span>ผลการเรียนเฉลี่ย</span><b>${esc(fmt2(get(sh, 32, 4)))}</b></div><div class="row sp"><span>คะแนนรวมได้ลำดับที่</span><b>${esc(fmt2(get(sh, 31, 4)))}</b></div><div class="row sp"><span>ผลการเรียนเฉลี่ยได้ลำดับที่</span><b>${esc(fmt2(get(sh, 33, 4)))}</b></div></div>
        <div class="card"><h3>กิจกรรมพัฒนาผู้เรียน</h3>${act}</div><div class="card full"><h3>ผลการประเมิน</h3>${ass}</div></div>${p6AttBlock(U.p6no - 1)}`;
  }

  // ---------------------------------------------------------------- เอกสาร (ชีทแบบ Excel)
  function vDocs() {
    const groups = A().sheetGroups();
    for (const g of groups) if (g.items.some(x => /^(คุณลักษณะ|ค่านิยม|กิจกรรม) /.test(x))) g.items = [...Object.keys(ASSESS_DOCS).filter(k => k.startsWith('ประเมิน:') && k !== 'ประเมิน:body'), ...g.items.filter(x => !/^(คุณลักษณะ|ค่านิยม|กิจกรรม) |^น้ำหนัก-ส่วนสูง$|^เวลาเรียน$/.test(x)), 'ประเมิน:body'];
    const gi = groups.findIndex(g => g.items.includes('ประเมิน:body'));
    groups.splice(gi + 1, 0, { title: 'เวลาเรียนประจำชั้น', items: [ATT] }, { title: 'เวลาเรียนรายวิชา', items: subjects().map(s => 'เวลาเรียนวิชา:' + s.slot) });
    for (const g of groups) if (g.items.some(x => /^(คะแนนรายชั้น|คะแนนร้อยละ|เกรดรายชั้น|เกรดร้อยละ) /.test(x))) g.items = ['ชื่อนักเรียน', 'บิดามารดา', 'ผู้ปกครอง', 'สรุป:score', 'สรุป:pct', 'สรุป:grade', 'สรุป:dist'];
    return `<div class="bar"><input class="in" data-search placeholder="ค้นหาเอกสาร เช่น ปพ.6, คะแนน, วิชาที่ 3" value="${esc(U.docq)}" style="max-width:420px"></div>
      <p class="hint">เปิดชีทแบบ Excel เพื่อดู แก้ไขช่องที่กรอกได้ หรือพิมพ์ — ข้อมูลเดียวกับที่กรอกในเมนูอื่น</p>
      ${groups.map(g => `<div class="card doc-g"><h3>${esc(g.title)}</h3><div class="dlist">${g.items.map(nm => { const m = /^วิชาที่ (\d+)$/.exec(nm), sn = m ? A().subjName(+m[1]) : '', ma = /^เวลาเรียนวิชา:(\d+)$/.exec(nm); return `<div class="ditem" data-q="${esc((nm + ' ' + sn + (ma ? ' ' + subjName(+ma[1]) : '')).toLowerCase())}"><button class="open" data-act="doc" data-name="${esc(nm)}">${esc(ma ? subjName(+ma[1]) + ' (รายเดือน + สรุปเทอม)' : nm === ATT ? 'เวลาเรียนประจำชั้น (รายเดือน + สรุปเทอม)' : ASSESS_DOCS[nm] ? ASSESS_DOCS[nm] + (/^(ประเมิน:(trait|value|act)|สรุป:)/.test(nm) ? ' (เทอม 1 · เทอม 2 · รวม)' : '') : nm)}${ma && !weekPeriods(+ma[1]) ? ' <span class="hint">ยังไม่ได้ตั้งตารางสอน</span>' : ''}${sn ? ` <span class="hint">${esc(sn)}</span>` : ''}</button><button class="btn" data-act="print-doc" data-name="${esc(nm)}">พิมพ์</button></div>`; }).join('')}</div></div>`).join('')}`;
  }
  function docFilter() {
    const q = U.docq.trim().toLowerCase();
    $$('#view .ditem').forEach(d => { d.hidden = !!q && !d.dataset.q.includes(q); });
    $$('#view .doc-g').forEach(g => { g.hidden = !$$('.ditem', g).some(d => !d.hidden); });
  }

  // ---------------------------------------------------------------- คู่มือ
  function vManual() {
    const S = [
      ['เริ่มต้นใช้งาน', `<ol><li>แท็บ <b>ข้อมูล › โรงเรียน</b> กรอกปีการศึกษา ชื่อ ผอ. และครูประจำชั้น</li><li><b>ข้อมูล › นักเรียน</b> กรอกหรือคัดลอกรายชื่อจาก Excel มาวาง (สูงสุด ${MAX_STU} คน)</li><li><b>ข้อมูล › นักเรียน</b> กรอกข้อมูลที่เล่ม ปพ.6 ใช้ (เลขประจำตัว วันเกิด ที่อยู่ ผู้ปกครอง บิดา-มารดา) หรือกด <b>นำเข้าไฟล์ DMC (.xlsx)</b> · ความเห็นครูกรอกที่ ข้อมูล › ความเห็นครู · วันที่ออกเอกสารตั้งที่ ข้อมูล › โรงเรียน</li><li><b>ข้อมูล › รายวิชา</b> กรอกรหัสวิชา ชื่อวิชา เวลาเรียน และครูผู้สอน</li><li>แท็บ <b>คะแนน</b> เลือกวิชา ตั้งตัวชี้วัดและคะแนนเต็ม แล้วกรอกคะแนนทีละภาคเรียน</li><li>แท็บ <b>ประเมิน</b> เลือกระดับคุณลักษณะ ค่านิยม และกิจกรรม · ปุ่ม <b>ปรับหัวข้อ</b> ใช้ตั้งชื่อกลุ่ม/หัวข้อ เพิ่มหรือลบหัวข้อของคุณลักษณะ ค่านิยม และสมรรถนะ</li><li>แท็บ <b>รายงาน ปพ.5 / ปพ.6</b> ดูสรุปผลและพิมพ์</li></ol>`],
      ['การกรอกข้อมูลในตาราง', `<ul><li>คลิกช่องแล้วพิมพ์ · <kbd>Enter</kbd> / <kbd>↓</kbd> ลงช่องล่าง · <kbd>Tab</kbd> ไปช่องขวา</li><li>คัดลอกหลายช่องจาก Excel แล้ววาง (<kbd>Ctrl</kbd>+<kbd>V</kbd>) ที่ช่องแรกได้ทั้งตาราง</li><li>ระบบบันทึกอัตโนมัติ ดูสถานะที่มุมขวาบน หรือกดปุ่ม “บันทึก”</li></ul>`],
      ['ปฏิทินและเวลาเรียน', `<ol><li><b>ข้อมูล › ปฏิทิน · เวลาเรียน › 1. ปฏิทินวันเรียน</b> กรอกวันเปิด-ปิดเรียนของ 2 ภาคเรียน กด “สร้างปฏิทินวันเรียน” แล้วคลิกวันหยุดราชการเพื่อปิดวันเรียน</li><li>กด “เสร็จแล้ว → กรอกเวลาเรียน” ระบบดึงวันเรียนไปสร้างตารางเช็คชื่อของแต่ละเดือนให้เอง</li><li><b>2. กรอกเวลาเรียน</b> แตะช่องเพื่อเปลี่ยน ✓ มา · ข ขาด · ป ป่วย · ล ลา หรือกด “มาครบทั้งเดือน” แล้วแก้เฉพาะคนที่ขาด</li><li>พิมพ์ได้เป็นแบบฟอร์มรายเดือน 1 หน้าต่อเดือน พร้อมหน้าสรุปทั้งปี</li></ol><p><b>เวลาเรียนรายวิชา</b> (แต่ละวิชาสอนคนละวัน): หลังเช็คเวลาเรียนรายห้อง ให้ไปที่ <b>3. ตารางสอน</b> กรอกจำนวนคาบต่อวันของแต่ละวิชา แล้วดูผลที่ <b>4. เวลาเรียนรายวิชา</b> — ระบบดึงจากเวลาเรียนรายห้อง (ครูประจำชั้นเช็คทุกวัน) มาในวันนั้นถือว่ามาทุกวิชาที่สอนในวันนั้น ครูประจำวิชาไม่ต้องลงซ้ำ แต่แก้เฉพาะวิชาของตนได้เมื่อนักเรียนเรียนไม่เต็มวัน (แตะช่อง) วันที่ไม่มีสอน (เช่น ครูติดราชการ) กดวันที่ที่หัวตารางเพื่อปิด ร้อยละคิดจากคาบที่ครูประจำชั้นเช็คแล้ว เกณฑ์ไม่น้อยกว่าร้อยละ 80 · ดูสรุปรายชั้นและรายวิชาได้ที่ รายงาน ปพ.5 และรายบุคคลที่ รายงาน ปพ.6</p>`],
      ['พิมพ์เอกสาร', `<p>เมนู <b>เอกสาร</b> เปิดชีทแบบทางการและกดพิมพ์ได้ หน้า ปพ.6 เลือกนักเรียนแล้วพิมพ์คนเดียว หรือพิมพ์ทั้งห้อง ระบบแสดงตัวอย่างก่อนพิมพ์เสมอ</p>`],
      ['สำรองข้อมูล', `<p>ข้อมูลเก็บบนระบบอัตโนมัติ ควรกด <b>ข้อมูล › สำรองข้อมูล</b> เพื่อดาวน์โหลดไฟล์เก็บไว้เป็นระยะ</p>`],
    ];
    return `<h2>คู่มือการใช้งาน</h2><p class="sub">โปรแกรมแจ้งผลการศึกษา ปพ.5 และ ปพ.6 หลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน พ.ศ. 2551 ระดับประถมศึกษา</p>
      <div class="man-toc">${S.map((s, i) => `<button class="chip" data-act="toc" data-k="${i}">${esc(s[0])}</button>`).join('')}</div>
      ${S.map((s, i) => `<details class="card man" id="man-${i}"${i === 0 ? ' open' : ''}><summary><span class="no">${i + 1}</span>${esc(s[0])}</summary><div class="man-body">${s[1]}</div></details>`).join('')}`;
  }

  root.PPUI = { mount, go, refresh, customPages, beforeSheet };
})(window);
