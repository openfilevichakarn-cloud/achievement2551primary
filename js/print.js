import * as C from './calc.js';

export const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const dl = (v, w) => `<span class="dl${w ? ' fix' : ''}"${w ? ` style="width:${w}mm"` : ''}>${esc(v) || '&nbsp;'}</span>`;
const sortStudents = room => [...(room.students ?? [])].sort((a, b) => (a.no || 0) - (b.no || 0));
const active = s => s.status !== 'ย้ายออก' && s.status !== 'ลาออก';
const ROWS = 45;

/* ---------- ข้อมูลที่ใช้ร่วมกัน ---------- */
export function evalOf(room, term, sid) {
  const get = t => room.eval?.[t]?.[sid] ?? {};
  const levelOf = k => {
    if (term === 'รวม') return C.modeLevel([...(get(1)[k] ?? []), ...(get(2)[k] ?? [])]);
    return C.modeLevel(get(term)[k] ?? []);
  };
  return { t: levelOf('t'), r: levelOf('r'), v: levelOf('v'), c: levelOf('c') };
}
export function actOf(room, term, sid) {
  const g = t => room.acts?.[t]?.[sid] ?? {};
  if (term === 'รวม') {
    const a = g(1), b = g(2), out = { club: '', clubName: b.clubName || a.clubName || '' };
    for (const k of ['guid', 'scout', 'club', 'social']) out[k] = a[k] && b[k] ? (a[k] === 'ผ' && b[k] === 'ผ' ? 'ผ' : 'มผ') : '';
    return out;
  }
  return g(term);
}

function classStats(room) {
  const st = room.students ?? [];
  const cnt = f => ({ m: st.filter(s => s.sex === 'ช' && f(s)).length, f: st.filter(s => s.sex === 'ญ' && f(s)).length });
  const rows = [
    ['นักเรียนต้นปีการศึกษา', cnt(s => s.status !== 'ย้ายเข้า')],
    ['ออกระหว่างปีการศึกษา', cnt(s => s.status === 'ย้ายออก' || s.status === 'ลาออก')],
    ['เข้าระหว่างปีการศึกษา', cnt(s => s.status === 'ย้ายเข้า')],
    ['รวมสิ้นปีการศึกษา', cnt(active)],
  ];
  return rows;
}
const statsBlock = room => `<div class="t-sm" style="margin:2mm 0 3mm">${classStats(room).map(([n, c]) =>
  `<div class="line" style="margin:0.6mm 0"><span style="width:44mm">${n}</span>ชาย${dl(c.m, 18)}คน <span style="margin-left:6mm">หญิง</span>${dl(c.f, 18)}คน <span style="margin-left:6mm">รวม</span>${dl(c.m + c.f, 18)}คน</div>`).join('')}</div>`;

function headLines(room, school, term) {
  return `<div class="c t-lg">สมุดบันทึกการพัฒนาคุณภาพผู้เรียน (ปพ.5)</div>
  <div class="c t-md">โรงเรียน${esc(school.name)}</div>
  <div class="c t-md" style="font-size:12.5pt">${esc(school.affiliation)}${school.affiliation ? '' : '&nbsp;'}</div>`;
}

function levelCounts(room, getter, n = 4) {
  const arr = Array(n).fill(0);
  sortStudents(room).filter(active).forEach(s => { const v = getter(s); if (v !== '' && v != null && v >= 0 && v < n) arr[v]++; });
  return arr;
}
const cell = v => (v ? v : '-');

function summaryBoxes(room, term) {
  const get = k => levelCounts(room, s => evalOf(room, term, s.id)[k]);
  const t = get('t'), r = get('r'), v = get('v'), c = get('c');
  const row = (title, h1, a, h2, b, lab1, lab2) => `<tr><td rowspan="3" class="b">${title}</td><td colspan="4" class="b">${h1}</td><td colspan="4" class="b">${h2}</td></tr>
    <tr>${lab1.map(x => `<td>${x}</td>`).join('')}${lab1.map(x => `<td>${x}</td>`).join('')}</tr>
    <tr>${[...a, ...b].map(x => `<td>${cell(x)}</td>`).join('')}</tr>`;
  const L = C.LV_TRAIT, K = C.LV_COMP;
  const fix = a => [...a];
  return `<table class="pt t-sm mt1" style="font-size:10pt">
    <colgroup><col style="width:26mm">${'<col>'.repeat(8)}</colgroup>
    <tr><td rowspan="3" class="b">สรุปการประเมิน<br>จำนวนนักเรียน</td><td colspan="4" class="b">คุณลักษณะอันพึงประสงค์</td><td colspan="4" class="b">การอ่าน คิดวิเคราะห์ และเขียน</td></tr>
    <tr>${L.map(x => `<td>${x}</td>`).join('')}${L.map(x => `<td>${x}</td>`).join('')}</tr>
    <tr>${[...fix(t), ...fix(r)].map(x => `<td>${cell(x)}</td>`).join('')}</tr>
    <tr><td rowspan="3" class="b">สรุปการประเมิน<br>จำนวนนักเรียน</td><td colspan="4" class="b">ค่านิยม 12 ประการ</td><td colspan="4" class="b">สมรรถนะสำคัญของผู้เรียน</td></tr>
    <tr>${L.map(x => `<td>${x}</td>`).join('')}${K.map(x => `<td>${x}</td>`).join('')}</tr>
    <tr>${[...fix(v), ...fix(c)].map(x => `<td>${cell(x)}</td>`).join('')}</tr>
  </table>`;
}

function approval(school, room, term, compact) {
  const d = room.signDate ? C.thaiDate(room.signDate) : '';
  const sq = compact ? '<span class="sq" style="width:8mm;height:8mm"></span>' : '<span class="sq"></span>';
  return `<div class="c t-md ${compact ? '' : 'mt2'}" style="${compact ? 'margin-top:2mm;font-size:13pt' : ''}">การอนุมัติผลการเรียน</div>
  <div style="margin:${compact ? '0 0 1mm 44%' : '2mm 0 4mm 44%'};font-size:${compact ? 11 : 12}pt"><div class="line"><span class="dl" style="flex:1">&nbsp;</span>ครูประจำวิชา/ครูประจำชั้น</div>
  <div class="line"><span class="dl" style="flex:1">&nbsp;</span>หัวหน้างานวิชาการโรงเรียน</div>
  <div class="line"><span class="dl" style="flex:1">&nbsp;</span>รองผู้อำนวยการโรงเรียน</div></div>
  <div class="row" style="display:flex;justify-content:space-around;margin:${compact ? 1 : 3}mm 0"><span>${sq} &nbsp;อนุมัติ</span><span>${sq} &nbsp;ไม่อนุมัติ</span></div>
  <div class="sig" style="${compact ? 'margin-top:1mm;font-size:11pt;line-height:1.2' : ''}">${dl('', 70)}<br>(${esc(school.director) || '..................................................'})<br>${esc(school.directorTitle) || 'ผู้อำนวยการโรงเรียน'}<br>วันที่ ${dl(d, 50)}</div>`;
}

/* ---------- ปก ปพ.5 ---------- */
function gradeCount(room, subj, term) {
  const out = { 'มส': 0, 'ร': 0, 0: 0, 1: 0, 1.5: 0, 2: 0, 2.5: 0, 3: 0, 3.5: 0, 4: 0 };
  sortStudents(room).filter(active).forEach(s => {
    const g = C.resultFor(subj, room.scores?.[subj.id]?.[s.id], term).grade;
    if (g !== '' && g in out) out[g]++;
  });
  return out;
}
const gradeHead = () => C.GRADE_LIST.map(g => `<th>${g}</th>`).join('');

function coverSubject(room, school, subj, term) {
  const cnt = gradeCount(room, subj, term);
  return `<div class="page frame frame2"><div class="box">
    ${headLines(room, school, term)}
    <div class="line mt2"><span>ชั้น</span>${dl(room.level + '/' + room.room)}<span>ภาคเรียนที่</span>${dl(term === 'รวม' ? '1-2' : term, 28)}<span>ปีการศึกษา</span>${dl(room.year, 36)}</div>
    <div class="line"><span>รายวิชา</span>${dl(subj.name)}<span>รหัส</span>${dl(subj.code, 30)}<span>เวลาเรียน</span>${dl(subj.hours, 22)}<span>ชั่วโมง/ปี</span></div>
    <div class="line"><span>กลุ่มสาระการเรียนรู้</span>${dl(subj.group)}</div>
    <div class="line mt2"><span class="dl" style="flex:1">${esc(subj.teacher) || '&nbsp;'}</span><span style="width:40%;flex:none"></span><span>ครูประจำวิชา</span></div>
    <div class="line"><span class="dl" style="flex:1">${esc(room.advisor) || '&nbsp;'}</span><span style="flex:none">ครูที่ปรึกษา/ครูประจำชั้น</span></div>
    ${statsBlock(room)}
    <div class="c t-md mt2" style="margin-bottom:2mm">สรุปผลสัมฤทธิ์ทางการเรียนรู้</div>
    <table class="pt"><tr><th style="width:30mm">ระดับ</th>${gradeHead()}</tr>
    <tr><td class="b">จำนวนนักเรียน</td>${C.GRADE_LIST.map(g => `<td>${cell(cnt[g])}</td>`).join('')}</tr></table>
    ${summaryBoxes(room, term)}
    ${approval(school, room, term)}
  </div></div>`;
}

function coverClass(room, school, term) {
  const subs = room.subjects ?? [];
  const rows = subs.map(sj => {
    const cnt = gradeCount(room, sj, term);
    return `<tr class="fill c4"><td>${esc(sj.code)}</td><td class="l">${esc(sj.name)}</td>${C.GRADE_LIST.map(g => `<td>${cell(cnt[g])}</td>`).join('')}</tr>`;
  });
  while (rows.length < 20) rows.push(`<tr class="fill c4"><td></td><td></td>${C.GRADE_LIST.map(() => '<td></td>').join('')}</tr>`);
  return `<div class="page frame frame2"><div class="box">
    ${headLines(room, school, term)}
    <div class="line mt2"><span>ชั้น</span>${dl(room.level + '/' + room.room)}<span>ภาคเรียนที่</span>${dl(term === 'รวม' ? '1-2' : term, 28)}<span>ปีการศึกษา</span>${dl(room.year, 36)}</div>
    <div class="line"><span class="dl" style="flex:1">${esc(room.advisor) || '&nbsp;'}</span><span>ครูสอน/ครูประจำชั้น</span></div>
    ${statsBlock(room)}
    <div class="c t-md" style="margin-bottom:1.5mm">สรุปผลสัมฤทธิ์ทางการเรียนรู้</div>
    <table class="pt"><colgroup><col style="width:17mm"><col style="width:46mm">${C.GRADE_LIST.map(() => '<col>').join('')}</colgroup>
      <tr><th>รหัส</th><th>รายวิชา</th>${gradeHead()}</tr>${rows.join('')}</table>
    ${summaryBoxes(room, term)}
    ${approval(school, room, term, true)}
  </div></div>`;
}

/* ---------- ตารางรายชั้น ---------- */
function titleBlock(lines) { return `<div class="c b t-sm" style="margin-bottom:1.5mm">${lines.map(esc).join('<br>')}</div>`; }
const signFoot = room => `<div class="c t-sm" style="position:absolute;bottom:8mm;left:0;right:0">ลงชื่อ....................................................ครูประจำชั้น/ครูประจำวิชา</div>`;
const roomLabel = room => `ชั้น${room.level}/${room.room} ปีการศึกษา ${room.year}`;

function pad(rows, n, cols) { const r = [...rows]; while (r.length < n) r.push(`<tr class="fill">${'<td></td>'.repeat(cols)}</tr>`); return r.join(''); }

function indPage(room, subj, term) {
  const c = C.ensureCfg(subj, term), N = c.inds.length;
  const st = sortStudents(room);
  const rows = st.map((s, i) => {
    const rec = room.scores?.[subj.id]?.[s.id];
    const r = C.termResult(subj, rec, term);
    const p = rec?.[term]?.p ?? [];
    return `<tr class="fill"><td>${i + 1}</td><td class="l">${esc(C.studentName(s))}</td>${c.inds.map((_, k) => `<td>${esc(p[k] ?? '')}</td>`).join('')}<td>${r.ind ? r.pass : ''}</td><td>${r.ind ? r.fail : ''}</td><td>${r.ind}</td></tr>`;
  });
  return `<div class="page"><div style="font-size:10pt">
    ${titleBlock([`ผลการประเมินมาตรฐาน/ตัวชี้วัด ${roomLabel(room)} ${C.termName(term)}`, `รายวิชา ${subj.code} ${subj.name}`])}
    <table class="pt" style="font-size:${N > 30 ? 7 : 9.5}pt"><colgroup><col style="width:8mm"><col style="width:${N > 30 ? 34 : 42}mm">${c.inds.map(() => '<col>').join('')}<col style="width:9mm"><col style="width:9mm"><col style="width:11mm"></colgroup>
    <tr><th rowspan="2">เลขที่</th><th rowspan="2">ชื่อ-สกุล</th><th colspan="${N}">ตัวชี้วัด</th><th rowspan="2"><div class="vt">ผ่าน</div></th><th rowspan="2"><div class="vt">ไม่ผ่าน</div></th><th rowspan="2"><div class="vt">สรุป</div></th></tr>
    <tr>${c.inds.map((n, k) => `<th style="font-size:8pt">${esc(n) || k + 1}</th>`).join('')}</tr>
    ${pad(rows, ROWS, N + 5)}</table></div>${signFoot()}</div>`;
}

function scorePage(room, subj, term) {
  const st = sortStudents(room);
  const rowsOut = [];
  const colAvg = {};
  if (term === 'รวม') {
    const body = st.map((s, i) => {
      const y = C.yearResult(subj, room.scores?.[subj.id]?.[s.id]);
      const f = r => `<td>${C.fmt(r.between)}</td><td>${C.fmt(r.fin, 0)}</td><td>${C.fmt(r.total)}</td><td>${C.fmt(r.pct)}</td><td>${esc(r.grade)}</td>`;
      return `<tr class="fill"><td>${i + 1}</td><td class="l">${esc(C.studentName(s))}</td>${f(y.a)}${f(y.b)}<td>${C.fmt(y.total)}</td><td>${C.fmt(y.pct)}</td><td>${esc(y.grade)}</td></tr>`;
    });
    return `<div class="page"><div style="font-size:10pt">
      ${titleBlock([`ผลสัมฤทธิ์ทางการเรียน สรุปรายปี ${roomLabel(room)}`, `รายวิชา ${subj.code} ${subj.name}`])}
      <table class="pt" style="font-size:9pt"><colgroup><col style="width:8mm"><col style="width:40mm">${'<col>'.repeat(13)}</colgroup>
      <tr><th rowspan="2">เลขที่</th><th rowspan="2">ชื่อ-สกุล</th><th colspan="5">ภาคเรียนที่ 1</th><th colspan="5">ภาคเรียนที่ 2</th><th colspan="3">รวมรายปี</th></tr>
      <tr>${['ระหว่างภาค', 'ปลายภาค', 'รวม', 'ร้อยละ', 'เกรด'].map(x => `<th><div class="vt">${x}</div></th>`).join('').repeat(2)}<th><div class="vt">รวมรายปี</div></th><th><div class="vt">ร้อยละ</div></th><th><div class="vt">ระดับผลการเรียน</div></th></tr>
      ${pad(body, ROWS, 15)}</table></div>${signFoot()}</div>`;
  }
  const c = C.ensureCfg(subj, term), K = c.max.length;
  const res = st.map(s => C.termResult(subj, room.scores?.[subj.id]?.[s.id], term));
  const avgOf = f => { const v = res.map(f).filter(C.isNum); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
  const body = st.map((s, i) => {
    const r = res[i], sc = room.scores?.[subj.id]?.[s.id]?.[term]?.s ?? [];
    return `<tr class="fill"><td>${i + 1}</td><td class="l">${esc(C.studentName(s))}</td>${c.max.map((_, k) => `<td>${esc(sc[k] ?? '')}</td>`).join('')}<td>${C.fmt(r.between)}</td><td>${C.fmt(r.fin, 0)}</td><td>${C.fmt(r.total)}</td><td>${C.fmt(r.pct)}</td><td>${esc(r.grade)}</td></tr>`;
  });
  const sumCols = (lbl, f) => `<tr><td colspan="${2 + K}" class="b">${lbl}</td>${[r => r.between, r => r.fin, r => r.total, r => r.pct].map(g => `<td>${C.fmt(f(g))}</td>`).join('')}<td></td></tr>`;
  const tot = g => { const v = res.map(g).filter(C.isNum); return v.length ? v.reduce((a, b) => a + b, 0) : null; };
  return `<div class="page"><div style="font-size:10pt">
    ${titleBlock([`ผลสัมฤทธิ์ทางการเรียน ${roomLabel(room)} ${C.termName(term)}`, `รายวิชา ${subj.code} ${subj.name}`, `อัตราส่วนคะแนน ระหว่างภาค : ปลายภาค = ${c.rb} : ${c.rf}`])}
    <table class="pt" style="font-size:9.5pt"><colgroup><col style="width:8mm"><col style="width:46mm">${c.max.map(() => '<col>').join('')}<col style="width:11mm"><col style="width:11mm"><col style="width:11mm"><col style="width:11mm"><col style="width:11mm"></colgroup>
    <tr><th rowspan="2">เลขที่</th><th rowspan="2">ชื่อ-สกุล</th><th colspan="${K}">คะแนนระหว่างเรียน</th><th rowspan="2"><div class="vt">ระหว่างภาค</div></th><th rowspan="2"><div class="vt">สอบปลายภาค</div></th><th rowspan="2"><div class="vt">รวมรายภาค</div></th><th rowspan="2"><div class="vt">ร้อยละ</div></th><th rowspan="2"><div class="vt">ระดับผลการเรียน</div></th></tr>
    <tr>${c.max.map((m, k) => `<th style="font-weight:400">${m}</th>`).join('')}</tr>
    ${pad(body, ROWS, K + 7)}
    <tr><td colspan="${2 + K}" class="b">รวม</td><td>${C.fmt(tot(r => r.between))}</td><td>${C.fmt(tot(r => r.fin))}</td><td>${C.fmt(tot(r => r.total))}</td><td></td><td></td></tr>
    <tr><td colspan="${2 + K}" class="b">เฉลี่ย</td><td>${C.fmt(avgOf(r => r.between))}</td><td>${C.fmt(avgOf(r => r.fin))}</td><td>${C.fmt(avgOf(r => r.total))}</td><td>${C.fmt(avgOf(r => r.pct))}</td><td></td></tr>
    </table></div>${signFoot()}</div>`;
}

/* ---------- คุณลักษณะ / ค่านิยม / กิจกรรม ---------- */
function evalSheet(room, term, kind) {
  const st = sortStudents(room);
  const isT = kind === 't';
  const A = isT ? { k: 't', n: 8, labels: C.TRAITS_SHORT, vert: true, h: 'คุณลักษณะอันพึงประสงค์', lv: C.LV_TRAIT } : { k: 'v', n: 12, labels: [...Array(12)].map((_, i) => i + 1), vert: false, h: 'ค่านิยมหลัก 12 ประการ', lv: C.LV_TRAIT };
  const B = isT ? { k: 'r', n: 5, labels: [1, 2, 3, 4, 5], h: 'การอ่าน คิดวิเคราะห์ และเขียนสื่อความ', lv: C.LV_TRAIT } : { k: 'c', n: 5, labels: [1, 2, 3, 4, 5], h: 'สมรรถนะสำคัญของผู้เรียน', lv: C.LV_COMP };
  const get = (s, k) => (term === 'รวม' ? [...(room.eval?.[1]?.[s.id]?.[k] ?? []), ...(room.eval?.[2]?.[s.id]?.[k] ?? [])] : room.eval?.[term]?.[s.id]?.[k] ?? []);
  const body = st.map((s, i) => {
    const e = evalOf(room, term, s.id);
    const cells = (X) => {
      const vals = term === 'รวม' ? null : get(s, X.k);
      const lv = e[X.k];
      return (vals ? Array.from({ length: X.n }, (_, j) => `<td>${C.isNum(vals[j]) ? vals[j] : ''}</td>`).join('') : Array.from({ length: X.n }, () => '<td></td>').join('')) + [0, 1, 2, 3].map(l => `<td>${lv !== '' && lv === l ? '/' : ''}</td>`).join('');
    };
    return `<tr class="fill"><td>${i + 1}</td>${cells(A)}${cells(B)}</tr>`;
  });
  const hdr = X => `<th colspan="${X.n}">${X.h}</th><th colspan="4">ผลการประเมิน</th>`;
  const sub = X => X.labels.map(l => `<th style="font-size:8.5pt;font-weight:400">${X.vert ? `<div class="vt">${esc(l)}</div>` : l}</th>`).join('') + X.lv.map(l => `<th style="font-size:8pt"><div class="vt">${l}</div></th>`).join('');
  const total = 1 + A.n + 4 + B.n + 4;
  return `<div class="page"><div style="font-size:10pt">
    ${titleBlock([`สรุปการประเมิน${isT ? 'คุณลักษณะอันพึงประสงค์ - การอ่าน คิดวิเคราะห์ และเขียนสื่อความ' : 'ค่านิยมหลัก 12 ประการ และสมรรถนะสำคัญของผู้เรียน'}`, `${roomLabel(room)} ${C.termName(term)}`])}
    <table class="pt" style="font-size:9pt"><colgroup><col style="width:9mm">${'<col>'.repeat(total - 1)}</colgroup>
    <tr><th rowspan="2">เลขที่</th>${hdr(A)}${hdr(B)}</tr><tr>${sub(A)}${sub(B)}</tr>
    ${pad(body, ROWS, total)}</table>
    <div class="t-xs mt1" style="line-height:1.3"><b>เกณฑ์การประเมิน</b> คะแนน 3 = ดีเยี่ยม, 2 = ดี, 1 = ผ่าน, 0 = ไม่ผ่าน (${isT ? '' : 'สมรรถนะ: 3 = ดีเยี่ยม, 2 = ดี, 1 = พอใช้, 0 = ปรับปรุง; '}ผลการประเมิน = ฐานนิยมของคะแนนรายข้อ)</div>
    </div>${signFoot()}</div>`;
}

function actsSheet(room, term) {
  const st = sortStudents(room);
  const mk = v => [v === 'มผ' ? '/' : '', v === 'ผ' ? '/' : ''];
  const body = st.map((s, i) => {
    const a = actOf(room, term, s.id);
    return `<tr class="fill"><td>${i + 1}</td><td class="l">${esc(C.studentName(s))}</td>${mk(a.guid).map(x => `<td>${x}</td>`).join('')}${mk(a.scout).map(x => `<td>${x}</td>`).join('')}<td class="l">${esc(a.clubName)}</td>${mk(a.club).map(x => `<td>${x}</td>`).join('')}${mk(a.social).map(x => `<td>${x}</td>`).join('')}</tr>`;
  });
  const cnt = (k, v) => st.filter(s => actOf(room, term, s.id)[k] === v).length;
  const tot = k => ['มผ', 'ผ'].map(v => `<td>${cnt(k, v)}</td>`).join('');
  const pc = k => { const n = st.filter(s => actOf(room, term, s.id)[k]).length; return ['มผ', 'ผ'].map(v => `<td>${n ? C.fmt(cnt(k, v) * 100 / n) : ''}</td>`).join(''); };
  return `<div class="page"><div style="font-size:10pt">
    ${titleBlock([`สรุปการประเมินกิจกรรมพัฒนาผู้เรียน ${roomLabel(room)} ${C.termName(term)}`])}
    <table class="pt" style="font-size:9.5pt"><colgroup><col style="width:9mm"><col style="width:40mm"><col><col><col><col><col style="width:30mm"><col><col><col><col></colgroup>
    <tr><th rowspan="2">เลขที่</th><th rowspan="2">ชื่อ-สกุล</th><th colspan="2">แนะแนว</th><th colspan="2">ลูกเสือ-เนตรนารี</th><th rowspan="2">ชุมนุม</th><th colspan="2">ผลชุมนุม</th><th colspan="2">กิจกรรมเพื่อสังคมและสาธารณประโยชน์</th></tr>
    <tr>${'<th style="font-weight:400"><div class="vt">ไม่ผ่าน</div></th><th style="font-weight:400"><div class="vt">ผ่าน</div></th>'.repeat(2)}<th style="font-weight:400"><div class="vt">ไม่ผ่าน</div></th><th style="font-weight:400"><div class="vt">ผ่าน</div></th><th style="font-weight:400"><div class="vt">ไม่ผ่าน</div></th><th style="font-weight:400"><div class="vt">ผ่าน</div></th></tr>
    ${pad(body, ROWS, 10)}
    <tr><td colspan="2" class="b">รวม</td>${tot('guid')}${tot('scout')}<td></td>${tot('club')}${tot('social')}</tr>
    <tr><td colspan="2" class="b">ร้อยละ</td>${pc('guid')}${pc('scout')}<td></td>${pc('club')}${pc('social')}</tr>
    </table></div>${signFoot()}</div>`;
}

function attendSheet(room) {
  const st = sortStudents(room);
  const days = room.att?.days ?? {};
  const full = C.MONTHS.reduce((a, m) => a + (Number(days[m]) || 0), 0);
  const body = st.map((s, i) => {
    const p = room.att?.present?.[s.id] ?? {};
    const sum = C.MONTHS.reduce((a, m) => a + (Number(p[m]) || 0), 0);
    return `<tr class="fill"><td>${i + 1}</td><td class="l">${esc(C.studentName(s))}</td>${C.MONTHS.map(m => `<td>${esc(p[m] ?? '')}</td>`).join('')}<td>${sum || ''}</td><td>${full && sum ? C.fmt(sum * 100 / full, 1) : ''}</td></tr>`;
  });
  return `<div class="page"><div style="font-size:10pt">
    ${titleBlock([`สรุปเวลาเรียน (จำนวนวันที่มาเรียน) ${roomLabel(room)}`])}
    <table class="pt" style="font-size:9pt"><colgroup><col style="width:8mm"><col style="width:38mm">${'<col>'.repeat(14)}</colgroup>
    <tr><th>เลขที่</th><th>ชื่อ-สกุล</th>${C.MONTHS.map(m => `<th><div class="vt" style="font-weight:400">${m}</div></th>`).join('')}<th><div class="vt">รวม</div></th><th><div class="vt">ร้อยละ</div></th></tr>
    <tr><td colspan="2" class="b">วันเรียนเต็ม</td>${C.MONTHS.map(m => `<td>${esc(days[m] ?? '')}</td>`).join('')}<td>${full || ''}</td><td></td></tr>
    ${pad(body, ROWS, 16)}</table></div>${signFoot()}</div>`;
}

function healthSheet(room) {
  const st = sortStudents(room);
  const dates = room.health?.dates ?? [];
  const body = st.map((s, i) => {
    const h = room.health?.of?.[s.id] ?? {};
    return `<tr class="fill"><td>${i + 1}</td><td class="l">${esc(C.studentName(s))}</td>${[0, 1, 2, 3].map(k => `<td>${esc(h.w?.[k] ?? '')}</td><td>${esc(h.h?.[k] ?? '')}</td>`).join('')}</tr>`;
  });
  return `<div class="page"><div style="font-size:11pt">
    ${titleBlock([`บันทึกน้ำหนัก - ส่วนสูง ${roomLabel(room)}`])}
    <table class="pt"><colgroup><col style="width:10mm"><col style="width:50mm">${'<col>'.repeat(8)}</colgroup>
    <tr><th rowspan="2">เลขที่</th><th rowspan="2">ชื่อ-สกุล</th><th colspan="4">ภาคเรียนที่ 1</th><th colspan="4">ภาคเรียนที่ 2</th></tr>
    <tr>${[0, 1, 2, 3].map(k => `<th colspan="2" style="font-weight:400">ครั้งที่ ${k % 2 + 1}<br><span class="t-xs">${C.thaiDate(dates[k])}</span></th>`).join('')}</tr>
    <tr>${'<th></th>'.repeat(2)}${'<th class="t-xs">น้ำหนัก</th><th class="t-xs">ส่วนสูง</th>'.repeat(4)}</tr>
    ${pad(body, ROWS, 10)}</table></div>${signFoot()}</div>`;
}

/* ---------- ปพ.6 ---------- */
function studentTotals(room, term) {
  const subs = room.subjects ?? [];
  return sortStudents(room).filter(active).map(s => {
    let tot = 0, full = 0, gw = 0, hw = 0;
    subs.forEach(sj => {
      const r = C.resultFor(sj, room.scores?.[sj.id]?.[s.id], term);
      if (r.total != null) { tot += r.total; full += r.full; }
      if (C.isNum(r.grade) && r.grade !== '') { gw += Number(r.grade) * (Number(sj.hours) || 1); hw += Number(sj.hours) || 1; }
    });
    return { id: s.id, tot, full, gpa: hw ? gw / hw : null };
  });
}
const rankOf = (arr, id, key) => {
  const me = arr.find(x => x.id === id);
  if (!me || me[key] == null || (key === 'tot' && !me.full)) return '';
  return 1 + arr.filter(x => x[key] != null && x[key] > me[key]).length;
};

function p6Cover(room, school, s, term) {
  const ag = C.ageOn(s.birth);
  const d = room.signDate ? C.thaiDate(room.signDate) : '';
  return `<div class="page frame frame2"><div class="box" style="padding-top:22mm">
    <div class="c t-xl">แบบรายงานประจำตัวนักเรียน</div><div class="c t-xl">ผลการพัฒนาคุณภาพผู้เรียนรายบุคคล (ปพ.6)</div>
    <div class="c t-xl">โรงเรียน${esc(school.name)}</div>
    <div class="c t-md mt3" style="font-weight:600">${esc(school.affiliation)}</div>
    <div class="c t-md mt1" style="font-weight:600">${esc([school.subdistrict && 'ต.' + school.subdistrict, school.district && 'อ.' + school.district, school.province && 'จ.' + school.province].filter(Boolean).join(' '))}</div>
    <div style="margin-top:20mm;font-size:13pt">
      <div class="line"><span>ชื่อ</span>${dl(s.prefix + s.first)}<span>นามสกุล</span>${dl(s.last)}</div>
      <div class="line"><span>วันเกิด</span>${dl(C.thaiDate(s.birth))}<span>อายุ</span>${dl(ag.y, 14)}<span>ปี</span>${dl(ag.m, 14)}<span>เดือน</span></div>
      <div class="line"><span>เลขประจำตัวนักเรียน</span>${dl(s.sid, 36)}<span>เลขประจำตัวประชาชน</span>${dl(s.cid)}</div>
      <div class="line"><span>ชั้น</span>${dl(room.level + '/' + room.room)}<span>เลขที่</span>${dl(s.no)}</div>
      <div class="line"><span>ภาคเรียนที่</span>${dl(term === 'รวม' ? '1-2' : term, 40)}<span>ปีการศึกษา</span>${dl(room.year)}</div>
    </div>
    <div style="margin:22mm 0 0 28mm;font-size:13pt;line-height:1.35">
      <div class="line"><span>ลงชื่อ</span>${dl('', 70)}</div>
      <div style="margin-left:14mm">(${esc(room.advisor) || '.............................................'})<br>ครูประจำชั้น/ครูที่ปรึกษา</div>
      <div class="line mt2"><span>ลงชื่อ</span>${dl('', 70)}</div>
      <div style="margin-left:14mm">(${esc(school.academic) || '.............................................'})<br>หัวหน้างานวิชาการโรงเรียน</div>
      <div class="line mt2"><span>ลงชื่อ</span>${dl('', 70)}</div>
      <div style="margin-left:14mm">(${esc(school.director) || '.............................................'})<br>${esc(school.directorTitle) || 'ผู้อำนวยการโรงเรียน'}</div>
      <div class="line mt2" style="margin-left:14mm"><span>วันที่</span>${dl(d, 60)}</div>
    </div></div></div>`;
}

const p6Advice = () => `<div class="page"><div style="padding:2mm 4mm;font-size:12pt;line-height:1.3">
  <div class="c t-md" style="margin-bottom:6mm">คำแนะนำสำหรับผู้ปกครอง</div>
  <p class="b">เรียน ท่านผู้ปกครองนักเรียน</p>
  <p style="text-indent:14mm">เมื่อท่านได้รับแบบรายงานประจำตัวนักเรียนนี้แล้ว โปรดสละเวลาพิจารณาข้อมูลต่าง ๆ ดังนี้</p>
  <p style="padding-left:14mm;text-indent:-0mm;margin:1mm 0">1. โปรดตรวจสอบความถูกต้องของข้อมูลนักเรียน และบันทึกหากมีการเปลี่ยนแปลงแก้ไขข้อมูล</p>
  <p style="padding-left:14mm;margin:1mm 0">2. โปรดตรวจสอบผลการประเมินภาวะโภชนาการ จากน้ำหนัก - ส่วนสูง ตามเกณฑ์มาตรฐาน ถ้ามีผลผิดปกติ เช่น น้ำหนักน้อยกว่าเกณฑ์ เตี้ย อ้วน เริ่มอ้วน หรือ ผอม ควรหาทางช่วยเหลือ หรือปรึกษาแพทย์ ในกรณีที่บุตรหลานของท่านมีโรคประจำตัว หรือมีสิ่งผิดปกติใดๆ ควรแจ้งให้ครูประจำชั้นทราบด้วย</p>
  <p style="padding-left:14mm;margin:1mm 0">3. โปรดตรวจสอบผลการมาเรียนของเด็กอย่างสม่ำเสมอ ติดต่อกับโรงเรียนทันทีเมื่อมีปัญหา เด็กหยุดเรียนโดยไม่ได้รับอนุญาต ทั้งนี้เพื่อป้องกันไม่ให้เกิดปัญหาพฤติกรรมหรือขาดเรียนนาน</p>
  <p style="padding-left:14mm;margin:1mm 0">4. โปรดตรวจสอบผลการเรียนรายวิชา ผลการประเมินกิจกรรมพัฒนาผู้เรียน ผลการประเมินคุณลักษณะอันพึงประสงค์ ผลการประเมินการอ่าน คิดวิเคราะห์และเขียน ผลการประเมินค่านิยม 12 ประการ และผลการประเมินสมรรถนะสำคัญของผู้เรียน โดยตรวจสอบว่านักเรียนมีผลการประเมินด้านใด อยู่ในระดับใด ผ่านเกณฑ์การประเมินการศึกษาของสถานศึกษาหรือไม่ และควรได้รับการช่วยเหลือด้านใด</p>
  <p style="padding-left:14mm;margin:1mm 0">5. เกณฑ์การประเมิน</p>
  <p style="padding-left:20mm;margin:.5mm 0">5.1 นักเรียนต้องมีผลการประเมินรายวิชาตั้งแต่ระดับ 1 ขึ้นไปทุกวิชา</p>
  <p style="padding-left:20mm;margin:.5mm 0">5.2 นักเรียนมีผลการประเมินการอ่าน คิดวิเคราะห์ และเขียน ผ่านเกณฑ์การประเมินในระดับดีเยี่ยม/ดี/ผ่าน</p>
  <p style="padding-left:20mm;margin:.5mm 0">5.3 นักเรียนมีผลการประเมินคุณลักษณะอันพึงประสงค์ ผ่านเกณฑ์การประเมินในระดับ ดีเยี่ยม/ดี/ผ่าน</p>
  <p style="padding-left:20mm;margin:.5mm 0">5.4 นักเรียนเข้าร่วมกิจกรรมพัฒนาผู้เรียน และได้ผลการประเมิน “ผ” ทุกกิจกรรม</p>
  <p style="padding-left:20mm;margin:.5mm 0">5.5 นักเรียนมีผลการประเมินค่านิยม 12 ประการ ผ่านเกณฑ์การประเมินในระดับ ดีเยี่ยม/ดี/ผ่าน</p>
  <p style="padding-left:20mm;margin:.5mm 0">5.6 นักเรียนมีผลการประเมินสมรรถนะสำคัญของผู้เรียน ผ่านเกณฑ์การประเมินในระดับ ดีเยี่ยม/ดี/ผ่าน</p>
  <p class="b" style="margin:3mm 0 1mm">คำอธิบายเกณฑ์ ผลการประเมินรายวิชา</p>
  <table class="pt dot" style="font-size:10.5pt"><tr><td>ช่วงคะแนน</td><td>ผลการเรียน</td><td>ความหมาย</td><td>ช่วงคะแนน</td><td>ผลการเรียน</td><td>ความหมาย</td><td>ผลการเรียน</td><td>ความหมาย</td></tr>
  ${[['80-100', '4', 'ดีเยี่ยม', '60-64', '2', 'ปานกลาง', 'ร', 'รอการตัดสิน'], ['75-79', '3.5', 'ดีมาก', '55-59', '1.5', 'พอใช้', 'มส', 'ไม่มีสิทธิ์สอบ'], ['70-74', '3', 'ดี', '50-54', '1', 'ผ่านเกณฑ์ขั้นต่ำ', 'ผ', 'ผ่าน'], ['65-69', '2.5', 'ค่อนข้างดี', '0-49', '0', 'ต่ำกว่าเกณฑ์', 'มผ', 'ไม่ผ่าน']].map(r => `<tr>${r.map(x => `<td>${x}</td>`).join('')}</tr>`).join('')}</table>
  <p style="padding-left:14mm;margin:3mm 0 1mm">6. โปรดตรวจสอบความคิดเห็นและข้อเสนอแนะของครูประจำชั้นต่อนักเรียน</p>
  <p style="padding-left:14mm;margin:1mm 0">7. โปรดสละเวลาให้ความคิดเห็นและเสนอแนะเกี่ยวกับตัวนักเรียน ความเห็นของท่านจะเป็นประโยชน์ต่อตัวนักเรียนในปกครองของท่านเอง เพราะจะช่วยให้ครูเข้าใจนักเรียนได้ดียิ่งขึ้น และจะช่วยพัฒนานักเรียนได้ถูกต้องต่อไป</p>
  <p class="r b" style="margin-top:8mm;padding-right:30mm">ขอขอบพระคุณ</p></div></div>`;

function p6Info(room, school, s) {
  const ag = C.ageOn(s.birth);
  const addr = [s.addrNo && 'บ้านเลขที่ ' + s.addrNo, s.addrMoo && 'หมู่ ' + s.addrMoo, s.addrRoad, s.addrSub && 'ต.' + s.addrSub, s.addrDist && 'อ.' + s.addrDist, s.addrProv && 'จ.' + s.addrProv].filter(Boolean).join(' ');
  const L = (a, b, c, d, wa = 0, wb = 0) => `<div class="line">${a}${dl(b)}${c ? c + dl(d) : ''}</div>`;
  return `<div class="page"><div style="padding:4mm;font-size:13pt">
    <div class="c t-md" style="margin-bottom:6mm">ข้อมูลนักเรียน</div>
    <div class="line"><span>ชื่อ</span>${dl(s.prefix + s.first)}<span>สกุล</span>${dl(s.last)}</div>
    <div class="line"><span>วันเกิด</span>${dl(C.thaiDate(s.birth))}<span>อายุ</span>${dl(ag.y, 14)}<span>ปี</span>${dl(ag.m, 14)}<span>เดือน</span></div>
    <div class="line"><span>เพศ</span>${dl(s.sex === 'ช' ? 'ชาย' : s.sex === 'ญ' ? 'หญิง' : '', 22)}<span>เชื้อชาติ</span>${dl(s.ethnic, 28)}<span>สัญชาติ</span>${dl(s.nation, 28)}<span>ศาสนา</span>${dl(s.religion, 28)}</div>
    <div class="line"><span>เลขประจำตัวนักเรียน</span>${dl(s.sid, 38)}<span>เลขประจำตัวประชาชน</span>${dl(s.cid)}</div>
    <div class="line"><span>ที่อยู่ปัจจุบัน</span>${dl(addr)}</div>
    <div class="line"><span>ชื่อ-สกุลบิดา</span>${dl(s.father)}<span>อาชีพ</span>${dl(s.fatherJob, 50)}</div>
    <div class="line"><span>ชื่อ-สกุลมารดา</span>${dl(s.mother)}<span>อาชีพ</span>${dl(s.motherJob, 50)}</div>
    <div class="line"><span>สถานภาพสมรสของบิดา-มารดา</span>${dl(s.marital)}</div>
    <div class="line"><span>ชื่อ-สกุลผู้ปกครอง</span>${dl(s.guardian)}<span>อาชีพ</span>${dl(s.guardianJob, 50)}</div>
    <div class="line"><span>ความเกี่ยวข้องกับนักเรียน</span>${dl(s.relation)}</div>
    <div class="c t-md" style="margin:12mm 0 4mm">บันทึกการเปลี่ยนแปลงหรือแก้ไขข้อมูล</div>
    ${'<div class="line" style="height:9.2mm"><span class="dl" style="flex:1">&nbsp;</span></div>'.repeat(11)}
  </div></div>`;
}

function p6Results(room, school, s, term, totals) {
  const subs = room.subjects ?? [];
  const sRows = subs.map(sj => {
    const r = C.resultFor(sj, room.scores?.[sj.id]?.[s.id], term);
    const all = sortStudents(room).filter(active).map(x => C.resultFor(sj, room.scores?.[sj.id]?.[x.id], term).total).filter(C.isNum);
    const mean = all.length ? all.reduce((a, b) => a + b, 0) / all.length : null;
    return `<tr class="fill"><td>${esc(sj.code)}</td><td class="l">${esc(sj.name)}</td><td>${esc(sj.hours)}</td><td>${C.fmt(r.full)}</td><td>${C.fmt(mean)}</td><td>${C.fmt(r.total)}</td><td>${C.fmt(r.pct)}</td><td>${esc(r.grade)}</td></tr>`;
  });
  while (sRows.length < 20) sRows.push(`<tr class="fill">${'<td></td>'.repeat(8)}</tr>`);
  const me = totals.find(x => x.id === s.id) ?? { tot: 0, full: 0, gpa: null };
  const hoursSum = subs.reduce((a, b) => a + (Number(b.hours) || 0), 0);
  const fullSum = me.full, totSum = me.tot;
  const a = actOf(room, term, s.id), e = evalOf(room, term, s.id);
  const ac = v => (v === 'ผ' ? 'ผ่าน' : v === 'มผ' ? 'ไม่ผ่าน' : '');
  const lv = (v, L) => (v === '' ? '' : L[v]);
  const d = room.signDate ? C.thaiDate(room.signDate) : '';
  const sigBlock = (name, role) => `<div class="c t-sm" style="margin-bottom:2mm">ลงชื่อ${dl('', 55)}<br>(${esc(name) || '..............................................'})<br>${role}<br>${d}</div>`;
  const sumRow = (l, v) => `<tr><td class="l b" style="padding-left:2mm" colspan="2">${l}</td><td>${v}</td></tr>`;
  return `<div class="page"><div style="padding:0 2mm;font-size:12pt">
    <div class="c t-md">แบบรายงานประจำตัวนักเรียน : ผลการพัฒนาคุณภาพผู้เรียนรายบุคคล (ปพ.6)</div>
    <div class="c t-md" style="font-size:12.5pt">โรงเรียน${esc(school.name)} ${esc(school.affiliation)}</div>
    <div class="c t-md" style="font-size:12.5pt">ชั้น${esc(room.level)}/${esc(room.room)} ${term === 'รวม' ? 'สรุปรายปี' : 'ภาคเรียนที่ ' + term} ปีการศึกษา ${esc(room.year)}</div>
    <div class="line mt2"><span>ชื่อ-สกุล</span>${dl(C.studentName(s))}<span>เลขประจำตัว</span>${dl(s.sid, 32)}<span>เลขที่</span>${dl(s.no, 14)}</div>
    <table class="pt mt1" style="font-size:10.5pt"><colgroup><col style="width:7mm"><col style="width:20mm"><col><col style="width:19mm"><col style="width:15mm"><col style="width:19mm"><col style="width:16mm"><col style="width:16mm"><col style="width:17mm"></colgroup>
    <tr><th rowspan="22"><div class="vt" style="font-size:11pt">ผลการเรียนรายวิชา</div></th><th>รหัสวิชา</th><th>รายวิชา</th><th>เวลาเรียน<br>(ชั่วโมง/ปี)</th><th>คะแนน<br>เต็ม</th><th>ค่าเฉลี่ย<br>ในชั้นเรียน</th><th>คะแนน<br>ที่ได้</th><th>คิดเป็น<br>ร้อยละ</th><th>ระดับผล<br>การเรียน</th></tr>
    ${sRows.join('')}
    <tr><td colspan="2" class="b">รวม</td><td>${hoursSum || ''}</td><td>${C.fmt(fullSum)}</td><td></td><td>${C.fmt(totSum)}</td><td>${fullSum ? C.fmt(totSum * 100 / fullSum) : ''}</td><td></td></tr></table>
    <div style="display:flex;border:.6pt solid #000;border-top:0;font-size:11pt">
      <table class="pt" style="flex:1.15;border:0;table-layout:auto"><colgroup><col><col style="width:22mm"></colgroup>
        ${sumRow('คะแนนคิดเป็นร้อยละ', fullSum ? C.fmt(totSum * 100 / fullSum) : '')}
        ${sumRow('คะแนนรวมได้ลำดับที่', rankOf(totals, s.id, 'tot'))}
        ${sumRow('ผลการเรียนเฉลี่ย', me.gpa == null ? '' : C.fmt(me.gpa))}
        ${sumRow('ผลการเรียนเฉลี่ยได้ลำดับที่', rankOf(totals, s.id, 'gpa'))}
        <tr><td colspan="3" class="l b" style="padding-left:2mm">ผลการประเมินกิจกรรมพัฒนาผู้เรียน</td></tr>
        ${sumRow('&nbsp;&nbsp;&nbsp;แนะแนว', ac(a.guid))}${sumRow('&nbsp;&nbsp;&nbsp;ลูกเสือ-เนตรนารี', ac(a.scout))}
        ${sumRow('&nbsp;&nbsp;&nbsp;ชุมนุม ' + esc(a.clubName), ac(a.club))}${sumRow('&nbsp;&nbsp;&nbsp;กิจกรรมเพื่อสังคมและสาธารณประโยชน์', ac(a.social))}
        ${sumRow('ผลการประเมินคุณลักษณะอันพึงประสงค์', lv(e.t, C.LV_TRAIT))}${sumRow('ผลการประเมินการอ่าน คิดวิเคราะห์ และเขียน', lv(e.r, C.LV_TRAIT))}
        ${sumRow('ผลการประเมินค่านิยมหลัก 12 ประการ', lv(e.v, C.LV_TRAIT))}${sumRow('ผลการประเมินสมรรถนะสำคัญของผู้เรียน', lv(e.c, C.LV_COMP))}
      </table>
      <div style="flex:1;padding:3mm 2mm;border-left:.6pt solid #000">
        ${sigBlock(room.advisor, 'ครูประจำชั้น/ครูที่ปรึกษา')}${sigBlock(school.academic, 'หัวหน้างานวิชาการโรงเรียน')}${sigBlock(school.director, school.directorTitle || 'ผู้อำนวยการโรงเรียน')}
        <div class="c t-sm">ลงชื่อ${dl('', 55)}<br>(${dl('', 50)})<br>ผู้ปกครองนักเรียน</div></div>
    </div></div></div>`;
}

function p6Health(room, s, term) {
  const dates = room.health?.dates ?? [], h = room.health?.of?.[s.id] ?? {};
  const blocks = term === 'รวม' ? [1, 2] : [term];
  const hb = t => {
    const o = (t - 1) * 2;
    return `<tr><td rowspan="2" class="b" style="width:42mm">น้ำหนัก - ส่วนสูง</td><td colspan="2" class="b">ภาคเรียนที่ ${t}</td></tr>
      <tr><td class="b">ครั้งที่ 1</td><td class="b">ครั้งที่ 2</td></tr>
      <tr><td class="b">วันที่ชั่งน้ำหนัก/วัดส่วนสูง</td><td>${C.thaiDate(dates[o])}</td><td>${C.thaiDate(dates[o + 1])}</td></tr>
      <tr><td class="b">น้ำหนัก (กิโลกรัม)</td><td>${esc(h.w?.[o] ?? '')}</td><td>${esc(h.w?.[o + 1] ?? '')}</td></tr>
      <tr><td class="b">ส่วนสูง (เซนติเมตร)</td><td>${esc(h.h?.[o] ?? '')}</td><td>${esc(h.h?.[o + 1] ?? '')}</td></tr>`;
  };
  const days = room.att?.days ?? {}, p = room.att?.present?.[s.id] ?? {};
  let F = 0, P = 0;
  const arows = C.MONTHS.map(m => {
    const f = Number(days[m]) || 0, q = p[m] === '' || p[m] == null ? null : Number(p[m]);
    F += f; P += q || 0;
    return `<tr><td>${m}</td><td>${f || ''}</td><td>${q ?? ''}</td><td>${f && q != null ? C.fmt(q * 100 / f) : '-'}</td><td></td></tr>`;
  }).join('');
  return `<div class="page"><div style="padding:4mm;font-size:12pt">
    <div class="c t-md" style="margin-bottom:3mm">ผลการประเมินภาวะโภชนาการ</div>
    <table class="pt" style="table-layout:auto;font-size:11.5pt">${blocks.map(hb).join('')}</table>
    <div class="c t-md" style="margin:9mm 0 3mm">สรุปเวลาเรียน</div>
    <table class="pt" style="font-size:11.5pt"><colgroup><col style="width:42mm"><col style="width:25mm"><col style="width:25mm"><col style="width:25mm"><col></colgroup>
    <tr><th rowspan="2" style="height:12mm">เดือน</th><th>เวลาเต็ม<br>(วัน)</th><th>เวลามา<br>(วัน)</th><th>คิดเป็น<br>ร้อยละ</th><th rowspan="2">หมายเหตุ</th></tr><tr style="display:none"><td></td></tr>
    ${arows}<tr><td class="b" style="height:10mm">รวม</td><td class="b">${F || ''}</td><td class="b">${P || ''}</td><td class="b">${F ? C.fmt(P * 100 / F) : ''}</td><td></td></tr></table></div></div>`;
}

function p6Remarks(room, s, term) {
  const t = term === 'รวม' ? 2 : term;
  const r = room.remarks?.[t]?.[s.id] ?? {};
  const topics = ['ด้านหน้าที่รับผิดชอบ<br>ความเอาใจใส่การเรียน', 'ด้านการใช้เวลาว่าง', 'ด้านความสัมพันธ์กับ<br>บุคคลรอบข้าง', 'ด้านอุปนิสัย บุคลิกภาพ', 'ด้านสุขภาพ'];
  const box = (title, vals, sign) => `<div class="c t-md" style="margin:2mm 0 2mm">${title}</div>
    <table class="pt" style="table-layout:auto;font-size:11.5pt"><colgroup><col style="width:46mm"><col></colgroup>${topics.map((tp, i) => `<tr><td class="b" style="height:${vals ? 20 : 19}mm">${tp}</td><td class="l" style="vertical-align:top;${vals ? '' : 'background-image:linear-gradient(transparent 8.9mm,#000 8.9mm,#000 9.3mm,transparent 9.3mm);background-size:100% 9.5mm;'}">${vals ? esc(vals[i] ?? '') : ''}</td></tr>`).join('')}</table>
    <div class="c t-sm" style="margin-top:3mm">${sign}</div>`;
  return `<div class="page"><div style="padding:4mm">
    ${box('ความคิดเห็นและข้อเสนอแนะของครูประจำชั้น', [r.a, r.b, r.c, r.d, r.e], 'ลงชื่อ...........................................................ครูประจำชั้น/ครูที่ปรึกษา')}
    <div style="height:7mm"></div>
    ${box('ความคิดเห็นและข้อเสนอแนะของผู้ปกครอง', null, 'ลงชื่อ...........................................................ผู้ปกครอง<br>(...........................................................)')}
  </div></div>`;
}

/* ---------- ตัวเลือกเอกสาร ---------- */
export const DOCS = [
  { id: 'p5cover', label: 'ปก ปพ.5 รายวิชา', needSubj: true, needTerm: true },
  { id: 'p5class', label: 'ปก ปพ.5 รายชั้น (สรุปทุกรายวิชา)', needTerm: true },
  { id: 'ind', label: 'ผลประเมินมาตรฐาน/ตัวชี้วัด', needSubj: true, needTerm: 'one' },
  { id: 'score', label: 'ผลสัมฤทธิ์ทางการเรียน (คะแนน)', needSubj: true, needTerm: true },
  { id: 'traits', label: 'คุณลักษณะฯ + การอ่าน คิดวิเคราะห์ เขียน', needTerm: true },
  { id: 'values', label: 'ค่านิยม 12 ประการ + สมรรถนะ', needTerm: true },
  { id: 'acts', label: 'กิจกรรมพัฒนาผู้เรียน', needTerm: true },
  { id: 'attend', label: 'สรุปเวลาเรียน' },
  { id: 'health', label: 'น้ำหนัก - ส่วนสูง' },
  { id: 'p6', label: 'ปพ.6 รายบุคคล (เล่ม)', needTerm: true, needStudent: true },
];

export function buildDoc(opts, room, school) {
  const { doc, term, subjId, studentId } = opts;
  const subj = (room.subjects ?? []).find(s => s.id === subjId) ?? room.subjects?.[0];
  const t = term ?? 1;
  switch (doc) {
    case 'p5cover': return subj ? coverSubject(room, school, subj, t) : '';
    case 'p5class': return coverClass(room, school, t);
    case 'ind': return subj ? indPage(room, subj, t === 'รวม' ? 2 : t) : '';
    case 'score': return subj ? scorePage(room, subj, t) : '';
    case 'traits': return evalSheet(room, t, 't');
    case 'values': return evalSheet(room, t, 'v');
    case 'acts': return actsSheet(room, t);
    case 'attend': return attendSheet(room);
    case 'health': return healthSheet(room);
    case 'p6': {
      const totals = studentTotals(room, t);
      const list = studentId === 'all' ? sortStudents(room) : sortStudents(room).filter(s => s.id === studentId);
      return list.map(s => p6Cover(room, school, s, t) + p6Advice() + p6Info(room, school, s) + p6Results(room, school, s, t, totals) + p6Health(room, s, t) + p6Remarks(room, s, t)).join('');
    }
  }
  return '';
}
