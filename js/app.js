import { store } from './store.js';
import * as C from './calc.js';
import { DOCS, buildDoc, esc } from './print.js';

const $ = (s, r = document) => r.querySelector(s);
const $$ = (s, r = document) => [...r.querySelectorAll(s)];
const uid = () => (crypto.randomUUID ? crypto.randomUUID().slice(0, 8) : Math.random().toString(36).slice(2, 10));
const S = { user: null, school: {}, rooms: [], room: null, ui: {} };
const app = $('#app');

/* ---------- ยูทิลิตี้ ---------- */
function toast(msg) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toast.t); toast.t = setTimeout(() => t.classList.remove('on'), 2200);
}
const thaiYearNow = () => { const d = new Date(); return d.getMonth() >= 4 ? d.getFullYear() + 543 : d.getFullYear() + 542; };
const sorted = () => [...S.room.students].sort((a, b) => (a.no || 0) - (b.no || 0));
const activeSt = () => sorted();

function setState(txt, cls) { const el = $('#savestate'); if (el) { el.textContent = txt; el.className = 'savestate ' + (cls || ''); } }
function touch() {
  setState('กำลังบันทึก…');
  store.saveRoom(S.room).then(() => setState('บันทึกแล้ว ✓'), e => { console.error(e); setState('บันทึกไม่สำเร็จ: ' + (e.message || e), 'err'); });
}
function touchSchool() {
  store.saveSchool(S.school).then(() => toast('บันทึกข้อมูลโรงเรียนแล้ว'), e => toast('บันทึกไม่สำเร็จ: ' + e.message));
}

function ask(title, body, { ok = 'บันทึก', cancel = 'ยกเลิก', onOpen } = {}) {
  const d = $('#dlg');
  d.innerHTML = `<form method="dialog"><div class="dlg-h"><span>${title}</span><button class="btn sm" value="x" formnovalidate>✕</button></div><div class="dlg-b">${body}</div><div class="dlg-f">${cancel ? `<button class="btn" value="x" formnovalidate>${cancel}</button>` : ''}<button class="btn primary" value="ok">${ok}</button></div></form>`;
  return new Promise(res => {
    d.onclose = () => res(d.returnValue === 'ok' ? d : null);
    d.showModal();
    onOpen?.(d);
  });
}

/* ---------- เริ่มต้น ---------- */
async function boot() {
  try { S.user = await store.user(); } catch (e) { console.error(e); }
  window.addEventListener('hashchange', route);
  route();
}

async function route() {
  if (!S.user) return loginView();
  const [, a, b, c] = location.hash.split('/');
  if (a === 'room' && b) return roomRoute(b, c || 'info');
  if (a === 'settings') return settingsView();
  S.room = null;
  return roomsView();
}

function shell(inner, { back, title, tabs } = {}) {
  app.innerHTML = `<div class="topbar">
    <a class="brand" href="#/">ปพ.5 · ปพ.6<small>หลักสูตร 2551 ประถมศึกษา</small></a>
    ${store.remote ? '' : '<span class="badge warn" title="ข้อมูลเก็บในเบราว์เซอร์นี้เท่านั้น ตั้งค่า Supabase ใน js/config.js เพื่อใช้งานจริง">โหมดทดลอง</span>'}
    <span class="grow"></span><span id="savestate" class="savestate"></span>
    <a class="btn sm" href="#/settings">ข้อมูลโรงเรียน</a>
    ${store.remote ? '<button class="btn sm" id="logout">ออกจากระบบ</button>' : ''}</div>${tabs || ''}<main class="wrap" id="main">${inner}</main>`;
  $('#logout')?.addEventListener('click', async () => { await store.signOut(); S.user = null; route(); });
}

/* ---------- เข้าสู่ระบบ ---------- */
function loginView() {
  app.innerHTML = `<div class="login"><div class="card"><h1>ปพ.5 · ปพ.6 ออนไลน์</h1><p class="muted" style="margin:0">หลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน 2551 · ระดับประถมศึกษา</p>
    <label class="f">อีเมล<input type="email" id="em" autocomplete="username"></label>
    <label class="f">รหัสผ่าน<input type="password" id="pw" autocomplete="current-password"></label>
    <div class="err" id="er"></div>
    <div class="row"><button class="btn primary grow" id="in">เข้าสู่ระบบ</button><button class="btn" id="up">สมัครใช้งาน</button></div></div></div>`;
  const go = async fn => {
    $('#er').textContent = '';
    try { const r = await fn($('#em').value.trim(), $('#pw').value); if (r === false) { $('#er').textContent = 'สมัครแล้ว กรุณายืนยันอีเมลก่อนเข้าสู่ระบบ'; return; } S.user = await store.user(); route(); }
    catch (e) { $('#er').textContent = e.message; }
  };
  $('#in').onclick = () => go(store.signIn);
  $('#up').onclick = () => go(store.signUp);
  $('#pw').onkeydown = e => { if (e.key === 'Enter') $('#in').click(); };
}

/* ---------- รายการห้องเรียน ---------- */
async function roomsView() {
  shell('<p class="muted">กำลังโหลด…</p>');
  try { S.school = await store.loadSchool(); S.rooms = await store.listRooms(); } catch (e) { $('#main').innerHTML = `<p class="err">โหลดข้อมูลไม่สำเร็จ: ${esc(e.message)}</p>`; return; }
  const rooms = S.rooms;
  $('#main').innerHTML = `<div class="row" style="margin-bottom:1rem"><h1 class="grow" style="margin:0">ห้องเรียนของฉัน</h1><button class="btn primary" id="new">+ สร้างห้องเรียน</button></div>
  ${S.school.name ? '' : '<div class="card" style="margin-bottom:1rem">ยังไม่ได้กรอกข้อมูลโรงเรียน — <a href="#/settings">กรอกข้อมูลโรงเรียนก่อน</a> เพื่อให้แสดงบนหน้าพิมพ์</div>'}
  ${rooms.length ? `<div class="grid-rooms">${rooms.map(r => `<a class="room-card" href="#/room/${r.id}/students"><b>${esc(r.title)}</b><div class="muted">ปีการศึกษา ${esc(r.year)}</div></a>`).join('')}</div>` : '<p class="muted">ยังไม่มีห้องเรียน กด “สร้างห้องเรียน” เพื่อเริ่มต้น</p>'}`;
  $('#new').onclick = async () => {
    const f = await ask('สร้างห้องเรียน', `<div class="form-grid">
      <label class="f">ระดับชั้น<select name="level">${C.LEVELS.map(l => `<option>${l}</option>`).join('')}</select></label>
      <label class="f">ห้อง<input type="text" name="room" value="1"></label>
      <label class="f">ปีการศึกษา (พ.ศ.)<input type="number" name="year" value="${thaiYearNow()}"></label>
      <label class="f span2">ครูประจำชั้น/ครูที่ปรึกษา<input type="text" name="advisor"></label></div>`, { ok: 'สร้าง' });
    if (!f) return;
    const v = Object.fromEntries(new FormData(f.querySelector('form')));
    const room = { title: `${v.level}/${v.room}`, year: +v.year, level: v.level, room: v.room, advisor: v.advisor, signDate: '',
      students: [], subjects: [], scores: {}, eval: { 1: {}, 2: {} }, acts: { 1: {}, 2: {} }, att: { days: {}, present: {} }, health: { dates: [], of: {} }, remarks: { 1: {}, 2: {} } };
    try { await store.createRoom(room); location.hash = `#/room/${room.id}/students`; } catch (e) { toast('สร้างไม่สำเร็จ: ' + e.message); }
  };
}

/* ---------- ข้อมูลโรงเรียน ---------- */
async function settingsView() {
  shell('<p class="muted">กำลังโหลด…</p>');
  try { S.school = await store.loadSchool(); } catch (e) { toast(e.message); }
  const F = [['name', 'ชื่อโรงเรียน (ไม่ต้องใส่คำว่า “โรงเรียน”)'], ['affiliation', 'สังกัด (ระดับจังหวัด/เขตพื้นที่การศึกษา)'], ['subdistrict', 'ตำบล/แขวง'], ['district', 'อำเภอ/เขต'], ['province', 'จังหวัด'], ['director', 'ชื่อ-สกุล ผู้อำนวยการ'], ['directorTitle', 'ตำแหน่ง (เช่น ผู้อำนวยการโรงเรียน / รักษาการ…)'], ['registrar', 'ชื่อ-สกุล นายทะเบียน'], ['academic', 'ชื่อ-สกุล หัวหน้างานวิชาการ']];
  $('#main').innerHTML = `<h1>ข้อมูลโรงเรียน</h1><div class="card"><div class="form-grid">${F.map(([k, l]) => `<label class="f">${l}<input type="text" data-k="${k}" value="${esc(S.school[k] ?? '')}"></label>`).join('')}</div><div class="row" style="margin-top:1rem"><button class="btn primary" id="sv">บันทึก</button><a class="btn" href="#/">กลับ</a></div></div>`;
  $('#sv').onclick = () => { $$('[data-k]').forEach(i => (S.school[i.dataset.k] = i.value.trim())); touchSchool(); };
}

/* ---------- พื้นที่ทำงานของห้อง ---------- */
const TABS = [['students', 'นักเรียน'], ['subjects', 'รายวิชา'], ['scores', 'คะแนน/ตัวชี้วัด'], ['eval', 'คุณลักษณะ/ค่านิยม'], ['acts', 'กิจกรรม'], ['attend', 'เวลาเรียน'], ['health', 'น้ำหนัก-ส่วนสูง'], ['remarks', 'ความเห็นครู'], ['print', 'พิมพ์เอกสาร'], ['info', 'ข้อมูลห้อง']];

async function roomRoute(id, tab) {
  if (S.room?.id !== id) {
    shell('<p class="muted">กำลังโหลด…</p>');
    try { S.school = await store.loadSchool(); S.room = await store.loadRoom(id); } catch (e) { $('#main').innerHTML = `<p class="err">${esc(e.message)}</p>`; return; }
    if (!S.room) { $('#main').innerHTML = '<p class="err">ไม่พบห้องเรียน</p>'; return; }
    for (const k of ['students', 'subjects']) S.room[k] ??= [];
    for (const k of ['scores', 'att', 'health', 'remarks', 'eval', 'acts']) S.room[k] ??= {};
    for (const k of ['eval', 'acts', 'remarks']) { S.room[k][1] ??= {}; S.room[k][2] ??= {}; }
    S.room.att.days ??= {}; S.room.att.present ??= {}; S.room.health.dates ??= []; S.room.health.of ??= {};
  }
  const tabs = `<nav class="tabs"><a href="#/" title="ห้องเรียนทั้งหมด">☰</a>${TABS.map(([k, l]) => `<a href="#/room/${id}/${k}" class="${k === tab ? 'on' : ''}">${l}</a>`).join('')}</nav>`;
  shell('', { tabs });
  const v = { students: studentsView, subjects: subjectsView, scores: scoresView, eval: evalView, acts: actsView, attend: attendView, health: healthView, remarks: remarksView, print: printView, info: infoView }[tab] ?? infoView;
  $('#main').innerHTML = `<div class="row" style="margin-bottom:.6rem"><h2 style="margin:0">${esc(S.room.title)}</h2><span class="muted">ปีการศึกษา ${esc(S.room.year)}</span></div><div id="view"></div>`;
  v($('#view'));
}

/* ---- ตัวช่วยตาราง ---- */
function gridNav(root) {
  root.addEventListener('keydown', e => {
    const el = e.target; if (el.tagName !== 'INPUT') return;
    const td = el.closest('td'), tr = td?.parentElement; if (!tr) return;
    let tgt = null;
    const col = td.cellIndex;
    const mv = r => r?.cells[col]?.querySelector('input');
    if (e.key === 'Enter' || e.key === 'ArrowDown') tgt = mv(tr.nextElementSibling);
    else if (e.key === 'ArrowUp') tgt = mv(tr.previousElementSibling);
    else if (e.key === 'ArrowRight' && el.selectionStart === el.value.length) tgt = td.nextElementSibling?.querySelector('input');
    else if (e.key === 'ArrowLeft' && el.selectionEnd === 0) tgt = td.previousElementSibling?.querySelector('input');
    if (tgt) { e.preventDefault(); tgt.focus(); tgt.select(); }
  });
}
const termSel = (cur, opts = [1, 2]) => `<label class="f">ภาคเรียน<select id="term">${opts.map(t => `<option value="${t}" ${String(t) === String(cur) ? 'selected' : ''}>${t === 'รวม' ? 'สรุปรายปี' : 'ภาคเรียนที่ ' + t}</option>`).join('')}</select></label>`;

/* ---------- นักเรียน ---------- */
const SFIELDS = [['no', 'เลขที่', 'number'], ['sid', 'รหัสนักเรียน'], ['cid', 'เลขประจำตัวประชาชน'], ['prefix', 'คำนำหน้า', ['เด็กชาย', 'เด็กหญิง', 'นาย', 'นางสาว']], ['first', 'ชื่อ'], ['last', 'นามสกุล'], ['sex', 'เพศ', ['ช', 'ญ']], ['birth', 'วันเกิด', 'date'], ['status', 'สถานะ', ['ปกติ', 'ย้ายเข้า', 'ย้ายออก', 'ลาออก']],
  ['nation', 'สัญชาติ'], ['ethnic', 'เชื้อชาติ'], ['religion', 'ศาสนา'], ['father', 'ชื่อ-สกุลบิดา'], ['fatherJob', 'อาชีพบิดา'], ['mother', 'ชื่อ-สกุลมารดา'], ['motherJob', 'อาชีพมารดา'], ['marital', 'สถานภาพสมรสของบิดา-มารดา'], ['guardian', 'ชื่อ-สกุลผู้ปกครอง'], ['guardianJob', 'อาชีพผู้ปกครอง'], ['relation', 'ความเกี่ยวข้องกับนักเรียน'], ['phone', 'โทรศัพท์ผู้ปกครอง'],
  ['addrNo', 'บ้านเลขที่'], ['addrMoo', 'หมู่'], ['addrRoad', 'ถนน'], ['addrSub', 'ตำบล'], ['addrDist', 'อำเภอ'], ['addrProv', 'จังหวัด']];

function studentForm(s) {
  return `<div class="form-grid">${SFIELDS.map(([k, l, t]) => {
    const v = s[k] ?? '';
    if (Array.isArray(t)) return `<label class="f">${l}<select name="${k}"><option value=""></option>${t.map(o => `<option ${o === v ? 'selected' : ''}>${o}</option>`).join('')}</select></label>`;
    return `<label class="f">${l}<input type="${t || 'text'}" name="${k}" value="${esc(v)}"></label>`;
  }).join('')}</div>`;
}

function studentsView(root) {
  const R = S.room;
  const draw = () => {
    root.innerHTML = `<div class="row" style="margin-bottom:.8rem"><button class="btn primary" id="add">+ เพิ่มนักเรียน</button><button class="btn" id="imp">นำเข้า (วางจาก Excel)</button><span class="grow"></span><span class="muted">${R.students.length} คน</span></div>
    <div class="gridwrap"><table class="g"><thead><tr><th>เลขที่</th><th>รหัสนักเรียน</th><th style="text-align:left">ชื่อ-สกุล</th><th>เพศ</th><th>สถานะ</th><th></th></tr></thead><tbody>
    ${sorted().map(s => `<tr class="${s.status === 'ย้ายออก' || s.status === 'ลาออก' ? 'gone' : ''}"><td class="no">${s.no ?? ''}</td><td class="out">${esc(s.sid)}</td><td class="stk" style="position:static;text-align:left">${esc(C.studentName(s))}</td><td class="out">${esc(s.sex)}</td><td class="out">${esc(s.status || 'ปกติ')}</td><td class="out"><button class="btn sm" data-ed="${s.id}">แก้ไข</button> <button class="btn sm danger" data-del="${s.id}">ลบ</button></td></tr>`).join('') || '<tr><td colspan="6" class="muted" style="padding:1rem">ยังไม่มีนักเรียน</td></tr>'}</tbody></table></div>
    <p class="hint">กรณีนักเรียนย้ายออก/ลาออก ให้เปลี่ยน “สถานะ” แทนการลบ เพื่อให้สรุปจำนวนนักเรียนในหน้าปก ปพ.5 ถูกต้อง</p>`;
    $('#add').onclick = () => edit(null);
    $('#imp').onclick = importStudents;
    $$('[data-ed]', root).forEach(b => b.onclick = () => edit(b.dataset.ed));
    $$('[data-del]', root).forEach(b => b.onclick = async () => {
      const s = R.students.find(x => x.id === b.dataset.del);
      if (!confirm(`ลบ ${C.studentName(s)} ? (คะแนนของนักเรียนคนนี้จะถูกลบด้วย)`)) return;
      R.students = R.students.filter(x => x.id !== s.id);
      for (const sub of Object.values(R.scores)) delete sub[s.id];
      touch(); draw();
    });
  };
  async function edit(id) {
    const s = id ? R.students.find(x => x.id === id) : { id: uid(), no: Math.max(0, ...R.students.map(x => x.no || 0)) + 1, status: 'ปกติ', nation: 'ไทย', ethnic: 'ไทย', religion: 'พุทธ' };
    const f = await ask(id ? 'แก้ไขนักเรียน' : 'เพิ่มนักเรียน', studentForm(s));
    if (!f) return;
    const v = Object.fromEntries(new FormData(f.querySelector('form')));
    Object.assign(s, v); s.no = +v.no || s.no;
    if (!s.sex) s.sex = /ชาย|นาย$/.test(s.prefix) || s.prefix === 'นาย' ? 'ช' : s.prefix ? 'ญ' : '';
    if (!id) R.students.push(s);
    touch(); draw();
  }
  async function importStudents() {
    const f = await ask('นำเข้านักเรียน', `<p class="hint">คัดลอกจาก Excel แล้ววางที่นี่ ลำดับคอลัมน์: <b>รหัสนักเรียน · คำนำหน้า · ชื่อ · นามสกุล · เพศ(ช/ญ) · วันเกิด · เลขประชาชน</b> (ใส่เฉพาะที่มี ปล่อยว่างได้)<br>หรือวางเฉพาะชื่อ บรรทัดละคน เช่น <i>เด็กชายสมชาย ใจดี</i></p><textarea name="t" rows="12" style="width:100%" placeholder="เลขที่จะรันต่อจากคนสุดท้าย"></textarea>`, { ok: 'นำเข้า' });
    if (!f) return;
    const lines = new FormData(f.querySelector('form')).get('t').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    let n = Math.max(0, ...R.students.map(x => x.no || 0)), c = 0;
    for (const ln of lines) {
      let s;
      if (ln.includes('\t')) {
        const [sid, prefix, first, last, sex, birth, cid] = ln.split('\t').map(x => x.trim());
        s = { sid, prefix, first, last, sex, birth: C.parseDate(birth), cid };
      } else {
        const m = ln.match(/^(เด็กชาย|เด็กหญิง|นางสาว|นาย|ด\.ช\.|ด\.ญ\.|น\.ส\.)?\s*(\S+)\s+(.+)$/);
        if (!m) continue;
        const pf = { 'ด.ช.': 'เด็กชาย', 'ด.ญ.': 'เด็กหญิง', 'น.ส.': 'นางสาว' }[m[1]] ?? m[1] ?? '';
        s = { prefix: pf, first: m[2], last: m[3] };
      }
      if (!s.first) continue;
      s.sex ||= s.prefix === 'เด็กชาย' || s.prefix === 'นาย' ? 'ช' : s.prefix ? 'ญ' : '';
      Object.assign(s, { id: uid(), no: ++n, status: 'ปกติ', nation: 'ไทย', ethnic: 'ไทย', religion: 'พุทธ' });
      R.students.push(s); c++;
    }
    touch(); draw(); toast(`นำเข้า ${c} คน`);
  }
  draw();
}

/* ---------- รายวิชา ---------- */
function subjectsView(root) {
  const R = S.room;
  const lv = Math.max(1, C.LEVELS.indexOf(R.level) + 1);
  const draw = () => {
    root.innerHTML = `<div class="row" style="margin-bottom:.8rem"><button class="btn primary" id="add">+ เพิ่มรายวิชา</button><button class="btn" id="base">เพิ่มรายวิชาพื้นฐาน 8 กลุ่มสาระ</button></div>
    <div class="gridwrap"><table class="g"><thead><tr><th>ที่</th><th>รหัสวิชา</th><th>รายวิชา</th><th>เวลาเรียน (ชม./ปี)</th><th>กลุ่มสาระการเรียนรู้</th><th>ครูประจำวิชา</th><th></th></tr></thead><tbody>
    ${R.subjects.map((s, i) => `<tr data-i="${i}"><td class="no">${i + 1}</td><td><input type="text" data-f="code" value="${esc(s.code)}"></td><td><input type="text" class="w" data-f="name" value="${esc(s.name)}"></td><td><input type="number" data-f="hours" value="${esc(s.hours)}"></td>
    <td><select data-f="group" style="border:0">${['', ...C.GROUPS].map(g => `<option ${g === s.group ? 'selected' : ''}>${g}</option>`).join('')}</select></td><td><input type="text" class="w" data-f="teacher" value="${esc(s.teacher)}"></td>
    <td class="out"><button class="btn sm" data-up="${i}">↑</button> <button class="btn sm danger" data-del="${i}">ลบ</button></td></tr>`).join('') || '<tr><td colspan="7" class="muted" style="padding:1rem">ยังไม่มีรายวิชา</td></tr>'}</tbody></table></div>
    <p class="hint">ใส่รายวิชาได้ไม่จำกัด (ไฟล์ Excel เดิมรองรับ 20 วิชา) เวลาเรียนใช้แสดงในหน้าปก ปพ.5 และคำนวณผลการเรียนเฉลี่ย (ถ่วงน้ำหนักตามชั่วโมง)</p>`;
    $$('[data-f]', root).forEach(el => el.addEventListener('input', () => { const s = R.subjects[+el.closest('tr').dataset.i]; s[el.dataset.f] = el.value; touch(); }));
    $('#add').onclick = () => { R.subjects.push({ id: uid(), code: '', name: '', hours: '', group: '', teacher: R.advisor || '' }); touch(); draw(); };
    $('#base').onclick = () => {
      const base = [['ท', 'ภาษาไทย', 'ภาษาไทย', 200], ['ค', 'คณิตศาสตร์', 'คณิตศาสตร์', 200], ['ว', 'วิทยาศาสตร์และเทคโนโลยี', 'วิทยาศาสตร์และเทคโนโลยี', 80], ['ส', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80], ['พ', 'สุขศึกษาและพลศึกษา', 'สุขศึกษาและพลศึกษา', 80], ['ศ', 'ศิลปะ', 'ศิลปะ', 80], ['ง', 'การงานอาชีพ', 'การงานอาชีพและเทคโนโลยี', 40], ['อ', 'ภาษาอังกฤษ', 'ภาษาต่างประเทศ', 80]];
      for (const [p, n, g, h] of base) if (!R.subjects.some(s => s.name === n)) R.subjects.push({ id: uid(), code: `${p}1${lv}101`, name: n, hours: h, group: g, teacher: R.advisor || '' });
      touch(); draw(); toast('เพิ่มแล้ว — ตรวจสอบรหัสวิชาและชั่วโมงให้ตรงกับหลักสูตรสถานศึกษา');
    };
    $$('[data-del]', root).forEach(b => b.onclick = () => { const s = R.subjects[+b.dataset.del]; if (confirm(`ลบวิชา ${s.name || '(ไม่มีชื่อ)'} ? คะแนนของวิชานี้จะถูกลบด้วย`)) { delete R.scores[s.id]; R.subjects.splice(+b.dataset.del, 1); touch(); draw(); } });
    $$('[data-up]', root).forEach(b => b.onclick = () => { const i = +b.dataset.up; if (i > 0) { [R.subjects[i - 1], R.subjects[i]] = [R.subjects[i], R.subjects[i - 1]]; touch(); draw(); } });
  };
  draw();
}

/* ---------- คะแนน / ตัวชี้วัด ---------- */
function scoresView(root) {
  const R = S.room, U = S.ui;
  U.subj = R.subjects.find(s => s.id === U.subj)?.id ?? R.subjects[0]?.id;
  U.term ??= 1; U.smode ??= 'score';
  if (!R.subjects.length) { root.innerHTML = '<p class="muted">เพิ่มรายวิชาในแท็บ “รายวิชา” ก่อน</p>'; return; }
  const subj = R.subjects.find(s => s.id === U.subj);
  const cfg = C.ensureCfg(subj, U.term);
  const rec = sid => ((R.scores[subj.id] ??= {})[sid] ??= {})[U.term] ??= { p: [], s: [], f: '', o: '' };
  const st = sorted();
  const isS = U.smode === 'score';

  const head = `<div class="printbar"><label class="f">รายวิชา<select id="sj">${R.subjects.map(s => `<option value="${s.id}" ${s.id === subj.id ? 'selected' : ''}>${esc(s.code)} ${esc(s.name)}</option>`).join('')}</select></label>${termSel(U.term)}
    <div class="row"><button class="btn ${isS ? 'primary' : ''}" data-m="score">คะแนน</button><button class="btn ${!isS ? 'primary' : ''}" data-m="ind">ตัวชี้วัด (ผ/มผ)</button></div>
    <div class="row">${isS ? `<label class="f">อัตราส่วน ระหว่างภาค : ปลายภาค<span class="row"><input type="number" id="rb" value="${cfg.rb}" style="width:4.5rem"> : <input type="number" id="rf" value="${cfg.rf}" style="width:4.5rem"></span></label>` : ''}
    <button class="btn sm" id="cadd">+ คอลัมน์</button><button class="btn sm" id="cdel">− คอลัมน์</button><button class="btn sm" id="copycfg">คัดลอกโครงสร้างไปอีกภาค</button></div></div>`;

  const cols = isS ? cfg.max : cfg.inds;
  let thead, rowsHtml;
  if (isS) {
    thead = `<tr><th rowspan="2">เลขที่</th><th rowspan="2" class="stk">ชื่อ-สกุล</th><th colspan="${cols.length}">คะแนนระหว่างเรียน (ใส่คะแนนเต็มในช่องล่าง)</th><th rowspan="2">รวม</th><th rowspan="2">ระหว่างภาค<br><small>(${cfg.rb})</small></th><th rowspan="2">ปลายภาค<br><small>(${cfg.rf})</small></th><th rowspan="2">รวม<br><small>(${cfg.rb + cfg.rf})</small></th><th rowspan="2">ร้อยละ</th><th rowspan="2">เกรด</th><th rowspan="2">ร/มส</th></tr>
      <tr>${cols.map((m, k) => `<th><input type="number" data-max="${k}" value="${m}" style="width:3.2rem;border:1px solid var(--line);border-radius:6px;text-align:center"></th>`).join('')}</tr>`;
  } else {
    thead = `<tr><th rowspan="2">เลขที่</th><th rowspan="2" class="stk">ชื่อ-สกุล</th><th colspan="${cols.length}">ตัวชี้วัด (กรอก ผ หรือ มผ — พิมพ์ 1 = ผ่าน, 0 = ไม่ผ่าน)</th><th rowspan="2">ผ่าน</th><th rowspan="2">ไม่ผ่าน</th><th rowspan="2">สรุป</th></tr>
      <tr>${cols.map((n, k) => `<th><input type="text" data-iname="${k}" value="${esc(n)}" placeholder="${k + 1}" style="width:3.4rem;border:1px solid var(--line);border-radius:6px;text-align:center"></th>`).join('')}</tr>`;
  }
  rowsHtml = st.map((s, i) => `<tr data-sid="${s.id}"><td class="no">${s.no ?? i + 1}</td><td class="stk">${esc(C.studentName(s))}</td>${cols.map((_, k) => `<td><input type="text" data-k="${k}" inputmode="${isS ? 'decimal' : 'text'}" ${isS ? '' : 'maxlength="2"'}></td>`).join('')}
    ${isS ? `<td class="out" data-o="sum"></td><td class="out" data-o="between"></td><td><input type="text" data-f="f" inputmode="decimal"></td><td class="out" data-o="total"></td><td class="out" data-o="pct"></td><td class="out" data-o="grade"></td><td><input type="text" data-f="o" maxlength="2" style="width:2.6rem" placeholder="-"></td>`
      : `<td class="out" data-o="pass"></td><td class="out" data-o="fail"></td><td class="out" data-o="ind"></td>`}</tr>`).join('');

  root.innerHTML = head + `<div class="gridwrap"><table class="g"><thead>${thead}</thead><tbody>${rowsHtml}</tbody></table></div>
    <p class="hint">${isS ? 'ช่อง “ปลายภาค” กรอกเป็นคะแนนเต็ม ' + cfg.rf + ' · ช่อง ร/มส ใช้กรณีรอการตัดสิน/ไม่มีสิทธิ์สอบ · ผลลัพธ์ในช่องสีเทาคำนวณอัตโนมัติ (ร้อยละ → เกรด 0-4 ตามเกณฑ์ 2551)' : 'ผลสรุป: ไม่ผ่านตัวชี้วัดใดตัวชี้วัดหนึ่ง = มผ · ผ่านทุกตัวที่ประเมิน = ผ'} · ใช้ปุ่มลูกศร/Enter เลื่อนช่อง</p>`;

  const refresh = tr => {
    const sid = tr.dataset.sid, r = C.termResult(subj, R.scores[subj.id]?.[sid], U.term);
    const set = (k, v) => { const c = $(`[data-o="${k}"]`, tr); if (c) c.textContent = v ?? ''; };
    if (isS) { set('sum', C.fmt(r.sum)); set('between', C.fmt(r.between)); set('total', C.fmt(r.total)); set('pct', C.fmt(r.pct)); set('grade', r.grade); }
    else { set('pass', r.ind ? r.pass : ''); set('fail', r.ind ? r.fail : ''); set('ind', r.ind); }
  };
  const normPF = v => { v = v.trim(); if (!v) return ''; if (v.startsWith('ม')) return 'มผ'; if (v.includes('ผ') || v === '1') return 'ผ'; if (v === '0') return 'มผ'; return ''; };

  $$('tbody tr', root).forEach(tr => {
    const sid = tr.dataset.sid, rc = (R.scores[subj.id]?.[sid]?.[U.term]) ?? { p: [], s: [], f: '', o: '' };
    $$('input[data-k]', tr).forEach(inp => {
      const k = +inp.dataset.k; inp.value = (isS ? rc.s[k] : rc.p[k]) ?? '';
      inp.addEventListener('input', () => {
        const r = rec(sid);
        if (isS) { const v = inp.value.replace(/[^\d.]/g, ''); if (v !== inp.value) inp.value = v; r.s[k] = v; inp.classList.toggle('bad', C.isNum(v) && +v > +cfg.max[k]); }
        else { const v = normPF(inp.value); inp.value = v; r.p[k] = v; }
        touch(); refresh(tr);
      });
      if (!isS) inp.addEventListener('focus', () => inp.select());
    });
    if (isS) {
      const fi = $('[data-f="f"]', tr); fi.value = rc.f ?? '';
      fi.addEventListener('input', () => { const v = fi.value.replace(/[^\d.]/g, ''); fi.value = v; rec(sid).f = v; fi.classList.toggle('bad', C.isNum(v) && +v > cfg.rf); touch(); refresh(tr); });
      const oi = $('[data-f="o"]', tr); oi.value = rc.o ?? '';
      oi.addEventListener('input', () => { const v = oi.value.trim(); oi.value = v === 'ม' ? 'มส' : v; rec(sid).o = ['ร', 'มส'].includes(oi.value) ? oi.value : ''; if (!['ร', 'มส', ''].includes(oi.value) && oi.value !== 'ม') oi.value = ''; touch(); refresh(tr); });
    }
    refresh(tr);
  });
  gridNav(root);

  const reload = () => scoresView(root);
  $('#sj').onchange = e => { U.subj = e.target.value; reload(); };
  $('#term').onchange = e => { U.term = +e.target.value; reload(); };
  $$('[data-m]', root).forEach(b => b.onclick = () => { U.smode = b.dataset.m; reload(); });
  $('#cadd').onclick = () => { if (isS) cfg.max.push(10); else cfg.inds.push(''); if (cols.length >= 50) return toast('สูงสุด 50 คอลัมน์'); touch(); reload(); };
  $('#cdel').onclick = () => { if (cols.length <= 1) return; if (confirm('ลบคอลัมน์สุดท้าย? ข้อมูลในคอลัมน์นั้นจะถูกตัดออกจากการคำนวณ')) { cols.pop(); for (const r of Object.values(R.scores[subj.id] ?? {})) { const x = r[U.term]; if (x) { isS ? (x.s.length = Math.min(x.s.length, cols.length)) : (x.p.length = Math.min(x.p.length, cols.length)); } } touch(); reload(); } };
  $('#copycfg').onclick = () => { const o = U.term === 1 ? 2 : 1; subj.cfg[o] = JSON.parse(JSON.stringify(cfg)); touch(); toast(`คัดลอกโครงสร้างไปภาคเรียนที่ ${o} แล้ว`); };
  if (isS) {
    $('#rb').oninput = e => { cfg.rb = +e.target.value || 0; touch(); $$('tbody tr', root).forEach(refresh); };
    $('#rf').oninput = e => { cfg.rf = +e.target.value || 0; touch(); $$('tbody tr', root).forEach(refresh); };
    $$('[data-max]', root).forEach(i => i.oninput = () => { cfg.max[+i.dataset.max] = +i.value || 0; touch(); $$('tbody tr', root).forEach(refresh); });
  } else $$('[data-iname]', root).forEach(i => i.oninput = () => { cfg.inds[+i.dataset.iname] = i.value; touch(); });
  // หัวตารางจัดตำแหน่งแถว 2 ให้ sticky ใต้แถวแรก
  const h1 = $('thead tr', root); if (h1) $$('thead tr:nth-child(2) th', root).forEach(th => (th.style.top = h1.offsetHeight + 'px'));
}

/* ---------- คุณลักษณะ / ค่านิยม ---------- */
function evalView(root) {
  const R = S.room, U = S.ui;
  U.term = U.term === 'รวม' ? 1 : U.term ?? 1; U.emode ??= 't';
  const isT = U.emode === 't';
  const A = isT ? { k: 't', n: 8, lab: C.TRAITS_SHORT, h: 'คุณลักษณะอันพึงประสงค์ 8 ประการ' } : { k: 'v', n: 12, lab: [...Array(12)].map((_, i) => i + 1), h: 'ค่านิยมหลัก 12 ประการ' };
  const B = isT ? { k: 'r', n: 5, lab: C.RTW, h: 'การอ่าน คิดวิเคราะห์ และเขียน', lv: C.LV_TRAIT } : { k: 'c', n: 5, lab: C.COMPS.map(x => x.replace('ความสามารถใน', '')), h: 'สมรรถนะสำคัญของผู้เรียน', lv: C.LV_COMP };
  A.lv = C.LV_TRAIT;
  const st = sorted();
  const rec = sid => (R.eval[U.term][sid] ??= {});
  root.innerHTML = `<div class="printbar">${termSel(U.term)}<div class="row"><button class="btn ${isT ? 'primary' : ''}" data-m="t">คุณลักษณะฯ + อ่าน คิด เขียน</button><button class="btn ${!isT ? 'primary' : ''}" data-m="v">ค่านิยม 12 + สมรรถนะ</button></div><button class="btn" id="fill">ตั้งค่าเดียวกันทั้งห้อง…</button></div>
  <div class="gridwrap"><table class="g"><thead><tr><th rowspan="2">เลขที่</th><th rowspan="2" class="stk">ชื่อ-สกุล</th><th colspan="${A.n}">${A.h}</th><th rowspan="2">ผล</th><th colspan="${B.n}">${B.h}</th><th rowspan="2">ผล</th></tr>
  <tr>${A.lab.map(l => `<th><div class="${isT ? 'vert' : ''}">${l}</div></th>`).join('')}${B.lab.map(l => `<th><div class="vert">${esc(l)}</div></th>`).join('')}</tr></thead><tbody>
  ${st.map((s, i) => `<tr data-sid="${s.id}"><td class="no">${s.no ?? i + 1}</td><td class="stk">${esc(C.studentName(s))}</td>${Array.from({ length: A.n }, (_, k) => `<td><input type="text" data-g="${A.k}" data-k="${k}" maxlength="1" style="width:2.4rem"></td>`).join('')}<td class="out" data-o="${A.k}"></td>${Array.from({ length: B.n }, (_, k) => `<td><input type="text" data-g="${B.k}" data-k="${k}" maxlength="1" style="width:2.4rem"></td>`).join('')}<td class="out" data-o="${B.k}"></td></tr>`).join('')}</tbody></table></div>
  <p class="hint">กรอก 3 = ดีเยี่ยม · 2 = ดี · 1 = ผ่าน · 0 = ไม่ผ่าน${isT ? '' : ' (สมรรถนะ: 1 = พอใช้, 0 = ปรับปรุง)'} · ผลสรุปคิดจากฐานนิยมของคะแนนรายข้อ (เท่ากันเลือกค่าสูง) · พิมพ์เฉพาะตัวเลข 0-3</p>`;
  const refresh = tr => { const sid = tr.dataset.sid, r = R.eval[U.term][sid] ?? {}; for (const X of [A, B]) { const m = C.modeLevel(r[X.k] ?? []); $(`[data-o="${X.k}"]`, tr).textContent = m === '' ? '' : X.lv[m]; } };
  $$('tbody tr', root).forEach(tr => {
    const sid = tr.dataset.sid, r = R.eval[U.term][sid] ?? {};
    $$('input', tr).forEach(inp => {
      inp.value = r[inp.dataset.g]?.[+inp.dataset.k] ?? '';
      inp.addEventListener('input', () => {
        const v = inp.value.replace(/[^0-3]/g, ''); inp.value = v;
        const a = (rec(sid)[inp.dataset.g] ??= []); a[+inp.dataset.k] = v === '' ? '' : +v; touch(); refresh(tr);
      });
      inp.addEventListener('focus', () => inp.select());
    });
    refresh(tr);
  });
  gridNav(root);
  $('#term').onchange = e => { U.term = +e.target.value; evalView(root); };
  $$('[data-m]', root).forEach(b => b.onclick = () => { U.emode = b.dataset.m; evalView(root); });
  $('#fill').onclick = async () => {
    const f = await ask('ตั้งค่าเดียวกันทั้งห้อง', `<p class="hint">ใส่ค่าให้ทุกคนทุกข้อในชุดนี้ แล้วค่อยแก้เป็นรายคน (ค่าเดิมจะถูกเขียนทับ)</p><label class="f">คะแนน (0-3)<select name="v"><option>3</option><option>2</option><option>1</option><option>0</option></select></label>`, { ok: 'ตั้งค่า' });
    if (!f) return;
    const v = +new FormData(f.querySelector('form')).get('v');
    for (const s of st) for (const X of [A, B]) rec(s.id)[X.k] = Array(X.n).fill(v);
    touch(); evalView(root);
  };
}

/* ---------- กิจกรรม ---------- */
function actsView(root) {
  const R = S.room, U = S.ui; U.term = U.term === 'รวม' ? 1 : U.term ?? 1;
  const st = sorted(), rec = sid => (R.acts[U.term][sid] ??= {});
  const sel = k => `<select data-a="${k}" style="border:0"><option value=""></option><option value="ผ">ผ่าน</option><option value="มผ">ไม่ผ่าน</option></select>`;
  root.innerHTML = `<div class="printbar">${termSel(U.term)}<button class="btn" id="allp">ตั้ง “ผ่าน” ทั้งห้อง (ทุกกิจกรรม)</button></div>
  <div class="gridwrap"><table class="g"><thead><tr><th>เลขที่</th><th class="stk">ชื่อ-สกุล</th><th>แนะแนว</th><th>ลูกเสือ-เนตรนารี</th><th>ชื่อชุมนุม</th><th>ชุมนุม</th><th>กิจกรรมเพื่อสังคมฯ</th></tr></thead><tbody>
  ${st.map((s, i) => `<tr data-sid="${s.id}"><td class="no">${s.no ?? i + 1}</td><td class="stk">${esc(C.studentName(s))}</td><td>${sel('guid')}</td><td>${sel('scout')}</td><td><input type="text" class="w" data-a="clubName"></td><td>${sel('club')}</td><td>${sel('social')}</td></tr>`).join('')}</tbody></table></div>`;
  $$('tbody tr', root).forEach(tr => {
    const r = R.acts[U.term][tr.dataset.sid] ?? {};
    $$('[data-a]', tr).forEach(el => { el.value = r[el.dataset.a] ?? ''; el.addEventListener('input', () => { rec(tr.dataset.sid)[el.dataset.a] = el.value; touch(); }); });
  });
  $('#term').onchange = e => { U.term = +e.target.value; actsView(root); };
  $('#allp').onclick = () => { for (const s of st) Object.assign(rec(s.id), { guid: 'ผ', scout: 'ผ', club: 'ผ', social: 'ผ' }); touch(); actsView(root); };
}

/* ---------- เวลาเรียน ---------- */
function attendView(root) {
  const R = S.room, st = sorted();
  root.innerHTML = `<div class="gridwrap"><table class="g"><thead><tr><th>เลขที่</th><th class="stk">ชื่อ-สกุล</th>${C.MONTHS.map(m => `<th>${m}</th>`).join('')}<th>รวม</th><th>ร้อยละ</th></tr>
  <tr><th></th><th class="stk" style="text-align:right">วันเรียนเต็ม →</th>${C.MONTHS.map(m => `<th><input type="number" data-d="${m}" min="0" max="31" style="width:3.2rem;border:1px solid var(--line);border-radius:6px;text-align:center"></th>`).join('')}<th id="fulltot"></th><th></th></tr></thead><tbody>
  ${st.map((s, i) => `<tr data-sid="${s.id}"><td class="no">${s.no ?? i + 1}</td><td class="stk">${esc(C.studentName(s))}</td>${C.MONTHS.map(m => `<td><input type="text" inputmode="numeric" data-m="${m}"></td>`).join('')}<td class="out" data-o="sum"></td><td class="out" data-o="pct"></td></tr>`).join('')}</tbody></table></div>
  <p class="hint">กรอกจำนวน “วันที่มาเรียน” ของแต่ละคนในแต่ละเดือน และ “วันเรียนเต็ม” ของเดือนนั้นที่แถวบนสุด (ภาคเรียนที่ 1 ปกติเดือน พ.ค.–ก.ย./ต.ค. ภาคเรียนที่ 2 ต.ค./พ.ย.–มี.ค.)</p>`;
  const full = () => C.MONTHS.reduce((a, m) => a + (Number(R.att.days[m]) || 0), 0);
  const refresh = tr => {
    const p = R.att.present[tr.dataset.sid] ?? {}, sum = C.MONTHS.reduce((a, m) => a + (Number(p[m]) || 0), 0), f = full();
    $('[data-o=sum]', tr).textContent = sum || ''; $('[data-o=pct]', tr).textContent = f && sum ? C.fmt(sum * 100 / f, 1) : '';
  };
  const all = () => { $('#fulltot').textContent = full() || ''; $$('tbody tr', root).forEach(refresh); };
  $$('[data-d]', root).forEach(i => { i.value = R.att.days[i.dataset.d] ?? ''; i.oninput = () => { R.att.days[i.dataset.d] = i.value; touch(); all(); }; });
  $$('tbody tr', root).forEach(tr => $$('input', tr).forEach(i => { i.value = R.att.present[tr.dataset.sid]?.[i.dataset.m] ?? ''; i.oninput = () => { i.value = i.value.replace(/\D/g, ''); ((R.att.present[tr.dataset.sid] ??= {}))[i.dataset.m] = i.value; touch(); refresh(tr); }; }));
  gridNav(root); all();
}

/* ---------- น้ำหนัก-ส่วนสูง ---------- */
function healthView(root) {
  const R = S.room, st = sorted();
  root.innerHTML = `<div class="gridwrap"><table class="g"><thead><tr><th rowspan="2">เลขที่</th><th rowspan="2" class="stk">ชื่อ-สกุล</th><th colspan="2">ภาคเรียนที่ 1 ครั้งที่ 1</th><th colspan="2">ภาคเรียนที่ 1 ครั้งที่ 2</th><th colspan="2">ภาคเรียนที่ 2 ครั้งที่ 1</th><th colspan="2">ภาคเรียนที่ 2 ครั้งที่ 2</th></tr>
  <tr>${[0, 1, 2, 3].map(k => `<th colspan="2"><input type="date" data-date="${k}"></th>`).join('')}</tr><tr><th></th><th class="stk"></th>${'<th>น้ำหนัก (กก.)</th><th>ส่วนสูง (ซม.)</th>'.repeat(4)}</tr></thead><tbody>
  ${st.map((s, i) => `<tr data-sid="${s.id}"><td class="no">${s.no ?? i + 1}</td><td class="stk">${esc(C.studentName(s))}</td>${[0, 1, 2, 3].map(k => `<td><input type="text" inputmode="decimal" data-w="${k}"></td><td><input type="text" inputmode="decimal" data-h="${k}"></td>`).join('')}</tr>`).join('')}</tbody></table></div>
  <p class="hint">การเทียบเกณฑ์น้ำหนัก-ส่วนสูงตามอายุ (ภาวะโภชนาการ) ยังไม่ได้ทำในเวอร์ชันนี้ ช่องผลเทียบเกณฑ์ในหน้าพิมพ์ ปพ.6 จะว่างไว้</p>`;
  $$('[data-date]', root).forEach(i => { i.value = R.health.dates[+i.dataset.date] ?? ''; i.onchange = () => { R.health.dates[+i.dataset.date] = i.value; touch(); }; });
  $$('tbody tr', root).forEach(tr => $$('input', tr).forEach(i => {
    const key = i.dataset.w !== undefined ? 'w' : 'h', k = +(i.dataset.w ?? i.dataset.h);
    i.value = R.health.of[tr.dataset.sid]?.[key]?.[k] ?? '';
    i.oninput = () => { const v = i.value.replace(/[^\d.]/g, ''); i.value = v; const o = (R.health.of[tr.dataset.sid] ??= { w: [], h: [] }); o[key] ??= []; o[key][k] = v; touch(); };
  }));
  gridNav(root);
}

/* ---------- ความเห็นครู ---------- */
function remarksView(root) {
  const R = S.room, U = S.ui; U.term = U.term === 'รวม' ? 1 : U.term ?? 1;
  const st = sorted(); U.rs = st.find(s => s.id === U.rs)?.id ?? st[0]?.id;
  if (!st.length) { root.innerHTML = '<p class="muted">เพิ่มนักเรียนก่อน</p>'; return; }
  const i = st.findIndex(s => s.id === U.rs), r = R.remarks[U.term][U.rs] ?? {};
  const T = [['a', 'ด้านหน้าที่รับผิดชอบ ความเอาใจใส่การเรียน'], ['b', 'ด้านการใช้เวลาว่าง'], ['c', 'ด้านความสัมพันธ์กับบุคคลรอบข้าง'], ['d', 'ด้านอุปนิสัย บุคลิกภาพ'], ['e', 'ด้านสุขภาพ']];
  root.innerHTML = `<div class="printbar">${termSel(U.term)}<label class="f">นักเรียน<select id="rs">${st.map(s => `<option value="${s.id}" ${s.id === U.rs ? 'selected' : ''}>${s.no}. ${esc(C.studentName(s))}</option>`).join('')}</select></label><button class="btn" id="pv">← ก่อนหน้า</button><button class="btn" id="nx">ถัดไป →</button></div>
  <div class="card form-grid" style="grid-template-columns:1fr">${T.map(([k, l]) => `<label class="f">${l}<textarea rows="2" data-r="${k}">${esc(r[k] ?? '')}</textarea></label>`).join('')}</div>`;
  $$('[data-r]', root).forEach(t => t.oninput = () => { ((R.remarks[U.term][U.rs] ??= {}))[t.dataset.r] = t.value; touch(); });
  $('#term').onchange = e => { U.term = +e.target.value; remarksView(root); };
  $('#rs').onchange = e => { U.rs = e.target.value; remarksView(root); };
  $('#pv').onclick = () => { if (i > 0) { U.rs = st[i - 1].id; remarksView(root); } };
  $('#nx').onclick = () => { if (i < st.length - 1) { U.rs = st[i + 1].id; remarksView(root); } };
}

/* ---------- พิมพ์ ---------- */
function printView(root) {
  const R = S.room, U = S.ui;
  U.doc ??= 'p5cover'; U.pterm ??= 1; U.pstud ??= 'all';
  const d = DOCS.find(x => x.id === U.doc);
  U.psubj = R.subjects.find(s => s.id === U.psubj)?.id ?? R.subjects[0]?.id;
  const terms = d.needTerm === 'one' ? [1, 2] : [1, 2, 'รวม'];
  if (!terms.map(String).includes(String(U.pterm))) U.pterm = 1;
  root.innerHTML = `<div class="printbar print-ctl"><label class="f">เอกสาร<select id="doc">${DOCS.map(x => `<option value="${x.id}" ${x.id === U.doc ? 'selected' : ''}>${x.label}</option>`).join('')}</select></label>
    ${d.needTerm ? termSel(U.pterm, terms) : ''}
    ${d.needSubj ? `<label class="f">รายวิชา<select id="psubj">${R.subjects.map(s => `<option value="${s.id}" ${s.id === U.psubj ? 'selected' : ''}>${esc(s.code)} ${esc(s.name)}</option>`).join('')}</select></label>` : ''}
    ${d.needStudent ? `<label class="f">นักเรียน<select id="pstud"><option value="all">ทั้งห้อง</option>${sorted().map(s => `<option value="${s.id}" ${s.id === U.pstud ? 'selected' : ''}>${s.no}. ${esc(C.studentName(s))}</option>`).join('')}</select></label>` : ''}
    <button class="btn primary" id="doprint">พิมพ์ / บันทึก PDF</button></div>
  <p class="hint">เลือก “ขนาดกระดาษ A4” และปิดตัวเลือก “ส่วนหัวและท้ายกระดาษ” ในหน้าต่างพิมพ์ของเบราว์เซอร์ เพื่อให้ตรงกับแบบพิมพ์</p>
  <div class="stage"><div class="paper">${buildDoc({ doc: U.doc, term: U.pterm, subjId: U.psubj, studentId: U.pstud }, R, S.school) || '<p style="color:#fff">ไม่มีข้อมูลสำหรับพิมพ์</p>'}</div></div>`;
  $('#doc').onchange = e => { U.doc = e.target.value; printView(root); };
  $('#term')?.addEventListener('change', e => { U.pterm = e.target.value === 'รวม' ? 'รวม' : +e.target.value; printView(root); });
  $('#psubj')?.addEventListener('change', e => { U.psubj = e.target.value; printView(root); });
  $('#pstud')?.addEventListener('change', e => { U.pstud = e.target.value; printView(root); });
  $('#doprint').onclick = () => {
    app.classList.add('print-root'); document.body.classList.add('printing');
    // แสดงเฉพาะเวทีพิมพ์
    const stage = $('.stage', root).cloneNode(true); stage.classList.add('print-root');
    const holder = document.createElement('div'); holder.id = 'printHolder'; holder.appendChild(stage); document.body.appendChild(holder);
    const clean = () => { holder.remove(); app.classList.remove('print-root'); document.body.classList.remove('printing'); window.removeEventListener('afterprint', clean); };
    window.addEventListener('afterprint', clean); window.print();
  };
}

/* ---------- ข้อมูลห้อง ---------- */
function infoView(root) {
  const R = S.room;
  root.innerHTML = `<div class="card"><div class="form-grid">
    <label class="f">ระดับชั้น<select data-i="level">${C.LEVELS.map(l => `<option ${l === R.level ? 'selected' : ''}>${l}</option>`).join('')}</select></label>
    <label class="f">ห้อง<input type="text" data-i="room" value="${esc(R.room)}"></label>
    <label class="f">ปีการศึกษา (พ.ศ.)<input type="number" data-i="year" value="${esc(R.year)}"></label>
    <label class="f span2">ครูประจำชั้น/ครูที่ปรึกษา<input type="text" data-i="advisor" value="${esc(R.advisor)}"></label>
    <label class="f">วันที่ลงนามในเอกสาร<input type="date" data-i="signDate" value="${esc(R.signDate)}"></label></div>
    <div class="row" style="margin-top:1rem"><span class="grow"></span><button class="btn danger" id="delroom">ลบห้องเรียนนี้</button></div></div>`;
  $$('[data-i]', root).forEach(i => i.oninput = () => {
    R[i.dataset.i] = i.type === 'number' ? +i.value : i.value; R.title = `${R.level}/${R.room}`; touch();
    $('h2').textContent = R.title;
  });
  $('#delroom').onclick = async () => {
    if (!confirm(`ลบห้อง ${R.title} ทั้งห้อง รวมนักเรียนและคะแนนทั้งหมด? ย้อนกลับไม่ได้`)) return;
    await store.deleteRoom(R.id); S.room = null; location.hash = '#/';
  };
}

boot();
