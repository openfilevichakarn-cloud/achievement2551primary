/* app.js — แอปหลัก ปพ.5/ปพ.6 หลักสูตร 2551: เข้าสู่ระบบด้วยรหัสโรงเรียน, โหลด/บันทึก Supabase,
 * แสดงชีทที่แปลงจาก Excel (data/template.json) · กรอกข้อมูลในตาราง · คำนวณสูตรด้วย engine.js · พิมพ์ */
(function () {
  'use strict';
  const P = window.PPEngine, R = window.PPRender;
  const $ = (s, el) => (el || document).querySelector(s);
  const $$ = (s, el) => Array.from((el || document).querySelectorAll(s));
  const esc = R.esc;
  const CFG = window.PP_CONFIG || {};
  const backend = window.PPBackend.create(CFG);
  const LS = {
    get(k) { try { return localStorage.getItem(k); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, v); } catch (e) { /* ignore */ } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { /* ignore */ } },
  };

  const INFO = 'ข้อมูลพื้นฐาน', SUBJ = 'รายวิชา', ROSTER = 'ข้อมูลเพิ่มเติม';
  const S = {
    code: null, school: null, grade: 1, room: 1, rooms: {}, template: null, wb: null, styles: null,
    view: { rowHidden: {}, colHidden: null, bg: {}, locked: new Set() },
    userEdits: {}, att: {}, version: 0, dirty: false, saving: false, active: null, zoom: 'fit', saveTimer: null, offline: false,
  };

  // ---------------------------------------------------------------- UI helpers
  function toast(msg, kind) {
    const el = document.createElement('div');
    el.className = 'toast ' + (kind || '');
    el.textContent = msg;
    $('#toasts').appendChild(el);
    setTimeout(() => el.classList.add('out'), 4200);
    setTimeout(() => el.remove(), 4800);
  }
  function modal(title, body, buttons) {
    return new Promise((resolve) => {
      const root = $('#modal');
      root.innerHTML = `<div class="modal-card" role="dialog" aria-modal="true"><h3>${esc(title)}</h3><div class="modal-body">${body}</div><div class="modal-actions"></div></div>`;
      const act = $('.modal-actions', root);
      (buttons || [{ label: 'ตกลง', value: true, primary: true }]).forEach((b) => {
        const btn = document.createElement('button');
        btn.textContent = b.label; btn.className = 'btn' + (b.primary ? ' primary' : '');
        btn.onclick = () => { root.hidden = true; root.innerHTML = ''; resolve(b.value); };
        act.appendChild(btn);
      });
      root.hidden = false;
      const first = $('.btn.primary', act) || $('.btn', act); if (first) first.focus();
    });
  }
  function setStatus(txt, cls) { const el = $('#savestat'); el.textContent = txt; el.className = 'savestat ' + (cls || ''); }
  async function doLogin(code, remember) {
    $('#login-err').textContent = '';
    const btn = $('#login-go'); btn.disabled = true; btn.textContent = 'กำลังตรวจสอบ...';
    try {
      const res = await backend.login(code);
      if (!res || !res.ok) throw new Error((res && res.error) || 'รหัสสิทธิ์ไม่ถูกต้อง');
      const ready = (window.PP_CONFIG && PP_CONFIG.GRADES_READY) || [1];
      const granted = res.grades && res.grades.length ? res.grades : [1];
      const grades = granted.filter(g => ready.includes(g));
      if (!grades.length) throw new Error('ชั้นที่โรงเรียนได้รับสิทธิ์ (ป.' + granted.join(', ป.') + ') ยังไม่เปิดให้ใช้งาน กรุณาติดต่อผู้ดูแล');
      S.code = code; S.school = { name: res.school_name, grades, logo: res.logo || null };
      // ห้องเรียนที่มีข้อมูลแล้ว (ชั้นที่ยังไม่มี = ห้อง 1)
      S.rooms = {};
      for (const x of res.rooms || []) if (grades.includes(x.grade)) (S.rooms[x.grade] = S.rooms[x.grade] || []).push(x.room);
      for (const g of grades) S.rooms[g] = [...new Set(S.rooms[g] && S.rooms[g].length ? S.rooms[g] : [1])].sort((a, b) => a - b);
      if (remember) LS.set('pp_code', code); else LS.del('pp_code');
      const last = +LS.get('pp_grade'), lastRoom = +LS.get('pp_room') || 1;
      const g0 = S.school.grades.includes(last) ? last : S.school.grades[0];
      await openGrade(g0, S.rooms[g0].includes(lastRoom) ? lastRoom : S.rooms[g0][0]);
    } catch (e) {
      $('#login-err').textContent = e.offline ? 'เชื่อมต่ออินเทอร์เน็ตไม่ได้ กรุณาลองใหม่' : (e.message || 'เข้าสู่ระบบไม่สำเร็จ');
      LS.del('pp_code');
    } finally { btn.disabled = false; btn.textContent = 'เข้าสู่ระบบ'; }
  }

  function logout() {
    if (S.dirty && !confirm('ยังมีข้อมูลที่ยังไม่ได้บันทึก ต้องการออกจากระบบหรือไม่')) return;
    LS.del('pp_code'); location.reload();
  }

  // ---------------------------------------------------------------- open workbook

  // ---------------------------------------------------------------- เปิดห้อง (โหลดเทมเพลต + ข้อมูลที่บันทึกไว้)
  const roomLabel = (g, r) => 'ป.' + g + '/' + r;
  function draftKey() { return 'pp_draft_' + P_hash(S.code) + '_' + S.grade + (S.room > 1 ? '_' + S.room : ''); }
  function P_hash(s) { let h = 5381; for (const ch of String(s)) h = ((h << 5) + h + ch.charCodeAt(0)) | 0; return (h >>> 0).toString(36); }

  async function loadTemplate() {
    if (S.template) return S.template;
    const resp = await fetch((CFG.TEMPLATE_URL || 'data/template.json') + '?v=' + (CFG.BUILD || ''));
    if (!resp.ok) throw new Error('ไม่พบข้อมูลโปรแกรม (data/template.json)');
    return (S.template = await resp.json());
  }

  async function openGrade(g, room) {
    room = room || (S.rooms[g] && S.rooms[g][0]) || 1;
    S.grade = g; S.room = room; LS.set('pp_grade', String(g)); LS.set('pp_room', String(room));
    if (!S.rooms[g].includes(room)) S.rooms[g] = [...S.rooms[g], room].sort((a, b) => a - b);
    $('#login').hidden = true; $('#app').hidden = false; $('#loading').hidden = false;
    $('#loading-msg').textContent = 'กำลังโหลดโปรแกรม ' + roomLabel(g, room) + ' ...';
    try { await loadTemplate(); } catch (e) { $('#loading-msg').textContent = e.message; return; }
    await new Promise(r => setTimeout(r, 30));   // ให้หน้าจอวาดข้อความ "กำลังโหลด" ก่อนสร้างเวิร์กบุ๊ก
    S.styles = S.template.styles;
    $('#xl-css').textContent = R.styleCss(S.styles);
    S.view = { rowHidden: {}, colHidden: null, bg: {}, locked: new Set(), logo: S.school.logo || null };
    S.userEdits = {}; S.att = {};
    $('#loading-msg').textContent = 'กำลังโหลดข้อมูลของโรงเรียน...';
    let doc = null;
    try {
      const res = await backend.load(S.code, g, room);
      S.version = (res && res.version) || 0; doc = res && res.data;
      S.offline = false;
    } catch (e) {
      S.offline = true;
      const draft = LS.get(draftKey());
      if (draft) { try { doc = JSON.parse(draft).data; S.version = JSON.parse(draft).version || 0; } catch (x) { /* ignore */ } }
      toast('โหลดจากเซิร์ฟเวอร์ไม่ได้ ใช้ข้อมูลที่เก็บในเครื่องนี้แทน', 'warn');
    }
    S.items = PPAssess.normalize(S.template, doc && doc.items);
    if (doc && doc.cells && !doc.items) for (const k of Object.keys(PPAssess.KINDS)) PPAssess.remapEdits(doc.cells, k, PPAssess.KINDS[k].base, S.items[k].n);   // เอกสารเก่า (ก่อนมีการปรับหัวข้อ)
    makeWb();
    applyDoc(doc);
    $('#school-name').textContent = S.school.name;
    $('#grade-label').textContent = roomLabel(g, room);
    buildGradeSwitch();
    $('#loading').hidden = true;
    window.PPUI.mount();
    setStatus(S.offline ? 'ออฟไลน์ (ใช้ข้อมูลในเครื่อง)' : 'พร้อมใช้งาน', S.offline ? 'warn' : 'ok');
  }

  // เวิร์กบุ๊กสร้างจากแม่แบบที่ปรับจำนวนหัวข้อ คุณลักษณะ/ค่านิยม ตามค่าตั้งของห้องนี้ (S.items)
  function makeWb() { S.wb = new P.Workbook(PPAssess.apply(S.template, S.items)); }
  function rebuildWb() {
    makeWb();
    for (const [sh, m] of Object.entries(S.userEdits)) if (S.wb.sheet(sh)) for (const [ref, v] of Object.entries(m)) S.wb.setValue(sh, ref, v);
    lockSchoolName(); roomClassLabel();
  }
  // แก้ค่าตั้งหัวข้อของ kind (trait|value): fn แก้สำเนา {n, g, names} · del = {g, i} เมื่อลบหัวข้อ · structure=true เมื่อต้องสร้างเวิร์กบุ๊กใหม่
  function editItems(kind, fn, del) {
    const cur = JSON.parse(JSON.stringify(S.items[kind])), old = cur.n.slice(); fn(cur);
    const nx = PPAssess.normalize(S.template, { [kind]: cur })[kind];
    S.items = { ...S.items, [kind]: nx };
    PPAssess.remapEdits(S.userEdits, kind, old, nx.n, del);
    rebuildWb(); markDirty();
  }
  // เปลี่ยนชื่อหัวข้อโดยไม่ต้องสร้างเวิร์กบุ๊กใหม่
  function renameItem(kind, g, i, name) {
    S.items[kind].names[g][i] = name;
    const col = PPAssess.numCol(PPAssess.layout(kind, S.items[kind].n).items[g][i]) + '3', v = /^\d+$/.test(name) ? +name : name;
    for (const sh of PPAssess.sheetsOf(kind)) S.wb.setValue(sh, col, v);
    markDirty();
  }
  // วันที่ออกเอกสาร (เลขวันที่แบบ Excel ปี พ.ศ.) → ช่องวันที่ในปก ปพ.5 (วัน/เดือน/ปี) และ ปพ.6 (ข้อความ)
  const TH_MONTHS = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
  // term: 1 = เทอม 1 · 2 = เทอม 2 และฉบับรวมทั้งปี (ใช้วันที่ปลายปีการศึกษา)
  function issueTargets(serial, term) {
    const out = [], dt = typeof serial === 'number' ? new Date(Date.UTC(1899, 11, 30) + serial * 864e5) : null;
    const d = dt ? dt.getUTCDate() : null, m = dt ? TH_MONTHS[dt.getUTCMonth()] : null, y = dt ? dt.getUTCFullYear() : null;
    for (const s of term === 1 ? ['1'] : ['2', 'รวม']) {
      out.push(['ปก ปพ.5 รายชั้น ' + s, 'AE57', d], ['ปก ปพ.5 รายชั้น ' + s, 'AP57', m], ['ปก ปพ.5 รายชั้น ' + s, 'BF57', y]);
      out.push(['ปก ปพ.5 รายวิชา ' + s, 'AE41', d], ['ปก ปพ.5 รายวิชา ' + s, 'AP41', m], ['ปก ปพ.5 รายวิชา ' + s, 'BF41', y]);
      out.push(['ปพ.6 ' + (s === 'รวม' ? 'รวม' : 'เทอม ' + s), 'F34', dt ? d + ' ' + m + ' ' + y : null]);
    }
    return out;
  }
  function applyDoc(doc) {
    S.att = (doc && doc.att) || {};
    if (doc && doc.cells) {
      for (const [sheet, m] of Object.entries(doc.cells)) {
        if (!S.wb.sheet(sheet)) continue;
        for (const [ref, val] of Object.entries(m)) { S.wb.setValue(sheet, ref, val); (S.userEdits[sheet] = S.userEdits[sheet] || {})[ref] = val; }
      }
    }
    // วันที่ออกเอกสาร ปพ.6: ถ้ายังไม่เคยกรอก ตั้งเป็นวันนี้ (พ.ศ.) ให้คำนวณอายุได้
    const now = new Date(), today = thDate(String(now.getDate()).padStart(2, '0') + '/' + String(now.getMonth() + 1).padStart(2, '0') + '/' + (now.getFullYear() + 543));
    for (const b of ['เล่ม ปพ.6 เทอม 1', 'เล่ม ปพ.6 เทอม 2', 'เล่ม ปพ.6 รวม']) if (S.wb.sheet(b) && !(S.userEdits[b] && S.userEdits[b].CX20 !== undefined) && today !== null) S.wb.setValue(b, 'CX20', today);
    if (today !== null) for (const [sh, ref, v] of [...issueTargets(today, 1), ...issueTargets(today, 2)]) if (S.wb.sheet(sh) && !(S.userEdits[sh] && S.userEdits[sh][ref] !== undefined)) S.wb.setValue(sh, ref, v);
    lockSchoolName(); roomClassLabel();
  }
  // ชื่อโรงเรียนมาจากเซิร์ฟเวอร์ตามรหัสสิทธิ์ แก้ไม่ได้
  function lockSchoolName() {
    S.wb.setValue(INFO, 'B3', S.school.name);
    S.view.locked.add(INFO + '!B3');
    delete (S.userEdits[INFO] || {}).B3;
  }
  // ชื่อชั้นในเอกสาร (รายวิชา!B2 ใช้ในสูตรของเอกสารทุกฉบับ) ล็อกตามชั้น/ห้องที่เปิดอยู่
  function classText() {
    const multi = S.room > 1 || (S.rooms[S.grade] || []).length > 1;
    return 'ประถมศึกษาปีที่ ' + S.grade + (multi ? '/' + S.room : '');
  }
  function roomClassLabel() {
    S.wb.setValue(SUBJ, 'B2', classText());
    S.view.locked.add(SUBJ + '!B2');
    delete (S.userEdits[SUBJ] || {}).B2;
  }
  function buildGradeSwitch() {
    const sel = $('#grade-switch');
    sel.hidden = false;
    sel.innerHTML = S.school.grades.map(g => `<optgroup label="ชั้น ป.${g}">${S.rooms[g].map(r => `<option value="${g}-${r}"${g === S.grade && r === S.room ? ' selected' : ''}>${roomLabel(g, r)}</option>`).join('')}</optgroup>`).join('')
      + '<option value="manage">⚙ จัดการห้องเรียน…</option>';
    sel.onchange = async () => {
      if (sel.value === 'manage') { sel.value = S.grade + '-' + S.room; manageRooms(); return; }
      const [g, r] = sel.value.split('-').map(Number);
      await flushSave(); openGrade(g, r);
    };
  }
  // จัดการห้องเรียน: เพิ่ม / ลบห้อง (ชั้นละไม่เกิน 10 ห้อง) — ใช้รหัสโรงเรียนเดียวกันทุกห้อง
  const MAX_ROOMS = 10;
  function manageRooms() {
    const root = $('#modal');
    const draw = () => {
      root.innerHTML = `<div class="modal-card rooms-card" role="dialog" aria-modal="true"><h3>จัดการห้องเรียน</h3><div class="modal-body">
        <p class="hint">แต่ละห้องมีข้อมูลแยกกัน (นักเรียน คะแนน ผลประเมิน) · ชั้นละไม่เกิน ${MAX_ROOMS} ห้อง · ทุกห้องใช้รหัสโรงเรียนเดียวกัน</p>
        ${S.school.grades.map(g => `<div class="rm-row"><b>ป.${g}</b><div class="rm-chips">${S.rooms[g].map(r => `<span class="rm-chip${g === S.grade && r === S.room ? ' cur' : ''}"><button class="rm-open" data-g="${g}" data-r="${r}" title="เปิดห้องนี้">${roomLabel(g, r)}</button>${S.rooms[g].length > 1 ? `<button class="rm-del" data-g="${g}" data-r="${r}" aria-label="ลบ ${roomLabel(g, r)}" title="ลบห้องนี้">×</button>` : ''}</span>`).join('')}
          ${S.rooms[g].length < MAX_ROOMS ? `<button class="btn sm rm-add" data-g="${g}">＋ เพิ่มห้อง</button>` : '<span class="hint">ครบ 10 ห้องแล้ว</span>'}</div></div>`).join('')}
        </div><div class="modal-actions"><button class="btn primary" data-x>ปิด</button></div></div>`;
      root.hidden = false;
    };
    const close = () => { root.hidden = true; root.innerHTML = ''; root.onclick = null; };
    root.onclick = async (e) => {
      const b = e.target.closest('button'); if (!b) return;
      if (b.hasAttribute('data-x')) { close(); return; }
      const g = +b.dataset.g, r = +b.dataset.r;
      if (b.classList.contains('rm-open')) { close(); await flushSave(); openGrade(g, r); return; }
      if (b.classList.contains('rm-add')) {
        const n = Array.from({ length: MAX_ROOMS }, (_, i) => i + 1).find(x => !S.rooms[g].includes(x)); if (!n) return;
        b.disabled = true;
        try {
          // ห้องเดิมที่ยังไม่เคยบันทึก (เช่น ห้อง 1 ที่ยังว่าง) บันทึกไว้ก่อน ไม่ให้หายจากรายการเมื่อเข้าระบบครั้งหน้า
          await flushSave();
          for (const r0 of S.rooms[g]) {
            if (g === S.grade && r0 === S.room && S.version > 0) continue;
            const cur = await backend.load(S.code, g, r0);
            if (cur && !cur.version) await backend.save(S.code, g, r0, cur.data || { v: 3, grade: g, room: r0, cells: {}, savedAt: new Date().toISOString() }, 0, false);
          }
          const res = await backend.save(S.code, g, n, { v: 3, grade: g, room: n, cells: {}, savedAt: new Date().toISOString() }, 0, false);
          if (!res || res.ok === false) throw new Error((res && res.error) || 'เพิ่มห้องไม่สำเร็จ');
        } catch (x) { toast(x.message || 'เพิ่มห้องไม่สำเร็จ', 'warn'); b.disabled = false; return; }
        S.rooms[g] = [...S.rooms[g], n].sort((a, c) => a - c);
        close(); await flushSave(); toast('เพิ่มห้อง ' + roomLabel(g, n) + ' แล้ว', 'ok'); openGrade(g, n); return;
      }
      if (b.classList.contains('rm-del')) {
        if (!confirm(`ลบห้อง ${roomLabel(g, r)} พร้อมข้อมูลทั้งหมดของห้องนี้ (นักเรียน คะแนน ผลประเมิน)?\nย้อนกลับไม่ได้ — แนะนำให้กด “สำรองข้อมูล” ที่หน้าแรกก่อน`)) return;
        try { const res = await backend.deleteRoom(S.code, g, r); if (res && res.ok === false) throw new Error(res.error); }
        catch (x) { toast('ลบห้องไม่สำเร็จ: ' + (x.message || ''), 'warn'); return; }
        S.rooms[g] = S.rooms[g].filter(x => x !== r);
        LS.del('pp_draft_' + P_hash(S.code) + '_' + g + (r > 1 ? '_' + r : ''));
        toast('ลบห้อง ' + roomLabel(g, r) + ' แล้ว', 'ok');
        if (g === S.grade && r === S.room) { S.dirty = false; clearTimeout(S.saveTimer); close(); openGrade(g, S.rooms[g][0]); return; }
        if (g === S.grade) { roomClassLabel(); renderSheet(true); }   // เหลือห้องเดียว: ชื่อชั้นตัด /ห้อง ออก
        buildGradeSwitch(); draw();
      }
    };
    draw();
  }

  function buildDoc() { return { v: 3, grade: S.grade, room: S.room, cells: S.userEdits, items: S.items, att: S.att, active: S.active, savedAt: new Date().toISOString() }; }
  function markDirty() {
    if (S.batch) { S.pendingDirty = true; S.dirty = true; return; }
    S.dirty = true; setStatus('มีการแก้ไข... (จะบันทึกอัตโนมัติ)', 'dirty');
    LS.set(draftKey(), JSON.stringify({ data: buildDoc(), version: S.version }));
    clearTimeout(S.saveTimer); S.saveTimer = setTimeout(save, 1500);
  }
  async function save(force) {
    if (S.saving) { S.saveTimer = setTimeout(save, 800); return; }
    if (!S.dirty && !force) return;
    S.saving = true; setStatus('กำลังบันทึก...', 'saving');
    const snapshot = buildDoc();
    try {
      const res = await backend.save(S.code, S.grade, S.room, snapshot, S.version, force === true);
      if (res && res.conflict) {
        S.saving = false;
        const choice = await modal('ข้อมูลถูกแก้ไขจากเครื่องอื่น', '<p>มีการบันทึกข้อมูลใหม่กว่าจากอีกเครื่องหรืออีกแท็บ</p><p>ต้องการ <b>โหลดข้อมูลล่าสุด</b> (ข้อมูลที่แก้ในหน้านี้จะหายไป) หรือ <b>บันทึกทับ</b> ด้วยข้อมูลในหน้านี้?</p>',
          [{ label: 'โหลดข้อมูลล่าสุด', value: 'reload' }, { label: 'บันทึกทับ', value: 'force', primary: true }]);
        if (choice === 'reload') { S.dirty = false; LS.del(draftKey()); location.reload(); return; }
        return save(true);
      }
      S.version = res.version; S.dirty = false; S.offline = false; LS.set(draftKey(), JSON.stringify({ data: snapshot, version: S.version }));
      setStatus('บันทึกแล้ว ' + new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' }), 'ok');
    } catch (e) {
      S.offline = !!e.offline;
      setStatus(e.offline ? 'ออฟไลน์ — เก็บไว้ในเครื่อง จะบันทึกขึ้นระบบเมื่อมีเน็ต' : 'บันทึกไม่สำเร็จ: ' + e.message, 'err');
      clearTimeout(S.saveTimer); S.saveTimer = setTimeout(save, 15000);
    } finally { S.saving = false; }
  }
  async function flushSave() { clearTimeout(S.saveTimer); if (S.dirty) await save(); }
  const BACKUP_TAG = 'pp56-backup';
  // ไฟล์รุ่น 2: rooms = { "ชั้น/ห้อง": ข้อมูล } + ตราโรงเรียน  (รุ่น 1: grades = { ชั้น: ข้อมูล } = ห้อง 1)
  async function exportBackup() {
    await flushSave();
    const rooms = {};
    for (const g of S.school.grades) for (const r of S.rooms[g]) {
      if (g === S.grade && r === S.room) { rooms[g + '/' + r] = buildDoc(); continue; }
      try { const res = await backend.load(S.code, g, r); if (res && res.data) rooms[g + '/' + r] = res.data; }
      catch (e) { toast('โหลดข้อมูล ' + roomLabel(g, r) + ' ไม่ได้ — ' + (e.message || ''), 'warn'); return; }
    }
    const file = { app: BACKUP_TAG, ver: 2, school: S.school.name, exportedAt: new Date().toISOString(), logo: S.school.logo || null, rooms };
    const d = new Date(), pad = (x) => String(x).padStart(2, '0');
    const name = `backup-pp5-pp6_${d.getFullYear() + 543}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}.json`;   // ชื่อไฟล์ภาษาอังกฤษ: บางเบราว์เซอร์ไม่รองรับชื่อภาษาไทย
    const a = document.createElement('a');
    a.href = URL.createObjectURL(new Blob([JSON.stringify(file)], { type: 'application/json' }));
    a.download = name; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    toast('ดาวน์โหลดไฟล์สำรองแล้ว (' + Object.keys(rooms).length + ' ห้อง)', 'ok');
  }
  function pickBackup() {
    const inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.json,application/json';
    inp.onchange = () => { if (inp.files[0]) importBackup(inp.files[0]); };
    inp.click();
  }
  async function importBackup(f) {
    let file;
    try { file = JSON.parse(await f.text()); } catch (e) { file = null; }
    const src = file && file.app === BACKUP_TAG ? (file.rooms || (file.grades && Object.fromEntries(Object.entries(file.grades).map(([g, d]) => [g + '/1', d])))) : null;
    if (!src || typeof src !== 'object') { toast('ไฟล์นี้ไม่ใช่ไฟล์สำรองข้อมูลของโปรแกรม', 'warn'); return; }
    const all = Object.entries(src).map(([k, doc]) => { const [g, r] = k.split('/').map(Number); return { g, r, doc }; }).filter(x => x.doc && x.doc.cells);
    const list = all.filter(x => S.school.grades.includes(x.g) && x.r >= 1 && x.r <= MAX_ROOMS).sort((a, b) => a.g - b.g || a.r - b.r);
    const skip = all.filter(x => !list.includes(x));
    if (!list.length) { toast('ไฟล์สำรองไม่มีข้อมูลของชั้นที่โรงเรียนนี้ใช้งาน', 'warn'); return; }
    const when = file.exportedAt ? new Date(file.exportedAt).toLocaleString('th-TH', { dateStyle: 'medium', timeStyle: 'short' }) : '-';
    const other = file.school && file.school !== S.school.name;
    const yes = await modal('นำเข้าข้อมูลสำรอง', `<p>ไฟล์สำรองของ <b>${esc(file.school || '-')}</b> เมื่อ ${esc(when)}</p>
      ${other ? '<p class="err">ชื่อโรงเรียนในไฟล์ไม่ตรงกับโรงเรียนที่เข้าสู่ระบบอยู่</p>' : ''}
      <p>ข้อมูลปัจจุบันของ <b>${list.map(x => roomLabel(x.g, x.r)).join(', ')}</b> จะถูก<b>แทนที่ทั้งหมด</b>ด้วยข้อมูลในไฟล์ (ย้อนกลับไม่ได้)</p>
      ${skip.length ? `<p class="hint">ข้าม ${skip.map(x => roomLabel(x.g, x.r)).join(', ')} (โรงเรียนนี้ไม่มีสิทธิ์ชั้นนี้)</p>` : ''}
      <p class="hint">แนะนำ: กด “สำรองข้อมูล” เก็บข้อมูลปัจจุบันไว้ก่อน</p>`,
      [{ label: 'ยกเลิก', value: false }, { label: 'แทนที่ด้วยไฟล์สำรอง', value: true, primary: true }]);
    if (!yes) return;
    clearTimeout(S.saveTimer); S.dirty = false;
    try {
      for (const { g, r, doc: d } of list) {
        const cur = await backend.load(S.code, g, r);
        const doc = { ...d, grade: g, room: r, savedAt: new Date().toISOString() }; delete doc.logo;
        const res = await backend.save(S.code, g, r, doc, (cur && cur.version) || 0, true);
        if (!res || res.ok === false) throw new Error((res && res.error) || 'บันทึกไม่สำเร็จ');
        if (!S.rooms[g].includes(r)) S.rooms[g] = [...S.rooms[g], r].sort((a, b) => a - b);
        LS.del('pp_draft_' + P_hash(S.code) + '_' + g + (r > 1 ? '_' + r : ''));
      }
      const logo = file.logo || (list.find(x => x.doc.logo) || {}).doc?.logo;
      if (logo && !S.school.logo) { await backend.setLogo(S.code, logo); S.school.logo = logo; }
    } catch (e) { toast('นำเข้าไม่สำเร็จ: ' + (e.message || e), 'warn'); return; }
    toast('นำเข้าข้อมูลสำรองแล้ว (' + list.length + ' ห้อง)', 'ok');
    const keep = list.some(x => x.g === S.grade && x.r === S.room);
    openGrade(keep ? S.grade : list[0].g, keep ? S.room : list[0].r);
  }


  // ---------------------------------------------------------------- เมนูชีท
  const GROUPS = [
    ['เริ่มต้นใช้งาน', /^(สารบัญ|คู่มือ|ข้อมูลพื้นฐาน|รายวิชา|ข้อมูลนักเรียน|ข้อมูลเพิ่มเติม|ปฏิทิน)$/],
    ['บันทึกผลรายวิชา', /^วิชาที่ \d+$/],
    ['คุณลักษณะ · ค่านิยม · กิจกรรม', /^(คุณลักษณะ|ค่านิยม|กิจกรรม) |^(เวลาเรียน|น้ำหนัก-ส่วนสูง|เอกสารแนบท้าย)$/],
    ['สรุปผลรายชั้น', /^(คะแนนรายชั้น|คะแนนร้อยละ|เกรดรายชั้น|เกรดร้อยละ) |^(ชื่อนักเรียน|บิดามารดา|ผู้ปกครอง)$/],
    ['ปก ปพ.5', /^ปก /],
    ['ปพ.6', /ปพ\.6|คำชี้แจง/],
  ];
  const subjName = (n) => { const v = S.wb.get(S.wb.sheet(SUBJ), 5 + n, 3); return v === null || v === '' || P.isErr(v) ? '' : String(v); };
  // รายการชีทแยกกลุ่ม (ใช้ในเมนู "เอกสาร")
  function sheetGroups() {
    const groups = GROUPS.map(g => ({ title: g[0], re: g[1], items: [] })); const other = { title: 'อื่น ๆ', items: [] };
    for (const sh of S.wb.sheets) {
      if (sh.def.hidden || /โภชนาการ|ความชุก|^ใบรับรอง$|^เล่ม ปพ\.6 เทอม/.test(sh.name)) continue;   // ไม่ใช้ในเว็บนี้
      (groups.find(g => g.re.test(sh.name)) || other).items.push(sh.name);
    }
    return [...groups, other].filter(g => g.items.length && g.title !== 'เริ่มต้นใช้งาน');   // ชีทกรอกข้อมูลต้นทางมีหน้ากรอกแบบเว็บอยู่แล้ว (ข้อมูล ›) ไม่แสดงในเมนูเอกสาร
  }

  // ---------------------------------------------------------------- แสดง / แก้ไขชีท
  const extent = {};
  // ขอบเขตที่แสดง: ช่วงที่มีข้อมูล/สูตร/ช่องกรอก (ตัดแถว-คอลัมน์ว่างท้ายชีทที่มีแต่สไตล์)
  const MIN_ROWS = { 'ข้อมูลนักเรียน': 55 };
  function viewArea(name) {
    if (extent[name]) return extent[name];
    const sh = S.wb.sheet(name); let r2 = 1, c2 = 1;
    for (const cell of sh.cells.values()) if (cell.f || cell.u || (cell.v !== null && cell.v !== '')) { if (cell.r > r2) r2 = cell.r; if (cell.c > c2) c2 = cell.c; }
    const pa = sh.def.print && sh.def.print.area ? R.parseRange(sh.def.print.area) : null;
    if (pa) { r2 = Math.max(r2, pa.r2); c2 = Math.max(c2, pa.c2); }
    r2 = Math.max(r2, Math.min(sh.def.dim[0], MIN_ROWS[name] || 0));
    return (extent[name] = 'A1:' + P.numToCol(c2) + r2);
  }
  // ตราโรงเรียนวางมุมซ้ายบนของปก ปพ.5 / ปพ.6 (ไม่ล้ำแถว limitRow)
  function logoAnchor(name) {
    if (/^ปก ปพ\.5 /.test(name)) return { r: 2, c: 2, limitRow: 5 };
    if (/^เล่ม ปพ\.6/.test(name)) return { r: 2, c: 1, limitRow: 7 };
    if (/^ปพ\.6 /.test(name)) return { r: 1, c: 1, limitRow: 2 };
    return null;
  }
  // ตราเก็บที่เดียวต่อโรงเรียน ใช้กับทุกชั้น ทุกห้อง
  function loadLogo(file) {
    if (!file) return;
    const img = new Image(), url = URL.createObjectURL(file);
    img.onload = () => {
      const k = Math.min(1, 300 / Math.max(img.width, img.height));
      const cv = document.createElement('canvas'); cv.width = Math.round(img.width * k); cv.height = Math.round(img.height * k);
      cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
      const logo = cv.toDataURL('image/png'); URL.revokeObjectURL(url);
      backend.setLogo(S.code, logo).then((res) => {
        if (res && res.ok === false) throw new Error(res.error);
        S.school.logo = S.view.logo = logo; if (window.PPUI) PPUI.refresh(); toast('ใส่ตราโรงเรียนแล้ว (ใช้กับทุกชั้น ทุกห้อง)', 'ok');
      }).catch((e) => toast('บันทึกตราไม่สำเร็จ: ' + (e.message || ''), 'warn'));
    };
    img.onerror = () => toast('อ่านไฟล์รูปไม่ได้', 'warn');
    img.src = url;
  }
  async function removeLogo() {
    const yes = await modal('ตราโรงเรียน', '<p>ต้องการลบตราโรงเรียนออกจากทุกเอกสารใช่หรือไม่</p>', [{ label: 'ยกเลิก', value: false }, { label: 'ลบตรา', value: true, primary: true }]);
    if (!yes) return;
    try { const res = await backend.setLogo(S.code, null); if (res && res.ok === false) throw new Error(res.error); }
    catch (e) { toast('ลบตราไม่สำเร็จ: ' + (e.message || ''), 'warn'); return; }
    S.school.logo = S.view.logo = null; if (window.PPUI) PPUI.refresh();
  }
  function canEdit(name, r, c, cell) {
    if (cell && cell.f) return false;
    if (name === 'ข้อมูลนักเรียน') return r >= 3 && c <= 90;
    return !!(cell && cell.u);
  }
  function openSheet(name) {
    if (!S.wb.sheet(name)) return;
    S.active = name; $('#sheet-title').textContent = name;
    $('#docviewer').hidden = false; document.body.classList.add('docopen');
    renderSheet(); renderToolbar();
    const sc = $('#sheet-scroll'); sc.scrollTop = 0; sc.scrollLeft = 0;
  }
  function closeDoc() { endEdit(false); $('#docviewer').hidden = true; document.body.classList.remove('docopen'); if (window.PPUI) PPUI.refresh(); }
  function sheetHtmlFor(name, edit, area, viewOv) {
    if (window.PPUI && PPUI.beforeSheet) PPUI.beforeSheet(name);
    return R.sheetHtml({ wb: S.wb, sheetName: name, view: viewOv ? { ...S.view, ...viewOv } : S.view, edit, area, styles: S.styles, canEdit, logo: S.view.logo, logoAt: logoAnchor });
  }
  function renderSheet(keep) {
    const sc = $('#sheet-scroll'), top = sc.scrollTop, left = sc.scrollLeft;
    const out = sheetHtmlFor(S.active, true, viewArea(S.active));
    const host = $('#sheet-host'); host.innerHTML = out.html;
    const avail = Math.max(200, sc.clientWidth - 28);
    const z = S.zoom === 'fit' ? Math.max(0.7, Math.min(1.6, avail / Math.max(1, out.width))) : +S.zoom;   // ชีทกว้างมาก: ไม่ย่อต่ำกว่า 70% (เลื่อนดูแนวนอน)
    const zz = Math.max(0.2, z);
    host.style.zoom = zz; host.style.setProperty('--bw', (1 / zz) + 'px');
    if (keep) { sc.scrollTop = top; sc.scrollLeft = left; }
  }

  // รายชื่อนักเรียนในห้อง (จากชีทข้อมูลเพิ่มเติม: A = เลขที่, C-E = คำนำหน้า ชื่อ สกุล)
  function roster() {
    const sh = S.wb.sheet(ROSTER), out = [];
    for (let r = 2; r <= 60; r++) {
      const n = S.wb.get(sh, r, 1); if (typeof n !== 'number') continue;
      const g = (c) => { const v = S.wb.get(sh, r, c); return v === null || v === 0 || P.isErr(v) ? '' : String(v); };   // ช่องว่างในไฟล์ต้นฉบับอ่านได้เป็น 0
      out.push({ n, name: (g(3) + g(4) + ' ' + g(5)).trim() });
    }
    return out;
  }
  const isP6 = (n) => /^ปพ\.6 /.test(n);
  const NO_CELL = (n) => /^เล่ม /.test(n) ? { r: 4, c: P.colToNum('CX') } : /^ปพ\.6 /.test(n) ? { r: 8, c: 11 } : null;

  function renderToolbar() {
    const tb = $('#sheet-tools'); tb.innerHTML = '';
    const add = (label, fn, cls) => { const b = document.createElement('button'); b.className = 'btn ' + (cls || ''); b.textContent = label; b.onclick = fn; tb.appendChild(b); return b; };
    const n = S.active;
    if (isP6(n)) {
      const list = roster(), pos = NO_CELL(n);
      const sel = document.createElement('select'); sel.className = 'docsel'; sel.setAttribute('aria-label', 'เลือกนักเรียน');
      const cur = S.wb.get(S.wb.sheet(n), pos.r, pos.c);
      sel.innerHTML = list.length ? list.map(x => `<option value="${x.n}"${x.n === cur ? ' selected' : ''}>${x.n}. ${esc(x.name)}</option>`).join('') : '<option>ยังไม่มีรายชื่อนักเรียน</option>';
      sel.onchange = () => { setCell(n, pos.r, pos.c, +sel.value); renderSheet(true); };
      tb.appendChild(sel);
      add('พิมพ์คนนี้', () => printSheets([n]), 'primary');
      add('พิมพ์ทั้งห้อง', () => printAllStudents(n));
    } else {
      add('พิมพ์', () => printSheets([n]), 'primary');
    }
    if (n === 'ข้อมูลนักเรียน') { const h = document.createElement('span'); h.className = 'hint'; h.textContent = 'คลิกช่องแรก แล้ววาง (Ctrl+V) ข้อมูลจาก Excel ได้ทั้งตาราง · วันเกิดกรอก วว/ดด/ปปปป (พ.ศ.)'; tb.appendChild(h); }
    const sp = document.createElement('span'); sp.className = 'grow'; tb.appendChild(sp);
    const z = document.createElement('label'); z.className = 'zoom';
    z.innerHTML = 'ขนาด <select id="zoom"><option value="fit"' + (S.zoom === 'fit' ? ' selected' : '') + '>พอดีหน้าจอ</option>' + [0.6, 0.75, 1, 1.25, 1.5].map(v => `<option value="${v}"${+S.zoom === v ? ' selected' : ''}>${v * 100}%</option>`).join('') + '</select>';
    tb.appendChild(z);
    $('#zoom').onchange = (e) => { S.zoom = e.target.value === 'fit' ? 'fit' : +e.target.value; renderSheet(true); };
  }

  // ---------------------------------------------------------------- ค่าที่กรอก
  const thDate = (s) => {   // วว/ดด/ปปปป (พ.ศ. หรือ ค.ศ.) -> เลขวันที่แบบ Excel
    const m = /^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/.exec(s); if (!m) return null;
    let y = +m[3]; if (y < 2400) y += 543;   // ไฟล์ต้นฉบับเก็บวันที่เป็นปี พ.ศ. ในช่องปีของ Excel (เช่น 16/5/2566 = ปี 2566)
    const t = Date.UTC(y, +m[2] - 1, +m[1]); if (isNaN(t)) return null;
    return Math.round((t - Date.UTC(1899, 11, 30)) / 864e5);
  };
  function parseInput(name, c, text) {
    let t = String(text).replace(/\r/g, '').trim();
    if (t === '') return null;
    if (name === 'ข้อมูลนักเรียน' && c === 13) { const d = thDate(t); if (d !== null) return d; }
    if (/^-?\d+(\.\d+)?$/.test(t) && !(name === 'ข้อมูลนักเรียน' && /^0\d/.test(t))) return Number(t);   // เลขนำหน้า 0 (รหัส/บัตร) เก็บเป็นข้อความ
    return t;
  }
  function setCell(name, r, c, val) {
    const ref = P.numToCol(c) + r;
    S.wb.setValue(name, ref, val);
    (S.userEdits[name] = S.userEdits[name] || {})[ref] = val;
    // ข้อมูลนักเรียน: ชีทอื่นนับว่ามีนักเรียนเมื่อคอลัมน์ A (รหัสโรงเรียน) ไม่ว่าง — ใส่ให้อัตโนมัติ
    if (name === 'ข้อมูลนักเรียน' && r >= 3 && c !== 1 && val !== null && !S.wb.get(S.wb.sheet(name), r, 1)) {
      S.wb.setValue(name, 'A' + r, '1'); S.userEdits[name]['A' + r] = '1';
    }
    markDirty();
  }

  let editing = null;
  function endEdit(commit) {
    if (!editing) return null;
    const { td, input, r, c, name } = editing; editing = null;
    const val = commit ? parseInput(name, c, input.value) : undefined;
    return { td, r, c, name, val };
  }
  function nextEditable(name, r, c, dr, dc) {
    const sh = S.wb.sheet(name), a = R.parseRange(viewArea(name));
    for (let i = 0; i < 400; i++) {
      r += dr; c += dc;
      if (c > a.c2) { c = 1; r++; } else if (c < 1) { c = a.c2; r--; }
      if (r < 1 || r > a.r2) return null;
      if (canEdit(name, r, c, sh.cells.get(P.KEY(r, c))) && !S.view.locked.has(name + '!' + P.numToCol(c) + r)) {
        const td = $(`#sheet-host td[data-r="${r}"][data-c="${c}"]`); if (td && td.classList.contains('inp')) return td;
      }
    }
    return null;
  }
  function startEdit(td) {
    if (editing) commitEdit();
    const name = S.active, r = +td.dataset.r, c = +td.dataset.c;
    const sh = S.wb.sheet(name);
    let raw = S.wb.get(sh, r, c); if (raw && raw.isRange) raw = null;
    let shown = td.textContent;
    const input = document.createElement('input'); input.className = 'cell-editor'; input.type = 'text'; input.value = raw === null || raw === undefined ? '' : (name === 'ข้อมูลนักเรียน' && c === 13 && typeof raw === 'number' ? shown : String(raw));
    td.classList.add('editing'); td.textContent = ''; td.appendChild(input);
    editing = { td, input, r, c, name };
    input.focus(); input.select();
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') { e.preventDefault(); endEdit(false); renderSheet(true); return; }
      if (e.key === 'Enter' || e.key === 'Tab' || e.key === 'ArrowDown' || e.key === 'ArrowUp') {
        e.preventDefault();
        const dir = e.key === 'Tab' ? (e.shiftKey ? [0, -1] : [0, 1]) : e.key === 'ArrowUp' || (e.key === 'Enter' && e.shiftKey) ? [-1, 0] : [1, 0];
        commitEdit(dir);
      }
    });
    input.addEventListener('paste', (e) => {
      const text = (e.clipboardData || window.clipboardData).getData('text');
      if (!/[\t\n]/.test(text.replace(/\n$/, ''))) return;   // วางข้อความธรรมดา: ปล่อยตามปกติ
      e.preventDefault(); pasteBlock(name, r, c, text);
    });
    input.addEventListener('blur', () => { if (editing && editing.input === input) setTimeout(() => { if (editing && editing.input === input) commitEdit(); }, 0); });
  }
  function commitEdit(dir) {
    const e = endEdit(true); if (!e) return;
    const sh = S.wb.sheet(e.name); const old = S.wb.get(sh, e.r, e.c);
    const same = (e.val === null && (old === null || old === '')) || e.val === old;
    if (!same) setCell(e.name, e.r, e.c, e.val);
    renderSheet(true);
    if (dir) { const nx = nextEditable(e.name, e.r, e.c, dir[0], dir[1]); if (nx) startEdit(nx); }
  }
  function pasteBlock(name, r0, c0, text) {
    const rows = text.replace(/\r/g, '').replace(/\n$/, '').split('\n').map(l => l.split('\t'));
    endEdit(false); let n = 0; const sh = S.wb.sheet(name);
    rows.forEach((row, i) => row.forEach((cell, j) => {
      const r = r0 + i, c = c0 + j;
      if (!canEdit(name, r, c, sh.cells.get(P.KEY(r, c))) || S.view.locked.has(name + '!' + P.numToCol(c) + r)) return;
      setCell(name, r, c, parseInput(name, c, cell)); n++;
    }));
    renderSheet(true); toast('วางข้อมูล ' + n + ' ช่อง', 'ok');
  }

  // ---------------------------------------------------------------- พิมพ์
  // เว็บตัวอย่าง: ลายน้ำทุกหน้าจอและทุกหน้าที่พิมพ์
  function wmLayer(n) { return CFG.WATERMARK ? `<div class="wm-layer" aria-hidden="true">${`<span>${esc(CFG.WATERMARK)}</span>`.repeat(n)}</div>` : ''; }
  const PX = 96 / 25.4;
  function usedRows(name) { return R.parseRange(viewArea(name)).r2; }
  // แบ่งชีทเป็นหน้ากระดาษ A4: ย่อให้พอดีความกว้าง แล้วตัดหน้าตามเส้นแบ่งหน้าของไฟล์หรือเมื่อเต็มหน้า
  function pagesFor(name) {
    const sh = S.wb.sheet(name), d = sh.def, pr = d.print || {};
    const land = pr.orient === 'landscape';
    const W = (land ? 297 : 210) * PX, H = (land ? 210 : 297) * PX;
    const m = (pr.m && pr.m.length === 4 ? pr.m : [0.7, 0.7, 0.75, 0.75]).map(v => Math.max(v, 0.5));   // ขอบกระดาษอย่างน้อย 0.5 นิ้ว ไม่ให้เนื้อหาชิดขอบ
    const availW = W - (m[0] + m[1]) * 96, availH = H - (m[2] + m[3]) * 96;
    const a = pr.area ? R.parseRange(pr.area) : { r1: 1, c1: 1, r2: usedRows(name), c2: R.parseRange(viewArea(name)).c2 };
    const colHid = (c) => !!(d.cw[c] && d.cw[c].h);
    const colW = (c) => colHid(c) ? 0 : Math.round(((d.cw[c] && d.cw[c].w) || d.dcw) * 7 + 5);
    const rowHid = (r) => !!(d.rh[r] && d.rh[r].x);
    const rowH = (r) => rowHid(r) ? 0 : Math.round(((d.rh[r] && d.rh[r].h) || d.drh) * 4 / 3);
    let tw = 0; for (let c = a.c1; c <= a.c2; c++) tw += colW(c);
    let th = 0; for (let r = a.r1; r <= a.r2; r++) th += rowH(r);
    let zoom = pr.scale && !pr.fitPage ? pr.scale / 100 : 1;
    if (tw * zoom > availW) zoom = availW / tw;
    if (pr.fitPage && pr.fh === 1 && th * zoom > availH) zoom = Math.min(zoom, availH / th);
    zoom = Math.max(0.1, zoom * 0.985);
    const brk = new Set(pr.breaks || []), pages = [];
    if (!brk.size && th * zoom > availH && th * zoom <= availH * 1.15) zoom = availH / th * 0.985;   // เกินหน้าเล็กน้อย: ย่อให้พอดี 1 หน้า
    let start = a.r1, acc = 0;
    for (let r = a.r1; r <= a.r2; r++) {
      const h = rowH(r) * zoom;
      if (acc + h > availH && r > start) { pages.push([start, r - 1]); start = r; acc = 0; }
      acc += h;
      if (brk.has(r) && r < a.r2) { pages.push([start, r]); start = r + 1; acc = 0; }
    }
    if (start <= a.r2) pages.push([start, a.r2]);
    return pages.map(([r1, r2]) => {
      const out = sheetHtmlFor(name, false, P.numToCol(a.c1) + r1 + ':' + P.numToCol(a.c2) + r2);
      // ขยายความสูงแถวให้เต็มหน้า A4 (หน้าที่เนื้อหาสั้นกว่าหน้ากระดาษ) — ตัวอักษรขนาดเดิม แถวห่างขึ้นตามแบบฟอร์มต้นฉบับ
      out.html = out.html.replace(/>\(กรอก[^<]*\)</g, '>(……………………………………)<').replace(/>กรอก[^<]*</g, '><');   // ข้อความตัวอย่าง “กรอก…” ที่ยังไม่ได้กรอก ไม่พิมพ์ออกมา
      let hh = 0; for (let r = r1; r <= r2; r++) hh += rowH(r);
      const k = Math.max(1, Math.min(1.7, availH * 0.985 / Math.max(1, hh * zoom)));
      const fixedRow = /^ปก ปพ\.5 รายชั้น/.test(name) ? 52 : /^ปก ปพ\.5 รายวิชา/.test(name) ? 36 : 0;   // แถวช่องสี่เหลี่ยม อนุมัติ/ไม่อนุมัติ: ไม่ยืด ให้ขนาดเท่ากันทุกปก
      if (k > 1.02) out.html = out.html.replace(/(<tr data-r=")(\d+)(" style="height:)([\d.]+)px/g, (mm, x, rr, y, h) => x + rr + y + (+rr === fixedRow ? h : (parseFloat(h) * k).toFixed(1)) + 'px');
      return { html: out.html, zoom, land, m, fs: /^ปพ\.6 /.test(name) ? 0.7 : 1, hc: !!pr.hc || tw * zoom < availW * 0.9 && false };
    });
  }
  function printAllStudents(name) {
    const list = roster(); if (!list.length) { toast('ยังไม่มีรายชื่อนักเรียน', 'warn'); return; }
    const pos = NO_CELL(name), prev = S.wb.get(S.wb.sheet(name), pos.r, pos.c);
    const blocks = [];
    try { for (const s of list) { S.wb.setValue(name, P.numToCol(pos.c) + pos.r, s.n); blocks.push(...docPages(name)); } }
    finally { S.wb.setValue(name, P.numToCol(pos.c) + pos.r, prev); }
    showPreview(blocks);
  }
  // หน้าพิมพ์ของชีท: ชีทที่มีหน้าเสริม (เช่น เวลาเรียนรายวิชาต่อท้าย ปพ.5/ปพ.6) ให้ ui.js ประกอบให้
  const docPages = (n) => (window.PPUI && PPUI.customPages && PPUI.customPages(n)) || pagesFor(n);
  function printSheets(names) { showPreview(names.flatMap(docPages)); }

  function showPreview(blocks) {
    if (!blocks.length) return;
    const root = $('#print-root');
    root.innerHTML = blocks.map(b => `<section class="pp-page ${b.land ? 'land' : 'port'}"${b.fitW ? ' data-fitw="1"' : ''} style="padding:${b.m[2]}in ${b.m[1]}in ${b.m[3]}in ${b.m[0]}in"><div class="pp-zoom" style="zoom:${b.zoom};--bw:${1 / b.zoom}px${b.fs && b.fs !== 1 ? ';--fs:' + b.fs : ''}${b.hc ? ';margin:0 auto;width:max-content' : ''}">${b.html}</div>${wmLayer(14)}</section>`).join('');
    let bar = $('#pv-bar');
    if (!bar) {
      document.body.insertAdjacentHTML('beforeend', `<div id="pv-bar" role="toolbar" aria-label="ตัวอย่างก่อนพิมพ์"><b>ตัวอย่างก่อนพิมพ์</b><span class="hint" id="pv-n"></span><span class="pv-act"><button class="btn primary" id="pv-print">พิมพ์</button><button class="btn" id="pv-close">ปิด</button></span></div>`);
      bar = $('#pv-bar');
    }
    const y = window.scrollY;
    const close = () => { document.body.classList.remove('previewing', 'printing'); root.innerHTML = ''; document.removeEventListener('keydown', onKey); window.scrollTo(0, y); };
    const onKey = (e) => { if (e.key === 'Escape') close(); };
    $('#pv-n').textContent = blocks.length + ' หน้า';
    $('#pv-close').onclick = close;
    $('#pv-print').onclick = async () => {
      document.body.classList.add('printing');
      await new Promise(r => setTimeout(r, 60));
      window.print();
      setTimeout(() => document.body.classList.remove('printing'), 300);
    };
    document.addEventListener('keydown', onKey);
    document.body.classList.add('previewing'); window.scrollTo(0, 0);
    // หน้าที่ออกแบบเอง (เช่น เวลาเรียน): ย่อให้พอดีกว้างและสูงของกระดาษ
    $$('.pp-page[data-fitw]', root).forEach(sec => {
      const z = sec.firstElementChild, cs = getComputedStyle(sec);
      const aw = sec.clientWidth - parseFloat(cs.paddingLeft) - parseFloat(cs.paddingRight), ah = sec.clientHeight - parseFloat(cs.paddingTop) - parseFloat(cs.paddingBottom);
      z.style.zoom = 1; z.style.width = 'max-content'; z.style.margin = '0 auto';
      const r = z.getBoundingClientRect(), k = Math.max(0.3, Math.min(1, aw / r.width, ah / r.height) * 0.99);
      z.style.zoom = k; z.style.setProperty('--bw', (1 / k) + 'px');
    });
  }

  // ---------------------------------------------------------------- เริ่มต้น
  function boot() {
    if (CFG.WATERMARK) document.body.insertAdjacentHTML('beforeend', wmLayer(80));
    if (CFG.SAMPLE_DIR) {   // เว็บตัวอย่าง: ปุ่มคืนค่าข้อมูลตัวอย่าง
      $('#save-now').insertAdjacentHTML('beforebegin', '<button id="sample-reset" class="btn" title="ล้างการแก้ไขทั้งหมด กลับเป็นข้อมูลตัวอย่างเดิม">คืนค่าตัวอย่าง</button>');
      $('#sample-reset').onclick = async () => {
        const yes = await modal('คืนค่าข้อมูลตัวอย่าง', '<p>ล้างการแก้ไขทั้งหมดในเบราว์เซอร์นี้ แล้วกลับเป็นข้อมูลตัวอย่างเดิม</p>', [{ label: 'ยกเลิก', value: false }, { label: 'คืนค่า', value: true, primary: true }]);
        if (!yes) return;
        S.dirty = false; clearTimeout(S.saveTimer); backend.reset(); location.reload();
      };
    }
    $('#login-form').addEventListener('submit', (e) => { e.preventDefault(); doLogin($('#login-code').value, $('#login-remember').checked); });
    $('#logout').onclick = logout;
    $('#logo-file').addEventListener('change', (e) => { loadLogo(e.target.files[0]); e.target.value = ''; });
    $('#save-now').onclick = () => save(true);
    $('#doc-close').onclick = closeDoc;
    $('#open-docs').onclick = () => { if (window.PPUI) PPUI.go('docs'); };
    $('#sheet-host').addEventListener('click', (e) => { const td = e.target.closest('td.inp'); if (td && !td.classList.contains('editing')) startEdit(td); });
    window.addEventListener('beforeunload', (e) => { if (S.dirty) { e.preventDefault(); e.returnValue = ''; } });
    window.addEventListener('online', () => { if (S.dirty) save(); });
    window.addEventListener('resize', () => { if (S.wb && S.zoom === 'fit' && !editing && !$('#docviewer').hidden) renderSheet(true); });
    if (backend.demo) { $('#demo-note').hidden = false; $('#login-code').value = 'DEMO'; }
    const saved = LS.get('pp_code');
    if (saved && !backend.demo) { $('#login-code').value = saved; $('#login-remember').checked = true; doLogin(saved, true); }
  }
  document.addEventListener('DOMContentLoaded', boot);
  window.PPApp = {
    S, CFG, toast, modal, roomLabel, openSheet, closeDoc, pagesFor, roster, setCell, markDirty, save, sheetGroups, subjName,
    printSheets, showPreview, printAllStudents, exportBackup, pickBackup, parseInput, canEdit, openGrade, manageRooms, editItems, renameItem, pickLogo: () => $('#logo-file').click(), removeLogo, issueTargets,
  };
})();
