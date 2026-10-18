// กติกาการคำนวณ หลักสูตรแกนกลางการศึกษาขั้นพื้นฐาน พ.ศ. 2551 (ประถมศึกษา)

export const TERMS = [1, 2, 'รวม'];
export const termName = t => (t === 'รวม' ? 'สรุปรายปี' : `ภาคเรียนที่ ${t}`);

export const TRAITS = ['รักชาติ ศาสน์ กษัตริย์', 'ซื่อสัตย์สุจริต', 'มีวินัย', 'ใฝ่เรียนรู้', 'อยู่อย่างพอเพียง', 'มุ่งมั่นในการทำงาน', 'รักความเป็นไทย', 'มีจิตสาธารณะ'];
export const TRAITS_SHORT = ['รักชาติฯ', 'ซื่อสัตย์', 'มีวินัย', 'ใฝ่เรียนรู้', 'พอเพียง', 'มุ่งมั่น', 'ความเป็นไทย', 'จิตสาธารณะ'];
export const RTW = ['การอ่าน', 'การคิดวิเคราะห์', 'การเขียน', 'การสื่อสาร', 'การสรุปความ'];
export const VALUES = [
  'มีความรักชาติ ศาสน์ กษัตริย์',
  'ซื่อสัตย์ เสียสละ อดทน มีอุดมการณ์ในสิ่งที่ดีงามเพื่อส่วนรวม',
  'กตัญญูต่อพ่อแม่ ผู้ปกครอง ครูบาอาจารย์',
  'ใฝ่หาความรู้ หมั่นศึกษาเล่าเรียนทั้งทางตรงและทางอ้อม',
  'รักษาวัฒนธรรมประเพณีไทยอันงดงาม',
  'มีศีลธรรม รักษาความสัตย์ หวังดีต่อผู้อื่น เผื่อแผ่และแบ่งปัน',
  'เข้าใจเรียนรู้การเป็นประชาธิปไตยอันมีพระมหากษัตริย์ทรงเป็นประมุขที่ถูกต้อง',
  'มีระเบียบวินัย เคารพกฎหมาย ผู้น้อยรู้จักการเคารพผู้ใหญ่',
  'มีสติรู้ตัวอยู่เสมอ ประพฤติตามพระราชดำรัสสอนของพระมหากษัตริย์',
  'รู้จักดำรงตนอยู่โดยใช้หลักปรัชญาของเศรษฐกิจพอเพียง',
  'มีความเข้มแข็งทั้งร่างกายและจิตใจ ไม่ยอมแพ้ต่ออำนาจฝ่ายต่ำ',
  'คำนึงถึงผลประโยชน์ของส่วนรวมและของชาติมากกว่าผลประโยชน์ของตนเอง',
];
export const COMPS = ['ความสามารถในการสื่อสาร', 'ความสามารถในการคิด', 'ความสามารถในการแก้ปัญหา', 'ความสามารถในการใช้ทักษะชีวิต', 'ความสามารถในการใช้เทคโนโลยี'];
export const LV_TRAIT = ['ไม่ผ่าน', 'ผ่าน', 'ดี', 'ดีเยี่ยม'];
export const LV_COMP = ['ปรับปรุง', 'พอใช้', 'ดี', 'ดีเยี่ยม'];
export const GRADE_LIST = ['มส', 'ร', 0, 1, 1.5, 2, 2.5, 3, 3.5, 4];
export const MONTHS = ['พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม', 'มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน'];
export const MONTHS_TH = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน', 'พฤษภาคม', 'มิถุนายน', 'กรกฎาคม', 'สิงหาคม', 'กันยายน', 'ตุลาคม', 'พฤศจิกายน', 'ธันวาคม'];
export const GROUPS = ['ภาษาไทย', 'คณิตศาสตร์', 'วิทยาศาสตร์และเทคโนโลยี', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 'สุขศึกษาและพลศึกษา', 'ศิลปะ', 'การงานอาชีพและเทคโนโลยี', 'ภาษาต่างประเทศ', 'รายวิชาเพิ่มเติม'];
export const LEVELS = ['ประถมศึกษาปีที่ 1', 'ประถมศึกษาปีที่ 2', 'ประถมศึกษาปีที่ 3', 'ประถมศึกษาปีที่ 4', 'ประถมศึกษาปีที่ 5', 'ประถมศึกษาปีที่ 6'];

export const isNum = v => v !== '' && v != null && !isNaN(Number(v));
export const r2 = n => Math.round((n + Number.EPSILON) * 100) / 100;
export const fmt = (n, d = 2) => (isNum(n) ? Number(n).toFixed(d) : '');

export function gradeFromPct(p) {
  return p < 50 ? 0 : p < 55 ? 1 : p < 60 ? 1.5 : p < 65 ? 2 : p < 70 ? 2.5 : p < 75 ? 3 : p < 80 ? 3.5 : 4;
}
export const gradeText = g => (g === '' || g == null ? '' : String(g));

export const defaultCfg = () => ({
  inds: Array.from({ length: 10 }, () => ''),
  max: [10, 10, 10, 10, 10, 10],
  rb: 70, rf: 30,
});
export function ensureCfg(subj, term) {
  subj.cfg ??= {};
  subj.cfg[term] ??= defaultCfg();
  return subj.cfg[term];
}

/** ผลของรายวิชาหนึ่งภาคเรียนของนักเรียนหนึ่งคน */
export function termResult(subj, rec, term) {
  const c = ensureCfg(subj, term);
  const r = rec?.[term] ?? {};
  const p = (r.p ?? []).slice(0, c.inds.length);
  const pass = p.filter(x => x === 'ผ').length;
  const fail = p.filter(x => x === 'มผ').length;
  const ind = pass + fail === 0 ? '' : fail > 0 ? 'มผ' : 'ผ';
  const M = c.max.reduce((a, b) => a + (Number(b) || 0), 0);
  const s = (r.s ?? []).slice(0, c.max.length);
  const hasS = s.some(isNum);
  const sum = s.reduce((a, b) => a + (isNum(b) ? Number(b) : 0), 0);
  const between = M > 0 && hasS ? r2(sum / (M / c.rb)) : null;
  const fin = isNum(r.f) ? Number(r.f) : null;
  const total = between == null && fin == null ? null : r2((between ?? 0) + (fin ?? 0));
  const full = c.rb + c.rf;
  const pct = total == null ? null : r2((total * 100) / full);
  const grade = r.o ? r.o : pct == null ? '' : gradeFromPct(pct);
  return { pass, fail, ind, sum, between, fin, total, pct, grade, full, over: r.o || '' };
}

/** ผลรวมรายปี (ภาค 1 + ภาค 2) */
export function yearResult(subj, rec) {
  const a = termResult(subj, rec, 1), b = termResult(subj, rec, 2);
  const over = a.over === 'มส' || b.over === 'มส' ? 'มส' : a.over === 'ร' || b.over === 'ร' ? 'ร' : '';
  const total = a.total != null && b.total != null ? r2(a.total + b.total) : null;
  const full = a.full + b.full;
  const pct = total == null ? null : r2((total * 100) / full);
  const grade = over || (pct == null ? '' : gradeFromPct(pct));
  return { a, b, total, full, pct, grade, ind: a.ind && b.ind ? (a.ind === 'ผ' && b.ind === 'ผ' ? 'ผ' : 'มผ') : '' };
}

export function resultFor(subj, rec, term) {
  if (term === 'รวม') {
    const y = yearResult(subj, rec);
    return { total: y.total, full: y.full, pct: y.pct, grade: y.grade, ind: y.ind, between: null, fin: null };
  }
  return termResult(subj, rec, term);
}

/** ฐานนิยมของคะแนน (เท่ากันให้เอาค่าสูง) ตามสูตรในไฟล์ต้นแบบ */
export function modeLevel(arr) {
  const v = (arr ?? []).filter(isNum).map(Number);
  if (!v.length) return '';
  const cnt = {};
  v.forEach(x => (cnt[x] = (cnt[x] || 0) + 1));
  const mx = Math.max(...Object.values(cnt));
  if (mx === 1) return Math.max(...v);
  return Math.max(...Object.keys(cnt).filter(k => cnt[k] === mx).map(Number));
}

export const avg = arr => {
  const v = arr.filter(isNum).map(Number);
  return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null;
};

export const thaiNum = n => String(n).replace(/\d/g, d => '๐๑๒๓๔๕๖๗๘๙'[d]);

export function thaiDate(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (isNaN(d)) return String(iso);
  return `${d.getDate()} ${MONTHS_TH[d.getMonth()]} ${d.getFullYear() + 543}`;
}

export function ageOn(iso, ref = new Date()) {
  if (!iso) return { y: '', m: '' };
  const b = new Date(iso);
  if (isNaN(b)) return { y: '', m: '' };
  let months = (ref.getFullYear() - b.getFullYear()) * 12 + ref.getMonth() - b.getMonth();
  if (ref.getDate() < b.getDate()) months--;
  return { y: Math.floor(months / 12), m: months % 12 };
}

/** แปลงวันที่ที่วางมา (YYYY-MM-DD, dd/mm/yyyy พ.ศ./ค.ศ.) เป็น ISO */
export function parseDate(s) {
  s = String(s ?? '').trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  let d, mo, y;
  if (m) { y = +m[1]; mo = +m[2]; d = +m[3]; }
  else if ((m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})$/))) { d = +m[1]; mo = +m[2]; y = +m[3]; }
  else return '';
  if (y > 2400) y -= 543;
  return `${y}-${String(mo).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
}

export function studentName(s) { return `${s.prefix ?? ''}${s.first ?? ''} ${s.last ?? ''}`.trim(); }
