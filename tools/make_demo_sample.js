/* สร้างข้อมูลตัวอย่างของเว็บตัวอย่าง (demo/sample/pN.json) — นักเรียนสมมติ 40 คน ทุกชั้น
 *   node tools/make_demo_sample.js
 * ชื่อ-นามสกุลทั้งหมดเป็นชื่อสมมติ (นามสกุลเป็นคำคุณธรรม ไม่ใช่ข้อมูลบุคคลจริง)
 * กรอกลงเวิร์กบุ๊กด้วย engine.js โดยตรง แล้วเก็บเฉพาะช่องที่กรอกในรูปแบบเดียวกับที่แอปบันทึก (cells) */
const fs = require('fs'), path = require('path');
const P = require('../js/engine.js');
const ROOT = path.join(__dirname, '..'), OUT = path.join(ROOT, 'demo', 'sample');
const tpl = JSON.parse(fs.readFileSync(path.join(ROOT, 'data', 'template.json'), 'utf8'));

const BOYS = ['กิตติพัฒน์', 'ธนกร', 'ภูมิพัฒน์', 'ณัฐวุฒิ', 'ปัณณวิชญ์', 'ศุภกร', 'พีรพัฒน์', 'ธีรภัทร', 'วรเมธ', 'อนุชิต', 'ชยพล', 'ปกรณ์', 'ภาคิน', 'รัชชานนท์', 'สิรภพ', 'ธนภัทร', 'กฤตเมธ', 'จิรายุ', 'ณภัทร', 'ปิยวัฒน์', 'วชิรวิทย์', 'อชิตะ', 'ภูริณัฐ', 'ตรัยคุณ'];
const GIRLS = ['ณัฐธิดา', 'กัญญาณัฐ', 'พิมพ์ชนก', 'ปุณยวีร์', 'ชนัญชิดา', 'ธัญชนก', 'อริสรา', 'ศิริญากร', 'วรัญญา', 'พัชราภา', 'กมลชนก', 'ปภาวรินทร์', 'ณิชาภัทร', 'สุพิชฌาย์', 'อัญชิสา', 'ภัทรวดี', 'ชุติกาญจน์', 'เบญญาภา', 'ธนพร', 'จิดาภา', 'มนัสนันท์', 'รินรดา', 'ปาณิสรา', 'ขวัญข้าว'];
const SURN = ['รักการอ่าน', 'ใฝ่เรียนรู้', 'ตั้งใจเรียนดี', 'มีน้ำใจงาม', 'ขยันหมั่นเพียร', 'ยิ้มแย้มเสมอ', 'ช่างสังเกตดี', 'มีวินัยดี', 'ใจดีเสมอ', 'รักความสะอาด', 'ซื่อตรงดี', 'อ่อนน้อมงาม', 'กตัญญูดี', 'รักษ์โลกดี', 'คิดดีทำดี', 'พอเพียงงาม', 'เรียนรู้ไว', 'สามัคคีดี', 'มุ่งมั่นดี', 'มีระเบียบดี', 'รับผิดชอบดี', 'ใฝ่ดีเสมอ', 'สุภาพงาม', 'อดทนดี', 'ประหยัดดี', 'รักเพื่อนดี', 'กล้าคิดดี', 'ร่าเริงดี', 'ตรงเวลาดี', 'ช่วยเหลือดี', 'สร้างสรรค์ดี', 'ใจเย็นดี', 'มีเหตุผลดี', 'รักธรรมชาติ', 'แบ่งปันดี', 'มีสติดี', 'ขยันอ่านดี', 'รักกีฬาดี', 'ใฝ่ศิลป์ดี', 'รักดนตรีดี'];
const STAFF = {
  B4: 'สำนักงานเขตพื้นที่การศึกษาประถมศึกษาตัวอย่าง เขต 1', B5: 'ตำบลตัวอย่าง', B6: 'อำเภอตัวอย่าง', B7: 'จังหวัดตัวอย่าง',
  B8: 'สำนักงานเขตพื้นที่การศึกษาประถมศึกษาตัวอย่าง เขต 1', B9: 'นายประเสริฐ นำทางดี', B10: 'ผู้อำนวยการโรงเรียน',
  B11: 'นางสาวสมใจ ใฝ่วิชาการ', B12: 'นายวิชัย บริหารดี',
};
const CLASS_T = ['นางสาวมาลี สอนดีมาก', 'นางสุดา ใจดีสอนเก่ง', 'นางสาวพรทิพย์ รักลูกศิษย์', 'นายสมชาย ครูตัวอย่าง', 'นางวันดี เมตตาศิษย์', 'นางสาวกานดา สอนสนุก'];
// [รหัสนำหน้า, ชื่อวิชา, กลุ่มสาระ, ชั่วโมง/ปี ป.1-3, ป.4-6, ครูประจำวิชา]
const SUBJECTS = [
  ['ท', 'ภาษาไทย', 'ภาษาไทย', 200, 160, 'นางสาวอรุณี รักภาษา'],
  ['ค', 'คณิตศาสตร์', 'คณิตศาสตร์', 200, 160, 'นายธนา เลขดีมาก'],
  ['ว', 'วิทยาศาสตร์และเทคโนโลยี', 'วิทยาศาสตร์และเทคโนโลยี', 80, 120, 'นางสาวนภา ทดลองดี'],
  ['ส', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 80, 80, 'นายชาติ รักถิ่นดี'],
  ['ส', 'ประวัติศาสตร์', 'สังคมศึกษา ศาสนา และวัฒนธรรม', 40, 40, 'นายชาติ รักถิ่นดี'],
  ['พ', 'สุขศึกษาและพลศึกษา', 'สุขศึกษาและพลศึกษา', 80, 80, 'นายพละ แข็งแรงดี'],
  ['ศ', 'ศิลปะ', 'ศิลปะ', 80, 80, 'นางสาวศิลปา วาดสวย'],
  ['ง', 'การงานอาชีพ', 'การงานอาชีพ', 40, 80, 'นางสายใจ ทำงานดี'],
  ['อ', 'ภาษาอังกฤษ', 'ภาษาต่างประเทศ', 80, 80, 'Ms. Anna Example'],
];
const col = (n) => P.numToCol(n);
const rng = (seed) => { let s = seed >>> 0; return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; }; };
const shuffle = (a, r) => { a = a.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(r() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
// ไฟล์ต้นฉบับเก็บวันที่เป็นปี พ.ศ. ในช่องปีของ Excel (ปี 2559 = เลขวันที่ของปี 2559)
const serial = (y, m, d) => Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 864e5);

function build(g) {
  const r = rng(2551 + g * 97), N = 40;
  const boys = shuffle(BOYS, r).slice(0, 20), girls = shuffle(GIRLS, r).slice(0, 20), sur = shuffle(SURN, r);
  const kids = shuffle([...boys.map(n => ['ช', 'เด็กชาย', n]), ...girls.map(n => ['ญ', 'เด็กหญิง', n])], r)
    .map(([sex, pre, n], i) => ({ sex, pre, first: n, last: sur[i], a: Math.min(0.97, Math.max(0.42, 0.72 + (r() + r() + r() - 1.5) * 0.24)) }));
  const wb = new P.Workbook(tpl);
  const edits = {};
  const set = (sh, ref, v) => { wb.setValue(sh, ref, v); (edits[sh] = edits[sh] || {})[ref] = v; };
  const rc = (row, c) => col(c) + row;
  const val = (sh, row, c) => { const s = wb.sheet(sh); let v = wb.get(s, row, c); if (v && v.isRange) v = wb.scalar(v); return v; };

  for (const [k, v] of Object.entries(STAFF)) set('ข้อมูลพื้นฐาน', k, v);
  set('ข้อมูลพื้นฐาน', 'B13', CLASS_T[g - 1]);

  // นักเรียน (แถว 3-42 ของชีทข้อมูลนักเรียน) — A = รหัสโรงเรียน (ต้องไม่ว่าง ชีทอื่นจึงนับว่ามีนักเรียน)
  kids.forEach((k, i) => {
    const row = 3 + i, S = 'ข้อมูลนักเรียน';
    set(S, 'A' + row, '1'); set(S, 'C' + row, '0000000' + String(g * 1000 + i + 1).padStart(6, '0'));
    set(S, 'D' + row, 'ประถมศึกษาปีที่ ' + g); set(S, 'E' + row, 1); set(S, 'F' + row, String(g) + String(i + 1).padStart(3, '0'));
    set(S, 'G' + row, k.sex); set(S, 'H' + row, k.pre); set(S, 'I' + row, k.first); set(S, 'J' + row, k.last);
    set(S, 'M' + row, serial(2566 - 5 - g, 1 + Math.floor(r() * 12), 1 + Math.floor(r() * 28)));
    set(S, 'BK' + row, Math.round(14 + g * 2.6 + r() * 5)); set(S, 'BL' + row, Math.round(100 + g * 6 + r() * 8));
  });

  // รายวิชา (แถว 6-14)
  const subs = SUBJECTS.map((a, j) => ({ a, slot: j + 1, row: 6 + j }));
  const seq = {};
  for (const s of subs) {
    const [pre, name, grp, h1, h2, teacher] = s.a;
    seq[pre] = (seq[pre] || 0) + 1;
    set('รายวิชา', 'B' + s.row, `${pre}1${g}10${seq[pre]}`); set('รายวิชา', 'C' + s.row, name + ' ' + g);
    set('รายวิชา', 'D' + s.row, g <= 3 ? h1 : h2); set('รายวิชา', 'E' + s.row, grp); set('รายวิชา', 'F' + s.row, teacher);
  }

  // คะแนนรายวิชา: ตัวชี้วัด 6 ตัว/ภาค (ผ/มผ) · คะแนนระหว่างเรียน 4 งาน (10 คะแนน) · สอบปลายภาค (30)
  const TERM = [{ ind: 3, work: P.colToNum('CA'), exam: P.colToNum('DA') }, { ind: P.colToNum('AO'), work: P.colToNum('DE'), exam: P.colToNum('EE') }];
  for (const s of subs) {
    const sh = 'วิชาที่ ' + s.slot, off = (r() - 0.5) * 0.14;
    TERM.forEach((t, ti) => {
      for (let j = 0; j < 6; j++) set(sh, rc(5, t.ind + j), 'ตัวชี้วัดที่ ' + (j + 1));
      for (let j = 0; j < 4; j++) { set(sh, rc(5, t.work + j), 'งานที่ ' + (j + 1)); set(sh, rc(6, t.work + j), 10); }
      set(sh, rc(5, t.exam), 'สอบปลายภาค');
      kids.forEach((k, i) => {
        const row = 7 + i, lv = k.a + off + ti * 0.02;
        for (let j = 0; j < 6; j++) set(sh, rc(row, t.ind + j), lv + (r() - 0.5) * 0.3 < 0.5 ? 'มผ' : 'ผ');
        for (let j = 0; j < 4; j++) set(sh, rc(row, t.work + j), Math.round(10 * Math.min(1, Math.max(0.3, lv + (r() - 0.5) * 0.2))));
        set(sh, rc(row, t.exam), Math.round(30 * Math.min(1, Math.max(0.3, lv + (r() - 0.5) * 0.18))));
      });
    });
  }

  // คุณลักษณะ / ค่านิยม / สมรรถนะ (0-3) และกิจกรรมพัฒนาผู้เรียน (ผ่าน/ไม่ผ่าน)
  const level = (k) => { const x = k.a + (r() - 0.5) * 0.5; return x < 0.46 ? 1 : x < 0.78 ? 2 : 3; };
  for (const t of [1, 2]) {
    kids.forEach((k, i) => {
      const row = 5 + i, fail = k.a < 0.5 && r() < 0.3;
      for (let c = 3; c <= 10; c++) set('คุณลักษณะ ' + t, rc(row, c), level(k));
      for (let c = 17; c <= 21; c++) set('คุณลักษณะ ' + t, rc(row, c), level(k));
      for (let c = 3; c <= 14; c++) set('ค่านิยม ' + t, rc(row, c), level(k));
      for (let c = 19; c <= 23; c++) set('ค่านิยม ' + t, rc(row, c), level(k));
      // กิจกรรม: C/D แนะแนว · E/F ลูกเสือ-เนตรนารี · G ชื่อชุมนุม · H/I ชุมนุม · J/K เพื่อสังคมฯ (ซ้าย = ไม่ผ่าน ขวา = ผ่าน)
      for (const [bad, ok] of [[3, 4], [5, 6], [8, 9], [10, 11]]) set('กิจกรรม ' + t, rc(row, fail && bad === 5 ? bad : ok), '/');
      set('กิจกรรม ' + t, rc(row, 7), 'ชุมนุมคอมพิวเตอร์');
    });
  }

  // น้ำหนัก-ส่วนสูง 4 ครั้ง (C/D, E/F, G/H, I/J)
  kids.forEach((k, i) => {
    const row = 7 + i, w0 = 14 + g * 2.6 + r() * 5, h0 = 100 + g * 6 + r() * 8;
    for (let j = 0; j < 4; j++) { set('น้ำหนัก-ส่วนสูง', rc(row, 3 + 2 * j), Math.round((w0 + j * 0.6) * 10) / 10); set('น้ำหนัก-ส่วนสูง', rc(row, 4 + 2 * j), Math.round((h0 + j * 0.8) * 10) / 10); }
  });

  // เวลาเรียน: ทุกวันเรียนตามปฏิทินมา (1) บางวัน ป่วย/ลา/ขาด — เดือนละ 1 บล็อก (แถว 3 = วันที่ แถว 4 = วันเรียนจากชีทปฏิทิน)
  const AT = 'เวลาเรียน', ats = wb.sheet(AT);
  const starts = []; for (let c = 3; c <= ats.maxC; c++) { const v = val(AT, 3, c); if (v === 1 || v === '1') starts.push(c); }
  kids.forEach((k, i) => {
    const row = 5 + i, miss = 0.06 * (1.3 - k.a);
    for (const c0 of starts) for (let d = 0; d < 31; d++) {
      const c = c0 + d, hd = val(AT, 4, c), dn = val(AT, 3, c);
      if (String(dn) !== String(d + 1) || hd === null || hd === '' || P.isErr(hd)) continue;
      const x = r();
      set(AT, rc(row, c), x < miss * 0.5 ? 'ป' : x < miss * 0.8 ? 'ล' : x < miss ? 'ข' : 1);
    }
  });
  return { g, edits };
}

fs.mkdirSync(OUT, { recursive: true });
for (const g of [1, 2, 3, 4, 5, 6]) {
  const { edits } = build(g);
  const doc = { v: 3, grade: g, room: 1, cells: edits, savedAt: new Date(0).toISOString() };
  fs.writeFileSync(path.join(OUT, `p${g}.json`), JSON.stringify(doc));
  console.log(`p${g}: ${Object.values(edits).reduce((a, m) => a + Object.keys(m).length, 0)} cells`);
}
