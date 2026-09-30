/* ================= CORE: đọc file raw, chuẩn hoá, tính toán ================= */
const STATUSES = ['Hoàn thành', 'Hoàn thành muộn', 'Đang làm', 'Quá hạn'];
const QBUCKETS = ['+3%', '+2%', '+1%', '0%', '-1%', '-2%', '-3%'];
const NOTE_DS = 'Công việc Doanh số (đánh giá theo chỉ tiêu doanh số, không quy % chất lượng)';
const UNKNOWN = '(Chưa xác định)';
const PROJECT_DEPT = 'Dự án';
const ALL = '(Tất cả)';
// Thứ tự sheet phòng ban giống file mẫu; phòng ban mới xếp sau theo ABC
const DEPT_ORDER = ['Admin', 'MKT', 'Kinh doanh', 'Kế toán', 'Tech', 'HCNS'];
const DEPT_CANON = {
  'admin': 'Admin', 'mkt': 'MKT', 'marketing': 'MKT', 'kinh doanh': 'Kinh doanh', 'kd': 'Kinh doanh', 'sale': 'Kinh doanh',
  'kế toán': 'Kế toán', 'ke toan': 'Kế toán', 'tech': 'Tech', 'hcns': 'HCNS', 'pm': 'PM', 'fsi': 'FSI', 'presale': 'Presale',
  'presales': 'Presale', 'sonic': 'Sonic', 'kỹ thuật': 'Kỹ thuật', 'solutions kỹ thuật': 'Solutions Kỹ thuật'
};

const norm = v => v == null ? '' : String(v).normalize('NFC').replace(/ /g, ' ').replace(/\s+/g, ' ').trim();
const lc = v => norm(v).toLowerCase();
const viCmp = (a, b) => a.localeCompare(b, 'vi', { sensitivity: 'base' });

function deptFromFileName(fn) {
  let s = norm(fn).replace(/\.(xlsx|xlsm|xls)$/i, '').replace(/[.\s]+$/, '');
  s = s.replace(/^T\s*\d{1,2}(?:[.\-_/]\d{2,4})?[\s_\-]+/i, '');   // bỏ tiền tố "T9 "
  s = s.replace(/[\s_\-]*\d+$/, '').trim();                           // "Tech 1" -> "Tech"
  const c = DEPT_CANON[s.toLowerCase()];
  if (c) return c;
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : 'Khác';
}
function monthFromFileName(fn) {
  const m = norm(fn).match(/^T\s*(\d{1,2})\b/i);
  return m ? +m[1] : null;
}

// Chuẩn hoá đánh giá chất lượng: "3%" -> "+3%", " +1 % " -> "+1%"
function normQuality(v) {
  const s = norm(v);
  if (!s) return { label: '', pct: 0, kind: 'blank' };
  const m = s.replace(/\s/g, '').match(/^([+-]?)(\d+(?:[.,]\d+)?)%$/);
  if (m) {
    let n = parseFloat(m[2].replace(',', '.')) * (m[1] === '-' ? -1 : 1);
    const label = n === 0 ? '0%' : (n > 0 ? '+' : '-') + Math.abs(n) + '%';
    return { label, pct: n / 100, kind: 'pct' };
  }
  if (/\bDS\b|doanh số/i.test(s)) return { label: s, pct: null, kind: 'ds' };
  return { label: s, pct: null, kind: 'text' };
}
function levelNum(v) { const s = norm(v); return /^\d+(\.0+)?$/.test(s) ? parseInt(s, 10) : null; }

function noteFor(level, q) {
  if (q.kind === 'ds' || lc(level) === 'doanh số') return NOTE_DS;
  if (q.kind === 'text') return `Đánh giá dạng chữ: "${q.label}" (không quy đổi %)`;
  return '';
}

function parsePeriodFromSheet(name) {
  const m = norm(name).match(/(\d{2})(\d{2})(\d{4})\s*-\s*(\d{2})(\d{2})(\d{4})/);
  if (!m) return null;
  return { from: new Date(+m[3], +m[2] - 1, +m[1]), to: new Date(+m[6], +m[5] - 1, +m[4]) };
}
const fmtDate = d => `${String(d.getDate()).padStart(2, '0')}/${String(d.getMonth() + 1).padStart(2, '0')}/${d.getFullYear()}`;

function parseCustomField(s) {
  const out = {};
  norm(s).split(/\s*\|\s*/).forEach(part => {
    const i = part.indexOf(':');
    if (i > 0) out[lc(part.slice(0, i))] = norm(part.slice(i + 1));
  });
  return out;
}

/** Đọc 1 file -> {type:'task'|'project'|'error', rows, dept, period, ...} */
async function readRawFile(file) {
  const buf = await file.arrayBuffer();
  let wb;
  try { wb = XLSX.read(buf, { type: 'array', cellDates: false }); }
  catch (e) { return { file, name: file.name, type: 'error', error: 'Không đọc được file: ' + e.message }; }
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    if (!ws || !ws['!ref']) continue;
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, blankrows: false });
    let hi = rows.findIndex((r, i) => i < 15 && r.some(c => lc(c) === 'tên công việc'));
    if (hi < 0) continue;
    const header = rows[hi].map(lc);
    const col = n => header.indexOf(n);
    const isProject = col('trường dữ liệu tùy chỉnh') >= 0 || col('tên dự án') >= 0;
    const get = (r, n) => { const i = col(n); return i >= 0 ? norm(r[i]) : ''; };
    const out = [];
    for (const r of rows.slice(hi + 1)) {
      const name = get(r, 'họ và tên người nhận việc') || get(r, 'người nhận việc');
      const status = get(r, 'trạng thái');
      const task = get(r, 'tên công việc');
      if (!task && !name && !status) continue;
      if (isProject) {
        const cf = parseCustomField(get(r, 'trường dữ liệu tùy chỉnh'));
        const hasCF = Object.keys(cf).length > 0;
        if (!name && !status && !hasCF) continue;   // dòng tiêu đề dự án / giai đoạn
        const lv = cf['mức độ công việc'] || '';
        const q = normQuality(cf['đánh giá chất lượng'] || '');
        out.push({
          project: get(r, 'tên dự án') || '(Không rõ dự án)', name: name || UNKNOWN, task, status,
          lvRaw: lv, lvNum: levelNum(lv), qRaw: q.label, q, note: noteFor(lv, q),
          phase: get(r, 'tasklist name'), goal: cf['mục tiêu bộ phận'] || '', detail: cf['chi tiết thành phần công việc'] || '',
          sales: cf['kinh doanh phụ trách'] || '', customer: cf['khách hàng'] || '', deadline: get(r, 'thời hạn'),
          done: get(r, 'hoàn thành thực tế'), cause: cf['nguyên nhân'] || '', fix: cf['phương án xử lý'] || '',
          custom: get(r, 'trường dữ liệu tùy chỉnh')
        });
      } else {
        if (!name && !status) continue;                // dòng nhóm kiểu "BÁO CÁO TUẦN:"
        const lv = get(r, 'mức độ công việc');
        const q = normQuality(get(r, 'đánh giá chất lượng'));
        out.push({ name: name || UNKNOWN, task, status, lvRaw: lv, lvNum: levelNum(lv), qRaw: q.label, q, note: noteFor(lv, q),
          deadline: get(r, 'thời hạn'), done: get(r, 'hoàn thành thực tế') });
      }
    }
    return {
      file, name: file.name, type: isProject ? 'project' : 'task', rows: out, sheet: sn,
      dept: isProject ? PROJECT_DEPT : deptFromFileName(file.name),
      period: parsePeriodFromSheet(sn), month: monthFromFileName(file.name)
    };
  }
  return { file, name: file.name, type: 'error', error: 'Không tìm thấy dòng tiêu đề có cột "Tên công việc"' };
}

/* ---------- thống kê ---------- */
function agg(rows) {
  const a = { total: rows.length, st: {}, other: 0, lv: { 1: 0, 2: 0, 3: 0 }, noLv: 0, q: {}, qSum: 0, qCnt: 0, ds: 0, text: 0, blankQ: 0 };
  STATUSES.forEach(s => a.st[s] = 0);
  QBUCKETS.forEach(b => a.q[b] = 0);
  for (const r of rows) {
    if (a.st[r.status] !== undefined) a.st[r.status]++; else a.other++;
    if (r.lvNum >= 1 && r.lvNum <= 3) a.lv[r.lvNum]++; else if (r.lvNum == null) a.noLv++;
    if (r.q.kind === 'blank') { a.q['0%']++; a.blankQ++; }
    else if (r.q.kind === 'pct' && a.q[r.q.label] !== undefined) a.q[r.q.label]++;
    if (r.q.kind === 'ds') a.ds++;
    if (r.q.kind === 'text') a.text++;
    if (r.q.pct != null) { a.qSum += r.q.pct; a.qCnt++; }
  }
  a.avg = a.qCnt ? a.qSum / a.qCnt : 0;
  a.score = 1 + a.avg;
  a.done = a.st['Hoàn thành'] + a.st['Hoàn thành muộn'];
  a.rate = a.total ? a.done / a.total : 0;
  return a;
}

function groupBy(arr, fn) { const m = new Map(); for (const x of arr) { const k = fn(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; }

const NO_DEPT = '(Chưa xác định phòng ban)';
const sortDepts = arr => [...arr].sort((a, b) => {
  const ia = DEPT_ORDER.indexOf(a), ib = DEPT_ORDER.indexOf(b);
  if (ia >= 0 || ib >= 0) return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
  return viCmp(a, b);
});

/** Phòng ban mặc định cho từng nhân sự trong file dự án:
 *  phòng ban nơi người đó có nhiều công việc nhất trong các file phòng ban (hoà -> theo thứ tự phòng ban). */
function projectDeptDefaults(files) {
  const cnt = new Map();
  for (const f of files.filter(f => f.type === 'task')) {
    const d = norm(f.dept) || 'Khác';
    for (const r of f.rows) {
      if (r.name === UNKNOWN) continue;
      if (!cnt.has(r.name)) cnt.set(r.name, new Map());
      const m = cnt.get(r.name); m.set(d, (m.get(d) || 0) + 1);
    }
  }
  const order = sortDepts([...new Set(files.filter(f => f.type === 'task').map(f => norm(f.dept) || 'Khác'))]);
  const out = new Map();
  for (const f of files.filter(f => f.type === 'project')) for (const r of f.rows) {
    if (!out.has(r.name)) out.set(r.name, { name: r.name, n: 0, dept: NO_DEPT });
    out.get(r.name).n++;
  }
  for (const o of out.values()) {
    const m = cnt.get(o.name);
    if (m) o.dept = [...m.entries()].sort((a, b) => b[1] - a[1] || order.indexOf(a[0]) - order.indexOf(b[0]))[0][0];
  }
  return [...out.values()].sort((a, b) => viCmp(a.name, b.name));
}

/** Dựng model tổng hợp từ các file đã đọc */
function buildModel(files, opts) {
  const taskFiles = files.filter(f => f.type === 'task');
  const projFiles = files.filter(f => f.type === 'project');
  const raw = [];
  for (const f of taskFiles) for (const r of f.rows) raw.push({ ...r, dept: norm(f.dept) || 'Khác', src: f.name, isProj: false });
  // Dữ liệu dự án: gán Phòng ban = phòng ban của nhân sự (mặc định tự nhận diện, người dùng sửa được)
  const assign = new Map(projectDeptDefaults(files).map(o => [o.name, o.dept]));
  const override = opts.projDept || {};
  const proj = [];
  for (const f of projFiles) for (const r of f.rows) {
    const d = r.name === UNKNOWN ? NO_DEPT : (norm(override[r.name]) || assign.get(r.name) || NO_DEPT);
    proj.push({ ...r, dept: d, src: f.name, isProj: true });
  }

  const deptNames = sortDepts(new Set(raw.map(r => r.dept)));   // phòng ban có sheet riêng
  const rawByDept = groupBy(raw, r => r.dept);
  const projByDept = groupBy(proj, r => r.dept);

  const depts = deptNames.map(d => {
    const own = rawByDept.get(d), pj = projByDept.get(d) || [];
    const rows = own.concat(pj);
    const ownByName = groupBy(own, r => r.name), pjByName = groupBy(pj, r => r.name);
    const persons = [...new Set(rows.map(r => r.name))].sort(viCmp).map(n => {
      const o = ownByName.get(n) || [], pr = pjByName.get(n) || [];
      const projCounts = [...groupBy(pr, r => r.project)].map(([p, l]) => ({ project: p, n: l.length }));
      return { name: n, own: o, proj: pr, projCounts, a: agg(o.concat(pr)) };
    });
    return { name: d, rows, persons, a: agg(rows) };
  });

  // Data = dữ liệu phòng ban + dữ liệu dự án đã gán phòng ban (xếp theo phòng ban)
  const unassigned = proj.filter(r => !deptNames.includes(r.dept));
  const data = depts.flatMap(d => d.rows).concat(unassigned);

  const projByName = groupBy(proj, r => r.name);
  const projPersons = [...projByName.keys()].sort(viCmp).map(n => {
    const l = projByName.get(n);
    return { name: n, rows: l, nProjects: new Set(l.map(r => r.project)).size, a: agg(l), lvSum: l.reduce((s, r) => s + (r.lvNum || 0), 0) };
  });
  const projects = [...new Set(proj.map(r => r.project))];

  // Tổng hợp theo người trên toàn bộ Data
  const byPerson = groupBy(data, r => r.name);
  const people = [...byPerson.keys()].sort(viCmp).map(n => {
    const l = byPerson.get(n);
    const dl = sortDepts(new Set(l.map(r => r.dept)));
    return { name: n, rows: l, depts: dl, deptLabel: dl.join(', '), nProjects: new Set(l.filter(r => r.isProj).map(r => r.project)).size, a: agg(l) };
  });
  const named = people.filter(p => p.name !== UNKNOWN);
  const top5 = [...named].sort((x, y) => y.a.total - x.a.total || y.a.done - x.a.done || viCmp(x.name, y.name)).slice(0, 5);
  const minT = Math.max(1, +opts.minTasks || 1);
  const top3 = named.filter(p => p.a.total >= minT && p.a.qCnt > 0)
    .sort((x, y) => y.a.score - x.a.score || y.a.qCnt - x.a.qCnt || y.a.total - x.a.total || viCmp(x.name, y.name)).slice(0, 3);

  // Bảng tổng hợp theo phòng ban (không có "Dự án"); chỉ thêm nhóm chưa xác định phòng ban nếu còn việc chưa gán
  const deptSummary = depts.map(d => ({ name: d.name, rows: d.rows }));
  const unDepts = sortDepts(new Set(unassigned.map(r => r.dept)));
  unDepts.forEach(d => deptSummary.push({ name: d, rows: unassigned.filter(r => r.dept === d) }));
  deptSummary.forEach(s => { s.a = agg(s.rows); s.nPeople = new Set(s.rows.map(r => r.name)).size; });

  const all = agg(data);
  const m = { files, raw, proj, data, depts, deptNames, projPersons, projects, people, top5, top3, deptSummary, unassigned, all, opts };
  m.insights = buildInsights(m);
  m.warnings = buildWarnings(m);
  return m;
}

const pct = (x, d = 1) => (x * 100).toFixed(d).replace('.', ',') + '%';
const pctSigned = x => (x >= 0 ? '+' : '') + pct(x, 2);

function buildInsights(m) {
  const A = m.all, out = [];
  const nPeople = m.people.filter(p => p.name !== UNKNOWN).length;
  if (!A.total) return ['Chưa có dữ liệu công việc.'];
  out.push({ i: '📌', t: `Kỳ ${m.opts.month || ''} ghi nhận <b>${A.total}</b> công việc của <b>${nPeople}</b> nhân sự thuộc <b>${m.deptNames.length}</b> phòng ban` + (m.projects.length ? ` và <b>${m.projects.length}</b> dự án` : '') + `; tỷ lệ hoàn thành (gồm hoàn thành muộn) đạt <b>${pct(A.rate)}</b>.` });
  const late = A.st['Hoàn thành muộn'], over = A.st['Quá hạn'], doing = A.st['Đang làm'];
  out.push({ i: '⏱️', t: `Đúng hạn: <b>${A.st['Hoàn thành']}</b> CV hoàn thành đúng hạn; <b>${late}</b> CV hoàn thành muộn (${pct(A.total ? late / A.total : 0)}), <b>${over}</b> CV quá hạn (${pct(A.total ? over / A.total : 0)}), <b>${doing}</b> CV đang làm chuyển kỳ.` });
  const ds = m.deptSummary.filter(d => d.a.total >= 5).map(d => ({ d, r: (d.a.st['Quá hạn'] + d.a.st['Hoàn thành muộn']) / d.a.total })).sort((x, y) => y.r - x.r);
  if (ds.length && ds[0].r > 0) out.push({ i: '🏢', t: `Phòng ban có tỷ lệ trễ hạn (quá hạn + hoàn thành muộn) cao nhất: <b>${ds[0].d.name}</b> với ${pct(ds[0].r)} (${ds[0].d.a.st['Quá hạn']} quá hạn, ${ds[0].d.a.st['Hoàn thành muộn']} muộn / ${ds[0].d.a.total} CV).` });
  const bestRate = m.deptSummary.filter(d => d.a.total >= 5).sort((x, y) => y.a.rate - x.a.rate)[0];
  if (bestRate) out.push({ i: '✅', t: `Phòng ban có tỷ lệ hoàn thành cao nhất: <b>${bestRate.name}</b> (${pct(bestRate.a.rate)} trên ${bestRate.a.total} CV).` });
  const overP = m.people.filter(p => p.a.st['Quá hạn'] > 0).sort((x, y) => y.a.st['Quá hạn'] - x.a.st['Quá hạn']);
  if (overP.length) out.push({ i: '⚠️', t: `<b>${overP.length}</b> nhân sự có công việc quá hạn; nhiều nhất: ` + overP.slice(0, 3).map(p => `<b>${p.name}</b> (${p.a.st['Quá hạn']})`).join(', ') + '.' });
  const doingP = m.people.filter(p => p.name !== UNKNOWN && p.a.st['Đang làm'] > 0).sort((x, y) => y.a.st['Đang làm'] - x.a.st['Đang làm']);
  if (doingP.length) out.push({ i: '🔄', t: `Tồn việc đang làm nhiều nhất: ` + doingP.slice(0, 3).map(p => `<b>${p.name}</b> (${p.a.st['Đang làm']} CV)`).join(', ') + '.' });
  if (m.top5.length) {
    const s = m.top5.reduce((t, p) => t + p.a.total, 0);
    out.push({ i: '🏋️', t: `Top 5 nhân sự xử lý nhiều việc nhất chiếm <b>${pct(s / A.total)}</b> tổng khối lượng; dẫn đầu là <b>${m.top5[0].name}</b> với ${m.top5[0].a.total} CV.` });
  }
  if (m.top3.length) out.push({ i: '🏆', t: `Điểm đánh giá cao nhất: <b>${m.top3[0].name}</b> đạt ${pct(m.top3[0].a.score)} (chất lượng bổ sung TB ${pctSigned(m.top3[0].a.avg)} trên ${m.top3[0].a.qCnt} CV). Điểm TB toàn công ty: <b>${pct(A.score)}</b>.` });
  const lvKnown = A.lv[1] + A.lv[2] + A.lv[3];
  out.push({ i: '🧩', t: `Mức độ: ${A.lv[1]} CV cơ bản, ${A.lv[2]} trung bình, <b>${A.lv[3]}</b> phức tạp` + (lvKnown ? ` (${pct(A.lv[3] / lvKnown)} số CV có mức độ là phức tạp)` : '') + `; <b>${A.noLv}</b> CV chưa được gán mức độ.` });
  const neg = A.q['-1%'] + A.q['-2%'] + A.q['-3%'], plus = A.q['+1%'] + A.q['+2%'] + A.q['+3%'];
  out.push({ i: '🎯', t: `Chất lượng: <b>${plus}</b> CV được cộng điểm (+1% đến +3%), <b>${neg}</b> CV bị trừ điểm; <b>${A.blankQ}</b> CV (${pct(A.blankQ / A.total)}) chưa được đánh giá → tính 0% theo quy tắc mặc định.` });
  return out;
}

function buildWarnings(m) {
  const w = [];
  const unk = m.data.filter(r => r.name === UNKNOWN);
  if (unk.length) w.push(`<b>${unk.length}</b> công việc không có người nhận việc → gom vào "${UNKNOWN}" (${[...new Set(unk.map(r => r.dept))].join(', ')}).`);
  const txt = m.data.filter(r => r.q.kind === 'text');
  if (txt.length) w.push(`<b>${txt.length}</b> công việc có đánh giá dạng chữ (VD "${txt[0].qRaw}") → không quy đổi %, loại khỏi mẫu số điểm TB, đánh dấu ở cột Ghi chú.`);
  const ds = m.data.filter(r => r.q.kind === 'ds');
  if (ds.length) w.push(`<b>${ds.length}</b> công việc Doanh số → đánh giá riêng theo chỉ tiêu, không tính vào thang %.`);
  const lvOut = m.data.filter(r => r.lvNum != null && (r.lvNum < 1 || r.lvNum > 3));
  if (lvOut.length) w.push(`<b>${lvOut.length}</b> công việc có mức độ ngoài thang 1–3 (${[...new Set(lvOut.map(r => r.lvRaw))].join(', ')}) → không đếm vào cột Mức 1/2/3.`);
  const stOther = m.data.filter(r => !STATUSES.includes(r.status));
  if (stOther.length) w.push(`<b>${stOther.length}</b> công việc có trạng thái khác 4 trạng thái chuẩn (${[...new Set(stOther.map(r => r.status || '(trống)'))].join(', ')}).`);
  const rawDepts = groupBy(m.raw.filter(r => r.name !== UNKNOWN), r => r.name);
  const multi = [...rawDepts].map(([n, l]) => [n, sortDepts(new Set(l.map(r => r.dept)))]).filter(x => x[1].length > 1);
  if (multi.length) w.push(`<b>${multi.length}</b> nhân sự xuất hiện ở nhiều file phòng ban: ` + multi.slice(0, 8).map(([n, d]) => `${n} (${d.join(', ')})`).join('; ') + (multi.length > 8 ? '…' : '') + '. Việc dự án của họ được gán vào phòng ban có nhiều việc nhất.');
  const both = m.depts.flatMap(d => d.persons.filter(p => p.proj.length).map(p => `${p.name} [${d.name}] +${p.proj.length}`));
  if (both.length) w.push(`Việc từ file Dự án đã gộp vào Data và cộng cho nhân sự trùng tên: ${both.join('; ')}.`);
  if (m.unassigned.length) w.push(`<b>${m.unassigned.length}</b> việc dự án chưa xác định được phòng ban (${[...new Set(m.unassigned.map(r => r.name))].join(', ')}) → ghi Phòng ban "${NO_DEPT}". Có thể gán phòng ban ở bước ① trước khi xuất.`);
  const disabled = m.people.filter(p => /vô hiệu hóa|vô hiệu hoá/i.test(p.name));
  if (disabled.length) w.push(`Tài khoản đã vô hiệu hoá vẫn có công việc: ${disabled.map(p => p.name).join(', ')}.`);
  if (!m.proj.length) w.push('Chưa nạp file Dự án → sheet "Tổng hợp Dự án" và "Data_Dự án" sẽ trống.');
  if (!w.length) w.push('Không phát hiện vấn đề dữ liệu.');
  return w;
}
