/* ================= APP: giao diện ================= */
const MAX_FILES = 20;
const ST_CLASS = { 'Hoàn thành': 'ht', 'Hoàn thành muộn': 'htm', 'Đang làm': 'dl', 'Quá hạn': 'qh' };
const ST_COLOR = { 'Hoàn thành': 'var(--ok)', 'Hoàn thành muộn': 'var(--late)', 'Đang làm': 'var(--doing)', 'Quá hạn': 'var(--over)' };
const Q_COLOR = { '+3%': '#1b9e5b', '+2%': '#3aa76d', '+1%': '#8fd19e', '0%': '#9aa5b4', '-1%': '#f7b267', '-2%': '#f08a4b', '-3%': '#e05353' };
const $ = id => document.getElementById(id);
const h = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const nf = n => (n || 0).toLocaleString('vi-VN');

let FILES = [];      // kết quả readRawFile
let MODEL = null;
let dirty = true;
let PROJ_DEPT = {};   // phòng ban người dùng chọn cho nhân sự trong file dự án
let QSEL = '+3%';

/* ---------- nạp file ---------- */
async function addFiles(list) {
  const arr = [...list].filter(f => /\.(xlsx|xls|xlsm)$/i.test(f.name));
  const skipped = list.length - arr.length;
  let msg = [];
  if (skipped) msg.push(`Bỏ qua ${skipped} file không phải Excel.`);
  for (const f of arr) {
    if (FILES.some(x => x.name === f.name && x.file.size === f.size)) { msg.push(`"${f.name}" đã có trong danh sách.`); continue; }
    if (FILES.length >= MAX_FILES) { msg.push(`Đã đạt giới hạn ${MAX_FILES} file — bỏ qua "${f.name}".`); continue; }
    const res = await readRawFile(f);
    FILES.push(res);
  }
  $('fileMsg').textContent = msg.join(' ');
  $('fileMsg').className = 'msg' + (msg.length ? ' err' : '');
  autoSettings();
  renderFiles();
  markDirty();
}
function renderFiles() {
  const tb = $('fileTable').querySelector('tbody');
  if (!FILES.length) { tb.innerHTML = `<tr><td colspan="7" style="color:var(--text-faint);text-align:center;padding:18px">Chưa có file nào.</td></tr>`; }
  else tb.innerHTML = FILES.map((f, i) => {
    const badge = f.type === 'project' ? '<span class="badge b-proj">Dự án</span>' : f.type === 'task' ? '<span class="badge b-task">Công việc phòng ban</span>' : '<span class="badge b-err">Lỗi</span>';
    const dept = f.type === 'task' ? `<input class="in" data-i="${i}" value="${h(f.dept)}" style="width:170px">` : f.type === 'project' ? '<span style="color:var(--text-faint)">→ Data_Dự án</span>' : `<span class="msg err">${h(f.error)}</span>`;
    const per = f.period ? `${fmtDate(f.period.from)} – ${fmtDate(f.period.to)}` : '<span style="color:var(--text-faint)">—</span>';
    return `<tr><td>${i + 1}</td><td class="fname">${h(f.name)}</td><td>${badge}</td><td class="num">${f.rows ? nf(f.rows.length) : '–'}</td><td>${dept}</td><td>${per}</td><td><button class="x" data-del="${i}" title="Bỏ file">✕</button></td></tr>`;
  }).join('');
  tb.querySelectorAll('input[data-i]').forEach(inp => inp.addEventListener('input', () => { FILES[+inp.dataset.i].dept = inp.value; markDirty(); }));
  tb.querySelectorAll('[data-del]').forEach(b => b.addEventListener('click', () => { FILES.splice(+b.dataset.del, 1); renderFiles(); autoSettings(); markDirty(); }));
  $('sFiles').textContent = FILES.length;
  $('btnProcess').disabled = !FILES.some(f => f.type === 'task' || f.type === 'project');
  renderProjAssign();
}
function autoSettings() {
  const pers = FILES.filter(f => f.period).map(f => f.period);
  let from = null, to = null;
  pers.forEach(p => { if (!from || p.from < from) from = p.from; if (!to || p.to > to) to = p.to; });
  if (to) {
    $('optPeriod').value = `${fmtDate(from)} - ${fmtDate(to)}`;
    $('optMonth').value = `${to.getMonth() + 1}.${to.getFullYear()}`;
  } else {
    const mo = FILES.map(f => f.month).find(Boolean);
    if (mo && !$('optMonth').value) $('optMonth').value = `${mo}.${new Date().getFullYear()}`;
  }
  updateOutName();
}
const outFileName = () => `Báo cáo KPI Tháng ${norm($('optMonth').value) || '?'}.xlsx`;
function updateOutName() { $('outName').innerHTML = `File đầu ra: <b style="color:var(--text-strong)">${h(outFileName())}</b>`; }
function markDirty() { dirty = true; $('btnDownload').disabled = $('btnDownload2').disabled = true; if (MODEL) $('procMsg').textContent = 'Dữ liệu/thiết lập đã thay đổi — bấm "Xử lý & phân tích" lại.'; }

/* ---------- xử lý ---------- */
function process() {
  const opts = { month: norm($('optMonth').value), period: norm($('optPeriod').value), minTasks: Math.max(1, +$('optMin').value || 1), projDept: { ...PROJ_DEPT } };
  MODEL = buildModel(FILES, opts);
  dirty = false;
  const m = MODEL;
  $('sTasks').textContent = nf(m.data.length);
  $('sPeople').textContent = nf(m.people.filter(p => p.name !== UNKNOWN).length);
  $('sDepts').textContent = m.deptNames.length;
  $('sProjects').textContent = m.projects.length;
  $('procMsg').className = 'msg ok';
  $('procMsg').textContent = `✔ Đã tổng hợp ${nf(m.data.length)} công việc (${nf(m.raw.length)} từ file phòng ban + ${nf(m.proj.length)} từ file dự án). Xem các tab phía trên hoặc tải file Excel.`;
  document.querySelectorAll('#tabs .chip').forEach(c => c.classList.remove('disabled'));
  $('btnDownload').disabled = $('btnDownload2').disabled = false;
  renderOverview(); renderPersonSelect(); renderQuality(); renderTop(); renderCheck();
}

async function download() {
  if (!MODEL || dirty) process();
  const btns = [$('btnDownload'), $('btnDownload2')];
  const old = btns.map(b => b.innerHTML);
  btns.forEach(b => { b.disabled = true; b.innerHTML = '<span class="spin"></span> Đang tạo file…'; });
  try {
    const { blob } = await buildWorkbook(MODEL);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = outFileName();
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
    $('procMsg').className = 'msg ok'; $('procMsg').textContent = `✔ Đã tạo "${outFileName()}".`;
  } catch (e) {
    console.error(e);
    $('procMsg').className = 'msg err'; $('procMsg').textContent = 'Lỗi khi tạo file Excel: ' + e.message;
  } finally { btns.forEach((b, i) => { b.disabled = false; b.innerHTML = old[i]; }); }
}

/* ---------- vẽ ---------- */
function kpi(label, value, sub, color) { return `<div class="kpi" style="--c:${color || 'var(--accent)'}"><div class="l">${label}</div><div class="v">${value}</div>${sub ? `<div class="s">${sub}</div>` : ''}</div>`; }
function donut(items, total) {
  const R = 60, Cc = 2 * Math.PI * R; let off = 0;
  const segs = items.filter(x => x.v > 0).map(x => { const len = total ? x.v / total * Cc : 0; const s = `<circle r="${R}" cx="80" cy="80" fill="none" stroke="${x.c}" stroke-width="26" stroke-dasharray="${len} ${Cc - len}" stroke-dashoffset="${-off}" transform="rotate(-90 80 80)"><title>${h(x.l)}: ${x.v}</title></circle>`; off += len; return s; }).join('');
  return `<div style="display:flex;gap:18px;align-items:center;flex-wrap:wrap"><svg width="160" height="160" viewBox="0 0 160 160"><circle r="${R}" cx="80" cy="80" fill="none" style="stroke:var(--track)" stroke-width="26"/>${segs}<text x="80" y="78" text-anchor="middle" style="fill:var(--text-strong)" font-size="24" font-weight="800">${nf(total)}</text><text x="80" y="98" text-anchor="middle" style="fill:var(--text-faint)" font-size="11">công việc</text></svg>` +
    `<div style="flex:1;min-width:170px">${items.map(x => `<div class="hbar" style="grid-template-columns:130px 1fr 44px"><span><i style="display:inline-block;width:10px;height:10px;border-radius:3px;background:${x.c};margin-right:6px"></i>${h(x.l)}</span><div class="track"><div class="fill" style="width:${total ? x.v / total * 100 : 0}%;background:${x.c}"></div></div><span class="n">${nf(x.v)}</span></div>`).join('')}</div></div>`;
}
function stackBar(a) {
  const parts = STATUSES.map(s => ({ s, v: a.st[s] })).concat([{ s: 'Khác', v: a.other }]);
  return `<div class="stack" title="${parts.map(p => p.s + ': ' + p.v).join(' · ')}">${parts.filter(p => p.v).map(p => `<span style="width:${p.v / a.total * 100}%;background:${ST_COLOR[p.s] || 'var(--other)'}"></span>`).join('')}</div>`;
}
const legendHtml = () => `<div class="legend">${STATUSES.map(s => `<span><i style="background:${ST_COLOR[s]}"></i>${s}</span>`).join('')}</div>`;
const stHtml = s => `<span class="st st-${ST_CLASS[s] || 'kh'}">${h(s || '—')}</span>`;

function renderOverview() {
  const m = MODEL, A = m.all;
  $('ovHint').textContent = `Kỳ báo cáo: ${m.opts.period || '—'} · Tháng ${m.opts.month || '—'}`;
  $('ovKpis').innerHTML = [
    kpi('Tổng công việc', nf(A.total), `${m.files.length} file nguồn`, 'var(--accent)'),
    kpi('Hoàn thành', nf(A.st['Hoàn thành']), pct(A.total ? A.st['Hoàn thành'] / A.total : 0), 'var(--ok)'),
    kpi('Hoàn thành muộn', nf(A.st['Hoàn thành muộn']), pct(A.total ? A.st['Hoàn thành muộn'] / A.total : 0), 'var(--late)'),
    kpi('Đang làm', nf(A.st['Đang làm']), pct(A.total ? A.st['Đang làm'] / A.total : 0), 'var(--doing)'),
    kpi('Quá hạn', nf(A.st['Quá hạn']), pct(A.total ? A.st['Quá hạn'] / A.total : 0), 'var(--over)'),
    kpi('Tỷ lệ hoàn thành', pct(A.rate), 'gồm hoàn thành muộn', 'var(--accent-2)'),
    kpi('Điểm đánh giá TB', pct(A.score), 'CL bổ sung ' + pctSigned(A.avg), 'var(--accent-3)')
  ].join('');
  $('ovDonut').innerHTML = donut(STATUSES.map(s => ({ l: s, v: A.st[s], c: ST_COLOR[s] })).concat(A.other ? [{ l: 'Khác', v: A.other, c: 'var(--other)' }] : []), A.total);
  $('ovDeptBars').innerHTML = m.deptSummary.map(d => `<div class="hbar" style="grid-template-columns:minmax(110px,150px) 1fr 44px"><span>${h(d.name)}</span>${stackBar(d.a)}<span class="n">${nf(d.a.total)}</span></div>`).join('') + legendHtml();
  const rows = m.deptSummary.map(d => `<tr><td class="fname">${h(d.name)}</td><td class="num">${d.nPeople}</td><td class="num">${nf(d.a.total)}</td>${STATUSES.map(s => `<td class="num">${nf(d.a.st[s])}</td>`).join('')}<td class="num">${pct(d.a.rate)}</td><td class="num">${pct(d.a.score)}</td><td style="min-width:140px">${stackBar(d.a)}</td></tr>`).join('');
  $('deptTable').innerHTML = `<thead><tr><th>Phòng ban</th><th class="num">Nhân sự</th><th class="num">Tổng CV</th>${STATUSES.map(s => `<th class="num">${s}</th>`).join('')}<th class="num">Tỷ lệ HT</th><th class="num">Điểm TB</th><th>Cơ cấu</th></tr></thead><tbody>${rows}</tbody>` +
    `<tfoot><tr><td>Tổng</td><td class="num">${m.people.filter(p => p.name !== UNKNOWN).length}</td><td class="num">${nf(A.total)}</td>${STATUSES.map(s => `<td class="num">${nf(A.st[s])}</td>`).join('')}<td class="num">${pct(A.rate)}</td><td class="num">${pct(A.score)}</td><td></td></tr></tfoot>`;
}

function renderPersonSelect(filter = '') {
  const m = MODEL, f = lc(filter);
  const sel = $('personSel'), cur = sel.value;
  const list = m.people.filter(p => !f || lc(p.name).includes(f) || lc(p.deptLabel).includes(f));
  sel.innerHTML = `<option value="${h(ALL)}">${h(ALL)} — toàn công ty</option>` + list.map(p => `<option value="${h(p.name)}">${h(p.name)} · ${h(p.deptLabel)} (${p.a.total})</option>`).join('');
  if ([...sel.options].some(o => o.value === cur)) sel.value = cur; else if (f && list.length) sel.value = list[0].name;
  renderPerson(sel.value);
}
function renderPerson(name) {
  const m = MODEL;
  const p = name === ALL ? { name: ALL, rows: m.data, deptLabel: 'Toàn công ty', nProjects: m.projects.length, a: m.all } : m.people.find(x => x.name === name);
  if (!p) return;
  const a = p.a;
  $('personInfo').innerHTML = `${h(p.deptLabel)} · ${p.nProjects} dự án`;
  $('pKpis').innerHTML = [
    kpi('Tổng công việc', nf(a.total), `Mức 1/2/3: ${a.lv[1]}/${a.lv[2]}/${a.lv[3]}`, 'var(--accent)'),
    kpi('Hoàn thành', nf(a.st['Hoàn thành']), '', 'var(--ok)'),
    kpi('Hoàn thành muộn', nf(a.st['Hoàn thành muộn']), '', 'var(--late)'),
    kpi('Đang làm', nf(a.st['Đang làm']), '', 'var(--doing)'),
    kpi('Quá hạn', nf(a.st['Quá hạn']), '', 'var(--over)'),
    kpi('Chất lượng bổ sung TB', pctSigned(a.avg), `${a.qCnt} CV tính điểm`, 'var(--accent-4)'),
    kpi('Điểm cuối cùng', pct(a.score), 'Tỷ lệ HT ' + pct(a.rate), 'var(--accent-3)')
  ].join('');
  $('pStatus').innerHTML = donut(STATUSES.map(s => ({ l: s, v: a.st[s], c: ST_COLOR[s] })).concat(a.other ? [{ l: 'Khác', v: a.other, c: 'var(--other)' }] : []), a.total);
  const qmax = Math.max(1, ...QBUCKETS.map(b => a.q[b]));
  $('pQuality').innerHTML = QBUCKETS.map(b => `<div class="hbar" style="grid-template-columns:90px 1fr 44px"><span>${b === '0%' ? '0% (gồm trống)' : b}</span><div class="track"><div class="fill" style="width:${a.q[b] / qmax * 100}%;background:${Q_COLOR[b]}"></div></div><span class="n">${a.q[b]}</span></div>`).join('') +
    (a.text || a.ds ? `<div class="legend">${a.text ? `${a.text} CV đánh giá dạng chữ · ` : ''}${a.ds ? `${a.ds} CV Doanh số` : ''} (không quy đổi %)</div>` : '');
  const rows = name === ALL ? [] : p.rows;
  $('pTaskTitle').textContent = name === ALL ? 'Chi tiết công việc — chọn một nhân sự để xem danh sách' : `Chi tiết ${rows.length} công việc của ${name}`;
  $('pTasks').innerHTML = rows.length ? `<thead><tr><th>#</th><th>Nguồn</th><th>Tên công việc</th><th>Trạng thái</th><th>Mức độ</th><th>Đánh giá CL</th><th>Ghi chú</th></tr></thead><tbody>` +
    rows.map((r, i) => `<tr><td>${i + 1}</td><td>${h(r.dept + (r.isProj ? ' · Dự án: ' + r.project : ''))}</td><td>${h(r.task)}</td><td>${stHtml(r.status)}</td><td class="num">${h(r.lvRaw)}</td><td class="num">${h(r.qRaw)}</td><td style="color:var(--text-faint);font-size:12.5px">${h(r.note)}</td></tr>`).join('') + '</tbody>' : '';
}

function renderTop() {
  const m = MODEL;
  const max5 = Math.max(1, ...m.top5.map(p => p.a.total));
  $('top5').innerHTML = m.top5.map((p, i) => `<div class="hbar" style="grid-template-columns:30px minmax(120px,190px) 1fr 44px"><span class="rank ${i < 3 ? 'r' + (i + 1) : ''}">${i + 1}</span><span><b style="color:var(--text-strong)">${h(p.name)}</b><br><small style="color:var(--text-faint)">${h(p.deptLabel)}</small></span><div class="track"><div class="fill" style="width:${p.a.total / max5 * 100}%"></div></div><span class="n">${p.a.total}</span></div>`).join('') || '<p class="hint">Không có dữ liệu.</p>';
  $('top3Hint').textContent = `Điểm = 100% + chất lượng bổ sung TB trên toàn bộ CV (phòng ban + dự án). Chỉ xét nhân sự có ≥ ${m.opts.minTasks} CV; hoà điểm thì ưu tiên nhiều CV được tính điểm hơn.`;
  $('top3').innerHTML = m.top3.length ? `<table class="t"><thead><tr><th></th><th>Nhân sự</th><th class="num">Tổng CV</th><th class="num">CV tính điểm</th><th class="num">CL TB</th><th class="num">Điểm</th></tr></thead><tbody>` +
    m.top3.map((p, i) => `<tr><td><span class="rank r${i + 1}">${i + 1}</span></td><td><b style="color:var(--text-strong)">${h(p.name)}</b><br><small style="color:var(--text-faint)">${h(p.deptLabel)}</small></td><td class="num">${p.a.total}</td><td class="num">${p.a.qCnt}</td><td class="num">${pctSigned(p.a.avg)}</td><td class="num" style="font-size:16px;font-weight:800;color:var(--accent-3)">${pct(p.a.score)}</td></tr>`).join('') + '</tbody></table>'
    : '<p class="hint">Chưa có nhân sự đủ điều kiện.</p>';
  $('insights').innerHTML = m.insights.map(x => `<li data-i="${x.i}">${x.t}</li>`).join('');
}

function renderCheck() {
  const m = MODEL;
  $('warnList').innerHTML = m.warnings.map(w => `<li>${w}</li>`).join('');
  const list = [['Mastersheet', 'Dashboard: KPI tổng quan, bảng & biểu đồ theo phòng ban, ô chọn nhân sự (biểu đồ + chi tiết tự cập nhật), Top 5, Top 3, insight'],
    ['Nguyên tắc đánh giá', 'Quy tắc tính điểm (giống file mẫu)']]
    .concat(m.depts.map(d => [d.name, `${d.persons.length} nhân sự · ${d.rows.length} CV` + (d.persons.some(p => p.proj.length) ? ' · có cộng việc Dự án cho người trùng tên' : '')]))
    .concat(m.proj.length ? [['Tổng hợp Dự án', `${m.projPersons.length} nhân sự · ${m.projects.length} dự án (form "Kỹ thuật (2)")`]] : [])
    .concat([['Data', `${m.data.length} dòng (phòng ban + dự án)`], ['Data_Dự án', `${m.proj.length} dòng, đã tách Mức độ / Đánh giá chất lượng từ "Trường dữ liệu tùy chỉnh"`]]);
  $('sheetTable').innerHTML = `<thead><tr><th>#</th><th>Sheet</th><th>Nội dung</th></tr></thead><tbody>${list.map((x, i) => `<tr><td>${i + 1}</td><td class="fname">${h(x[0])}</td><td>${h(x[1])}</td></tr>`).join('')}</tbody>`;
}


/* ---------- gán phòng ban cho nhân sự trong file dự án ---------- */
function renderProjAssign() {
  const list = projectDeptDefaults(FILES);
  $('projAssignBox').classList.toggle('hidden', !list.length);
  if (!list.length) return;
  const depts = sortDepts(new Set(FILES.filter(f => f.type === 'task').map(f => norm(f.dept) || 'Khác'))).concat([NO_DEPT]);
  Object.keys(PROJ_DEPT).forEach(k => { if (!depts.includes(PROJ_DEPT[k])) delete PROJ_DEPT[k]; });
  $('projAssign').innerHTML = `<thead><tr><th>Nhân sự</th><th class="num">Số việc dự án</th><th>Phòng ban để gộp vào Data</th></tr></thead><tbody>` +
    list.map((o, i) => {
      const cur = o.name === UNKNOWN ? NO_DEPT : (PROJ_DEPT[o.name] || o.dept);
      const opts = depts.map(d => `<option ${d === cur ? 'selected' : ''} value="${h(d)}">${h(d)}${d === o.dept && o.dept !== NO_DEPT ? ' (tự nhận diện)' : ''}</option>`).join('');
      return `<tr><td class="fname">${h(o.name)}</td><td class="num">${o.n}</td><td><select class="in" data-pa="${i}" ${o.name === UNKNOWN ? 'disabled' : ''} style="min-width:230px">${opts}</select></td></tr>`;
    }).join('') + '</tbody>';
  $('projAssign').querySelectorAll('select[data-pa]').forEach(sel => sel.addEventListener('change', () => {
    const o = list[+sel.dataset.pa];
    if (sel.value === o.dept) delete PROJ_DEPT[o.name]; else PROJ_DEPT[o.name] = sel.value;
    markDirty();
  }));
}

/* ---------- lọc theo đánh giá chất lượng ---------- */
const QOPTS = ['+3%', '+2%', '+1%', '0%', '(Trống)', '-1%', '-2%', '-3%', 'Dạng chữ / Doanh số'];
const qMatch = (r, opt) => opt === '(Trống)' ? r.q.kind === 'blank'
  : opt === 'Dạng chữ / Doanh số' ? (r.q.kind === 'text' || r.q.kind === 'ds')
  : r.q.kind === 'pct' && r.qRaw === opt;
function renderQuality() {
  const m = MODEL;
  $('qChips').innerHTML = QOPTS.map(o => `<div class="chip ${o === QSEL ? 'active' : ''}" data-q="${h(o)}">${h(o)} <span style="opacity:.75">(${m.data.filter(r => qMatch(r, o)).length})</span></div>`).join('');
  $('qChips').querySelectorAll('[data-q]').forEach(c => c.addEventListener('click', () => { QSEL = c.dataset.q; renderQuality(); }));
  const dsel = $('qDept'), cur = dsel.value;
  dsel.innerHTML = `<option value="">Tất cả phòng ban</option>` + m.deptSummary.map(d => `<option value="${h(d.name)}">${h(d.name)}</option>`).join('');
  if ([...dsel.options].some(o => o.value === cur)) dsel.value = cur;
  renderQualityTable();
}
function renderQualityTable() {
  const m = MODEL, dept = $('qDept').value, f = lc($('qSearch').value);
  const rows = m.data.filter(r => qMatch(r, QSEL) && (!dept || r.dept === dept) && (!f || lc(r.name + ' ' + r.task).includes(f)));
  const people = new Set(rows.map(r => r.name)).size;
  $('qInfo').textContent = `${rows.length} công việc · ${people} nhân viên`;
  $('qTable').innerHTML = rows.length ? `<thead><tr><th>#</th><th>Nhân viên</th><th>Phòng ban</th><th>Tên công việc</th><th>Trạng thái</th><th class="num">Mức độ</th><th>Đánh giá</th><th>Ghi chú</th></tr></thead><tbody>` +
    rows.map((r, i) => `<tr><td>${i + 1}</td><td class="fname">${h(r.name)}</td><td>${h(r.dept)}</td><td>${h(r.task)}</td><td>${stHtml(r.status)}</td><td class="num">${h(r.lvRaw)}</td><td>${h(r.qRaw || '(trống)')}</td>` +
      `<td style="color:var(--text-faint);font-size:12.5px">${h([r.isProj ? 'Dự án: ' + r.project : '', r.note].filter(Boolean).join(' · '))}</td></tr>`).join('') + '</tbody>'
    : '<tbody><tr><td style="color:var(--text-faint);padding:18px">Không có công việc phù hợp.</td></tr></tbody>';
}

/* ---------- sự kiện ---------- */
const drop = $('drop'), fi = $('fileInput');
drop.addEventListener('click', () => fi.click());
fi.addEventListener('change', () => { addFiles(fi.files); fi.value = ''; });
['dragenter', 'dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', ev => addFiles(ev.dataTransfer.files));
$('btnProcess').addEventListener('click', process);
$('btnDownload').addEventListener('click', download);
$('btnDownload2').addEventListener('click', download);
['optMonth', 'optPeriod', 'optMin'].forEach(id => $(id).addEventListener('input', () => { updateOutName(); markDirty(); }));
$('personSel').addEventListener('change', e => renderPerson(e.target.value));
$('personSearch').addEventListener('input', e => renderPersonSelect(e.target.value));
$('qDept').addEventListener('change', renderQualityTable);
$('qSearch').addEventListener('input', renderQualityTable);
document.querySelectorAll('#tabs .chip').forEach(ch => ch.addEventListener('click', () => {
  document.querySelectorAll('#tabs .chip').forEach(c => c.classList.remove('active'));
  ch.classList.add('active');
  document.querySelectorAll('.tabview').forEach(v => v.classList.toggle('show', v.dataset.view === ch.dataset.tab));
  window.scrollTo({ top: document.querySelector('.controls').offsetTop, behavior: 'smooth' });
}));
renderFiles(); updateOutName();
