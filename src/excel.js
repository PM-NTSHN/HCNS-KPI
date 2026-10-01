/* ================= EXCEL: dựng file báo cáo theo form mẫu ================= */
const colL = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const qs = s => '"' + String(s).replace(/"/g, '""') + '"';
const C = { blue: 'FF305496', hdr: 'FF4F81BD', orange: 'FFC55A11', orange2: 'FFED7D31', total: 'FFFCE4D6', grey: 'FFB7B7B7', black: 'FF000000',
  ok: 'FF3AA76D', late: 'FFF2A93B', doing: 'FF4F81BD', over: 'FFE05353', navy: 'FF1F3864', pale: 'FFF3F6FB', yellow: 'FFFFF2CC' };
const thin = c => ({ style: 'thin', color: { argb: c } });
const box = c => ({ top: thin(c), left: thin(c), bottom: thin(c), right: thin(c) });
const solidFill = c => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: c } });
const center = { horizontal: 'center', vertical: 'middle', wrapText: true };
const left = { horizontal: 'left', vertical: 'middle', wrapText: true };

function styleRange(ws, range, st) {
  const [a, b] = range.split(':');
  const s = ws.getCell(a), e = ws.getCell(b || a);
  for (let r = s.row; r <= e.row; r++) for (let c = s.col; c <= e.col; c++) {
    const cell = ws.getCell(r, c);
    if (st.font) cell.font = st.font;
    if (st.fill) cell.fill = st.fill;
    if (st.border) cell.border = st.border;
    if (st.alignment) cell.alignment = st.alignment;
    if (st.numFmt) cell.numFmt = st.numFmt;
  }
}
function put(ws, addr, value, st) {
  const [a] = addr.split(':');
  if (addr.includes(':')) ws.mergeCells(addr);
  ws.getCell(a).value = value;
  if (st) styleRange(ws, addr, st);
  return ws.getCell(a);
}
const F = (formula, result) => ({ formula, result: result === undefined || result === null ? '' : result });

function safeSheetName(name, used) {
  let s = norm(name).replace(/[\\\/\?\*\[\]:]/g, ' ').slice(0, 31).trim() || 'Sheet';
  let base = s, i = 2;
  while (used.has(s.toLowerCase())) { const suf = ` (${i++})`; s = base.slice(0, 31 - suf.length) + suf; }
  used.add(s.toLowerCase());
  return s;
}

/* ---------- Sheet: Nguyên tắc đánh giá ---------- */
const RULES = [
  '- Gộp toàn bộ công việc của một nhân viên trên tất cả các dự án trong file Dự án.xlsx ',
  '- "Số dự án tham gia" là số dự án khác nhau mà nhân viên có công việc được giao.',
  '- "Mức độ công việc": lấy theo trường "Mức độ công việc" trên hệ thống gốc 1 = cơ bản, 2 = trung bình, 3 = phức tạp',
  '- "Phân bố số lượng công việc theo thang điểm đánh giá chất lượng": đếm theo cột "Đánh giá chất lượng" gốc (+3%/+2%/+1%/0%/-1%/-2%/-3%). Ô để trống được tính vào cột 0% theo đúng quy tắc mặc định.',
  '- "Tỷ lệ hoàn thành (mặc định)": mặc định 100% cho mọi công việc theo yêu cầu.',
  '- "Chất lượng bổ sung TB (%)" = trung bình cộng các mức % ở trên (ô trống tính 0%). "Điểm đánh giá cuối cùng" = 100% + Chất lượng bổ sung TB.',
  '- Cộng thêm số liệu công việc từ file Dự án cho các nhân sự trùng tên.'
];
function rulesSheet(wb, name) {
  const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', xSplit: 4, ySplit: 1 }] });
  const widths = { A: 154, B: 17.2, C: 9, D: 4.9, E: 8, H: 5.3, O: 12, R: 50 };
  Object.entries(widths).forEach(([k, w]) => ws.getColumn(k).width = w);
  put(ws, 'A1:D2', 'NGUYÊN TẮC ĐÁNH GIÁ', { font: { name: 'Times New Roman', size: 14, bold: true }, alignment: center });
  ws.getRow(1).height = 17.25;
  const extra = [
    '- Công việc có "Đánh giá chất lượng" ghi bằng chữ (VD "Hoàn thành tốt") hoặc công việc Doanh số KHÔNG được quy đổi vào thang %, được đánh dấu ở cột "Ghi chú".',
    '- "(Chưa xác định)": các công việc không có tên người nhận việc trong file gốc.',
    '- Sheet "Data" = dữ liệu các file phòng ban + dữ liệu file Dự án (tách từ "Trường dữ liệu tùy chỉnh", Phòng ban = phòng ban của nhân sự); sheet "Mastersheet" là dashboard tổng quan, chọn tên nhân sự để xem chi tiết.'
  ];
  RULES.concat(extra).forEach((t, i) => {
    const r = 3 + i;
    put(ws, `A${r}:R${r}`, t, { font: { name: 'Times New Roman', size: 12 }, alignment: { vertical: 'middle', wrapText: true } });
    ws.getRow(r).height = 27;
  });
  return ws;
}

/* ---------- Sheet phòng ban (form Admin/MKT/...) ---------- */
function deptSheet(wb, sheetName, dept, ctx) {
  const { DR, period } = ctx;
  const ws = wb.addWorksheet(sheetName, { views: [{ state: 'frozen', xSplit: 3, ySplit: 5 }] });
  [5, 22, 11, 8.8, 9.8, 7.9, 7, 7, 7, 7, 7, 7, 7, 11.9, 13.8, 13.4, 48.3].forEach((w, i) => ws.getColumn(i + 1).width = w);
  const FN = 'Times New Roman';
  put(ws, 'A1:Q1', `BÁO CÁO ĐÁNH GIÁ CÔNG VIỆC THEO NHÂN SỰ - PHÒNG BAN: ${dept.name.toUpperCase()}`, { font: { name: FN, size: 14, bold: true, color: { argb: C.blue } }, alignment: center });
  put(ws, 'A2:Q2', `Kỳ báo cáo: ${period}`, { font: { name: FN, size: 9, italic: true, color: { argb: 'FF7F7F7F' } }, alignment: center });
  ws.getRow(1).height = 20; ws.getRow(2).height = 15; ws.getRow(4).height = 30.6; ws.getRow(5).height = 40.8;
  const H = { font: { name: FN, size: 10, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(C.hdr), border: box(C.black), alignment: center };
  put(ws, 'A4:A5', 'STT', H); put(ws, 'B4:B5', 'Người nhận việc', H); put(ws, 'C4:C5', 'SL CV\nthực hiện trong tháng', H);
  put(ws, 'D4:F4', 'MỨC ĐỘ CÔNG VIỆC', H); put(ws, 'G4:M4', 'PHÂN BỐ SỐ LƯỢNG CÔNG VIỆC THEO THANG ĐIỂM ĐÁNH GIÁ CHẤT LƯỢNG', H);
  put(ws, 'N4:P4', 'ĐÁNH GIÁ CHẤT LƯỢNG CÔNG VIỆC', H); put(ws, 'Q4:Q5', 'Ghi chú', H);
  ['Mức 1\n- Cơ bản', 'Mức 2- Trung bình', 'Mức 3- Phức tạp'].concat(QBUCKETS, ['Tỷ lệ hoàn thành\n(mặc định)', 'Chất lượng bổ sung\nTB (%)', 'Điểm đánh giá\ncuối cùng'])
    .forEach((t, i) => put(ws, colL(4 + i) + '5', t, H));

  const B = { font: { name: FN, size: 10 }, border: box(C.black), alignment: center };
  const d = qs(dept.name);
  let r = 6;
  dept.persons.forEach((p, i) => {
    const n = qs(p.name);
    const base = `${DR('A')},${d},${DR('B')},${n}`;
    // Việc từ file Dự án đã nằm trong sheet Data với Phòng ban của nhân sự -> chỉ cần đếm trên Data
    const cf = (col, crit) => `COUNTIFS(${base},${DR(col)},${crit})`;
    const a = p.a;
    ws.getRow(r).height = 25.2;
    put(ws, `A${r}`, i + 1, B);
    put(ws, `B${r}`, p.name, { ...B, alignment: left });
    put(ws, `C${r}`, F(`COUNTIFS(${base})`, a.total), B);
    [1, 2, 3].forEach((lv, k) => put(ws, colL(4 + k) + r, F(cf('F', lv), a.lv[lv]), B));
    QBUCKETS.forEach((b, k) => {
      let f = cf('G', qs(b));
      if (b === '0%') f += `+COUNTIFS(${base},${DR('G')},"")`;
      put(ws, colL(7 + k) + r, F(f, a.q[b]), B);
    });
    put(ws, `N${r}`, 1, { ...B, numFmt: '0%' });
    put(ws, `O${r}`, F(`IFERROR(AVERAGEIFS(${DR('H')},${base}),0)`, a.avg), { ...B, numFmt: '0.0%' });
    put(ws, `P${r}`, F(`N${r}+O${r}`, 1 + a.avg), { ...B, numFmt: '0.0%' });
    // Ghi chú
    const dsCnt = `COUNTIFS(${base},${DR('I')},${qs(NOTE_DS)})`;
    const dsTxt = `IF(${dsCnt}>0,"Có "&${dsCnt}&" việc Doanh số (đánh giá riêng theo chỉ tiêu, không nằm trong thang điểm trên)","")`;
    const staticNotes = [];
    if (a.text) staticNotes.push(`Có ${a.text} việc đánh giá dạng chữ (không quy đổi %)`);
    if (p.proj.length) staticNotes.push(`Đã cộng thêm ${p.proj.length} việc từ file Dự án - Dự án: ` + p.projCounts.map(x => `${x.project} (${x.n} việc)`).join('; '));
    const dsRes = a.ds ? `Có ${a.ds} việc Doanh số (đánh giá riêng theo chỉ tiêu, không nằm trong thang điểm trên)` : '';
    const qF = staticNotes.length ? `TRIM(${dsTxt}&" "&${qs(staticNotes.join('. '))})` : dsTxt;
    put(ws, `Q${r}`, F(qF, [dsRes, staticNotes.join('. ')].filter(Boolean).join(' ')), { ...B, alignment: left });
    r++;
  });
  // dòng Tổng
  const first = 6, last = r - 1;
  const T = { font: { name: FN, size: 10, bold: true }, fill: solidFill(C.total), border: box(C.grey), alignment: center };
  styleRange(ws, `A${r}:Q${r}`, T);
  put(ws, `B${r}`, 'Tổng', { ...T, alignment: left });
  const tot = dept.persons.reduce((acc, p) => {
    acc.total += p.a.total; [1, 2, 3].forEach(l => acc.lv[l] += p.a.lv[l]); QBUCKETS.forEach(b => acc.q[b] += p.a.q[b]); acc.avg.push(p.a.avg); return acc;
  }, { total: 0, lv: { 1: 0, 2: 0, 3: 0 }, q: Object.fromEntries(QBUCKETS.map(b => [b, 0])), avg: [] });
  const avgAvg = tot.avg.length ? tot.avg.reduce((s, x) => s + x, 0) / tot.avg.length : 0;
  if (last >= first) {
    put(ws, `C${r}`, F(`SUM(C${first}:C${last})`, tot.total), T);
    [1, 2, 3].forEach((l, k) => put(ws, colL(4 + k) + r, F(`SUM(${colL(4 + k)}${first}:${colL(4 + k)}${last})`, tot.lv[l]), T));
    QBUCKETS.forEach((b, k) => put(ws, colL(7 + k) + r, F(`SUM(${colL(7 + k)}${first}:${colL(7 + k)}${last})`, tot.q[b]), T));
    put(ws, `N${r}`, 1, { ...T, numFmt: '0%' });
    put(ws, `O${r}`, F(`IFERROR(AVERAGE(O${first}:O${last}),0)`, avgAvg), { ...T, numFmt: '0.0%' });
    put(ws, `P${r}`, F(`N${r}+O${r}`, 1 + avgAvg), { ...T, numFmt: '0.0%' });
  }
  ws.getRow(r).height = 22;
  ws.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };
  return ws;
}

/* ---------- Sheet tổng hợp Dự án (form "Kỹ thuật (2)") ---------- */
function projectSheet(wb, sheetName, m, ctx) {
  const { PR, period } = ctx;
  const ws = wb.addWorksheet(sheetName, { views: [{ state: 'frozen', xSplit: 4, ySplit: 5 }], properties: { tabColor: { argb: C.orange } } });
  [5, 22, 9, 11, 8, 8, 8, 9, 7, 7, 7, 7, 7, 7, 7, 12, 12, 12, 50].forEach((w, i) => ws.getColumn(i + 1).width = w);
  const FN = 'Arial';
  put(ws, 'A1:S1', 'BÁO CÁO ĐÁNH GIÁ CÔNG VIỆC THEO NHÂN VIÊN - TẤT CẢ CÁC DỰ ÁN', { font: { name: FN, size: 14, bold: true, color: { argb: C.orange } }, alignment: center });
  put(ws, 'A2:S2', `Kỳ báo cáo: ${period} · Nguồn dữ liệu: sheet "Data_Dự án" (đã chuẩn hoá từ file Dự án) - gộp số liệu của một nhân viên trên tất cả các dự án`,
    { font: { name: FN, size: 9, italic: true, color: { argb: 'FF7F7F7F' } }, alignment: center });
  ws.getRow(1).height = 20; ws.getRow(2).height = 15; ws.getRow(4).height = 34.8; ws.getRow(5).height = 44;
  const H = { font: { name: FN, size: 10, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(C.orange), border: box(C.grey), alignment: center };
  const H2 = { ...H, fill: solidFill(C.orange2) };
  put(ws, 'A4:A5', 'STT', H); put(ws, 'B4:B5', 'Người nhận việc', H); put(ws, 'C4:C5', 'Số dự án\ntham gia', H);
  put(ws, 'D4:D5', 'Số lượng công việc\nthực hiện', H); put(ws, 'E4:H4', 'MỨC ĐỘ CÔNG VIỆC', H);
  put(ws, 'I4:O4', 'PHÂN BỐ SỐ LƯỢNG CÔNG VIỆC THEO THANG ĐIỂM ĐÁNH GIÁ CHẤT LƯỢNG', H);
  put(ws, 'P4:R4', 'ĐÁNH GIÁ CHẤT LƯỢNG CÔNG VIỆC', H); put(ws, 'S4:S5', 'Ghi chú', H);
  ['Mức 1\n(số việc)', 'Mức 2\n(số việc)', 'Mức 3\n(số việc)', 'Tổng điểm\nmức độ'].concat(QBUCKETS, ['Tỷ lệ hoàn thành\n(mặc định)', 'Chất lượng bổ sung\nTB (%)', 'Điểm đánh giá\ncuối cùng'])
    .forEach((t, i) => put(ws, colL(5 + i) + '5', t, H2));
  const B = { font: { name: FN, size: 10 }, border: box(C.grey), alignment: center };
  let r = 6;
  m.projPersons.forEach((p, i) => {
    const n = qs(p.name), pb = `${PR('B')},${n}`;
    ws.getRow(r).height = p.rows.some(x => x.note) ? 30 : 18;
    put(ws, `A${r}`, i + 1, B);
    put(ws, `B${r}`, p.name, { ...B, alignment: left });
    put(ws, `C${r}`, F(`SUMPRODUCT((${PR('B')}=${n})/COUNTIFS(${PR('B')},${PR('B')},${PR('A')},${PR('A')}))`, p.nProjects), B);
    put(ws, `D${r}`, F(`COUNTIFS(${pb})`, p.a.total), B);
    [1, 2, 3].forEach((lv, k) => put(ws, colL(5 + k) + r, F(`COUNTIFS(${pb},${PR('F')},${lv})`, p.a.lv[lv]), B));
    put(ws, `H${r}`, F(`SUMIFS(${PR('F')},${pb})`, p.lvSum), B);
    QBUCKETS.forEach((b, k) => {
      let f = `COUNTIFS(${pb},${PR('G')},${qs(b)})`;
      if (b === '0%') f += `+COUNTIFS(${pb},${PR('G')},"")`;
      put(ws, colL(9 + k) + r, F(f, p.a.q[b]), B);
    });
    put(ws, `P${r}`, 1, { ...B, numFmt: '0%' });
    put(ws, `Q${r}`, F(`IFERROR(AVERAGEIFS(${PR('H')},${pb}),0)`, p.a.avg), { ...B, numFmt: '0.0%' });
    put(ws, `R${r}`, F(`P${r}+Q${r}`, 1 + p.a.avg), { ...B, numFmt: '0.0%' });
    const hasNote = p.rows.some(x => x.note);
    put(ws, `S${r}`, F(`IF(COUNTIFS(${pb},${PR('I')},"<>")>0,"Có công việc đánh giá dạng chữ (không quy đổi %) - xem chi tiết ở sheet Data_Dự án","")`,
      hasNote ? 'Có công việc đánh giá dạng chữ (không quy đổi %) - xem chi tiết ở sheet Data_Dự án' : ''), { ...B, alignment: left });
    r++;
  });
  const first = 6, last = r - 1;
  const T = { font: { name: FN, size: 10, bold: true }, fill: solidFill(C.total), border: box(C.grey), alignment: center };
  styleRange(ws, `A${r}:S${r}`, T);
  put(ws, `B${r}`, 'TỔNG CỘNG / TRUNG BÌNH TẤT CẢ NHÂN VIÊN', { ...T, alignment: left });
  ws.getRow(r).height = 30;
  if (last >= first) {
    const sums = { D: 'total' };
    for (let c = 4; c <= 15; c++) {
      const L = colL(c);
      const res = m.projPersons.reduce((s, p) => s + (c === 4 ? p.a.total : c <= 7 ? p.a.lv[c - 4] : c === 8 ? p.lvSum : p.a.q[QBUCKETS[c - 9]]), 0);
      put(ws, `${L}${r}`, F(`SUM(${L}${first}:${L}${last})`, res), T);
    }
    const av = m.projPersons.reduce((s, p) => s + p.a.avg, 0) / m.projPersons.length;
    put(ws, `P${r}`, 1, { ...T, numFmt: '0%' });
    put(ws, `Q${r}`, F(`IFERROR(AVERAGE(Q${first}:Q${last}),0)`, av), { ...T, numFmt: '0.0%' });
    put(ws, `R${r}`, F(`P${r}+Q${r}`, 1 + av), { ...T, numFmt: '0.0%' });
  }
  r += 2;
  put(ws, `A${r}`, 'Ghi chú / Giả định của báo cáo:', { font: { name: FN, size: 10, bold: true } });
  [
    '- Bảng này gộp toàn bộ công việc của một nhân viên trên TẤT CẢ các dự án trong file Dự án (không tách theo từng dự án).',
    '- "Số dự án tham gia": số dự án khác nhau mà nhân viên có công việc được giao.',
    '- "Mức độ công việc": lấy theo trường "Mức độ công việc" trên hệ thống gốc (1 = thấp, 2 = trung bình, 3 = cao); "Tổng điểm mức độ" = tổng mức độ của các công việc.',
    '- "Phân bố số lượng công việc theo thang điểm đánh giá chất lượng": đếm theo cột "Đánh giá chất lượng" gốc (+3%/+2%/+1%/0%/-1%/-2%/-3%). Ô để trống được TÍNH VÀO cột 0% theo đúng quy tắc mặc định.',
    '- "Tỷ lệ hoàn thành (mặc định)": mặc định 100% cho mọi công việc theo yêu cầu.',
    '- "Chất lượng bổ sung TB (%)" = trung bình cộng các mức % ở trên (ô trống tính 0%). "Điểm đánh giá cuối cùng" = 100% + Chất lượng bổ sung TB.',
    '- Một số công việc có ô "Đánh giá chất lượng" ghi bằng chữ (ví dụ "Hoàn thành tốt") thay vì %, KHÔNG được quy đổi vào thang % ở trên; các nhân viên có công việc dạng này được đánh dấu ở cột "Ghi chú".',
    '- "(Chưa xác định)": các công việc không có tên người nhận việc trong file gốc.'
  ].forEach(t => { r++; ws.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 }; put(ws, `A${r}:S${r}`, t, { font: { name: FN, size: 9, italic: true, color: { argb: 'FF7F7F7F' } }, alignment: { vertical: 'middle' } }); });
  return ws;
}

/* ---------- Sheet Data / Data_Dự án ---------- */
const H_FORMULA = r => `IF(G${r}="",0,IF(ISNUMBER(SEARCH("%",G${r})),VALUE(SUBSTITUTE(SUBSTITUTE(G${r},"%",""),"+",""))/100,""))`;
function dataSheet(wb, m) {
  const ws = wb.addWorksheet('Data', { views: [{ state: 'frozen', ySplit: 1 }], properties: { tabColor: { argb: C.blue } } });
  const heads = ['Phòng ban', 'Người nhận việc', 'Tên công việc', 'Trạng thái', 'Mức độ (gốc)', 'Mức độ số', 'Đánh giá chất lượng (gốc)', 'Chất lượng % (số)', 'Ghi chú'];
  [12, 26.5, 57.9, 14, 10, 10, 20, 14, 46].forEach((w, i) => ws.getColumn(i + 1).width = w);
  const H = { font: { name: 'Arial', size: 11, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(C.blue), border: box(C.grey), alignment: center };
  heads.forEach((h, i) => put(ws, colL(i + 1) + '1', h, H));
  ws.getRow(1).height = 42.6;
  const B = { font: { name: 'Arial', size: 11 }, border: box(C.grey), alignment: center };
  const BL = { ...B, alignment: left };
  m.data.forEach((d, i) => {
    const r = i + 2, row = ws.getRow(r);
    row.values = [d.dept, d.name, d.task, d.status, d.lvRaw || null, F(`IFERROR(VALUE(E${r}),"")`, d.lvNum == null ? '' : d.lvNum),
      d.qRaw || null, F(H_FORMULA(r), d.q.pct == null ? '' : d.q.pct), d.note || null];
    for (let c = 1; c <= 9; c++) { const cell = row.getCell(c); const st = (c === 3 || c === 9 || c === 2) ? BL : B; cell.font = st.font; cell.border = st.border; cell.alignment = st.alignment; }
    row.getCell(8).numFmt = '0%';
  });
  ws.autoFilter = `A1:I${Math.max(2, m.data.length + 1)}`;
  return ws;
}
function projDataSheet(wb, m) {
  const ws = wb.addWorksheet('Data_Dự án', { views: [{ state: 'frozen', ySplit: 1 }], properties: { tabColor: { argb: C.orange } } });
  const heads = ['Dự án', 'Người nhận việc', 'Tên công việc', 'Trạng thái', 'Mức độ (gốc)', 'Mức độ số', 'Đánh giá chất lượng (gốc)', 'Chất lượng % (số)', 'Ghi chú',
    'Giai đoạn (Tasklist)', 'Mục tiêu bộ phận', 'Chi tiết thành phần công việc', 'Kinh doanh phụ trách', 'Khách hàng', 'Thời hạn', 'Hoàn thành thực tế', 'Nguyên nhân', 'Phương án xử lý', 'Trường dữ liệu tùy chỉnh (gốc)'];
  [37.5, 20, 42, 14, 10, 10, 20, 14, 40, 26, 20, 36, 26, 22, 16, 16, 18, 18, 60].forEach((w, i) => ws.getColumn(i + 1).width = w);
  const H = { font: { name: 'Arial', size: 10, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(C.orange), border: box(C.grey), alignment: center };
  const H2 = { ...H, fill: solidFill('FF7F7F7F') };
  heads.forEach((h, i) => put(ws, colL(i + 1) + '1', h, i < 9 ? H : H2));
  ws.getRow(1).height = 32;
  const B = { font: { name: 'Arial', size: 10 }, border: box(C.grey), alignment: center };
  const BL = { ...B, alignment: left };
  m.proj.forEach((d, i) => {
    const r = i + 2, row = ws.getRow(r);
    row.values = [d.project, d.name, d.task, d.status, d.lvRaw || null, F(`IFERROR(VALUE(E${r}),"")`, d.lvNum == null ? '' : d.lvNum),
      d.qRaw || null, F(H_FORMULA(r), d.q.pct == null ? '' : d.q.pct), d.note || null,
      d.phase, d.goal, d.detail, d.sales, d.customer, d.deadline, d.done, d.cause, d.fix, d.custom].map(v => v === '' ? null : v);
    for (let c = 1; c <= 19; c++) { const cell = row.getCell(c); const st = [1, 2, 3, 9, 10, 11, 12, 13, 14, 19].includes(c) ? BL : B; cell.font = st.font; cell.border = st.border; cell.alignment = st.alignment; }
    row.getCell(8).numFmt = '0%';
    row.height = 23.25;
  });
  ws.autoFilter = `A1:S${Math.max(2, m.proj.length + 1)}`;
  return ws;
}

/* ---------- Mastersheet ---------- */
function masterSheet(wb, m, ctx) {
  const { DR, N, period } = ctx;
  const ws = wb.addWorksheet('Mastersheet', { views: [{ showGridLines: false, zoomScale: 90 }], properties: { tabColor: { argb: 'FF7030A0' } } });
  ws.getColumn(1).width = 2; for (let c = 2; c <= 17; c++) ws.getColumn(c).width = 12.5; ws.getColumn(18).width = 2;
  ['S', 'T', 'U', 'V', 'W'].forEach(k => { ws.getColumn(k).width = 18; ws.getColumn(k).hidden = true; });
  const FN = 'Arial';
  const charts = [];
  const A = m.all;
  const secBar = (r, range, text, color = C.blue) => { put(ws, range, text, { font: { name: FN, size: 11, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(color), alignment: { vertical: 'middle', horizontal: 'left', indent: 1 } }); ws.getRow(r).height = 22; };
  const HD = { font: { name: FN, size: 9, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(C.hdr), border: box('FFD9D9D9'), alignment: center };
  const TD = { font: { name: FN, size: 10 }, border: box('FFD9D9D9'), alignment: center };
  const TDL = { ...TD, alignment: left };

  // Tiêu đề
  put(ws, 'B1:Q1', `MASTERSHEET · DASHBOARD KPI THÁNG ${m.opts.month || ''}`, { font: { name: FN, size: 18, bold: true, color: { argb: C.navy } }, alignment: { vertical: 'middle' } });
  ws.getRow(1).height = 32;
  put(ws, 'B2:Q2', `Kỳ báo cáo: ${period} · Nguồn: sheet "Data" (${m.data.length} công việc từ ${m.files.length} file) · Số liệu là công thức, tự cập nhật khi sửa sheet Data`,
    { font: { name: FN, size: 9, italic: true, color: { argb: 'FF7F7F7F' } }, alignment: { vertical: 'middle' } });

  // ① KPI tổng quan
  secBar(4, 'B4:Q4', '①  TỔNG QUAN TOÀN CÔNG TY');
  const DD = DR('D');
  const nPeople = m.people.filter(p => p.name !== UNKNOWN).length;
  const cards = [
    ['Tổng công việc', F(`COUNTA(${DR('C')})`, A.total), '#,##0', C.navy],
    ['Hoàn thành', F(`COUNTIFS(${DD},"Hoàn thành")`, A.st['Hoàn thành']), '#,##0', C.ok],
    ['Hoàn thành muộn', F(`COUNTIFS(${DD},"Hoàn thành muộn")`, A.st['Hoàn thành muộn']), '#,##0', C.late],
    ['Đang làm', F(`COUNTIFS(${DD},"Đang làm")`, A.st['Đang làm']), '#,##0', C.doing],
    ['Quá hạn', F(`COUNTIFS(${DD},"Quá hạn")`, A.st['Quá hạn']), '#,##0', C.over],
    ['Tỷ lệ hoàn thành', F(`IFERROR((D6+F6)/B6,0)`, A.rate), '0.0%', 'FF2E75B6'],
    ['Nhân sự', F(`SUMPRODUCT((${DR('B')}<>${qs(UNKNOWN)})/COUNTIF(${DR('B')},${DR('B')}))`, nPeople), '#,##0', 'FF7030A0'],
    ['Điểm đánh giá TB', F(`1+IFERROR(AVERAGE(${DR('H')}),0)`, A.score), '0.0%', 'FFBF8F00']
  ];
  cards.forEach(([lbl, val, fmt, col], i) => {
    const c1 = colL(2 + i * 2), c2 = colL(3 + i * 2);
    put(ws, `${c1}5:${c2}5`, lbl, { font: { name: FN, size: 9, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(col), alignment: center });
    put(ws, `${c1}6:${c2}6`, val, { font: { name: FN, size: 20, bold: true, color: { argb: col } }, fill: solidFill(C.pale), alignment: center, numFmt: fmt });
    const sub = i >= 1 && i <= 4 ? F(`IFERROR(${c1}6/$B$6,0)`, A.total ? [0, A.st['Hoàn thành'], A.st['Hoàn thành muộn'], A.st['Đang làm'], A.st['Quá hạn']][i] / A.total : 0) : '';
    put(ws, `${c1}7:${c2}7`, sub, { font: { name: FN, size: 9, italic: true, color: { argb: 'FF7F7F7F' } }, fill: solidFill(C.pale), alignment: center, numFmt: '0.0%" tổng CV"' });
  });
  ws.getRow(5).height = 20; ws.getRow(6).height = 36; ws.getRow(7).height = 16;

  // ② Theo phòng ban
  secBar(9, 'B9:K9', '②  TỔNG HỢP THEO PHÒNG BAN');
  put(ws, 'M9:Q9', 'Cơ cấu trạng thái toàn công ty', { font: { name: FN, size: 11, bold: true, color: { argb: 'FFFFFFFF' } }, fill: solidFill(C.blue), alignment: { vertical: 'middle', indent: 1 } });
  const dh = ['Phòng ban', null, 'Số nhân sự', 'Tổng CV', 'Hoàn thành', 'Hoàn thành muộn', 'Đang làm', 'Quá hạn', 'Tỷ lệ HT', 'Điểm TB'];
  put(ws, 'B10:C10', 'Phòng ban', HD); dh.slice(2).forEach((h, i) => put(ws, colL(4 + i) + '10', h, HD));
  ws.getRow(10).height = 30;
  let r = 11;
  const dFirst = r;
  m.deptSummary.forEach(ds => {
    const d = qs(ds.name), a = ds.a;
    put(ws, `B${r}:C${r}`, ds.name, { ...TDL, font: { name: FN, size: 10, bold: true } });
    put(ws, `D${r}`, F(`SUMPRODUCT((${DR('A')}=${d})/COUNTIFS(${DR('A')},${DR('A')},${DR('B')},${DR('B')}))`, ds.nPeople), TD);
    put(ws, `E${r}`, F(`COUNTIFS(${DR('A')},${d})`, a.total), TD);
    STATUSES.forEach((s, k) => put(ws, colL(6 + k) + r, F(`COUNTIFS(${DR('A')},${d},${DD},${qs(s)})`, a.st[s]), TD));
    put(ws, `J${r}`, F(`IFERROR((F${r}+G${r})/E${r},0)`, a.rate), { ...TD, numFmt: '0.0%' });
    put(ws, `K${r}`, F(`1+IFERROR(AVERAGEIFS(${DR('H')},${DR('A')},${d}),0)`, a.score), { ...TD, numFmt: '0.0%' });
    r++;
  });
  const dLast = r - 1;
  const TT = { font: { name: FN, size: 10, bold: true }, fill: solidFill(C.total), border: box('FFD9D9D9'), alignment: center };
  put(ws, `B${r}:C${r}`, 'Tổng', { ...TT, alignment: left });
  put(ws, `D${r}`, F(`$N$6`, nPeople), TT);
  ['E', 'F', 'G', 'H', 'I'].forEach((L, k) => put(ws, `${L}${r}`, F(`SUM(${L}${dFirst}:${L}${dLast})`, k === 0 ? A.total : A.st[STATUSES[k - 1]]), TT));
  put(ws, `J${r}`, F(`IFERROR((F${r}+G${r})/E${r},0)`, A.rate), { ...TT, numFmt: '0.0%' });
  put(ws, `K${r}`, F(`$P$6`, A.score), { ...TT, numFmt: '0.0%' });
  const dTot = r;
  charts.push({ kind: 'doughnut', title: 'Trạng thái công việc toàn công ty', from: [13, 10], to: [17, 24],
    cat: { ref: `Mastersheet!$F$10:$I$10`, vals: STATUSES }, val: { ref: `Mastersheet!$F$${dTot}:$I$${dTot}`, vals: STATUSES.map(s => A.st[s]) },
    colors: ['3AA76D', 'F2A93B', '4F81BD', 'E05353'] });
  const chartDeptTop = dTot + 2;
  charts.push({ kind: 'bar', dir: 'bar', grouping: 'stacked', title: 'Trạng thái công việc theo phòng ban', from: [2, chartDeptTop], to: [11, chartDeptTop + Math.max(12, m.deptSummary.length * 2 + 4)],
    legend: true, series: STATUSES.map((s, k) => ({ name: s, color: ['3AA76D', 'F2A93B', '4F81BD', 'E05353'][k],
      cat: { ref: `Mastersheet!$B$${dFirst}:$B$${dLast}`, vals: m.deptSummary.map(d => d.name) },
      val: { ref: `Mastersheet!$${colL(6 + k)}$${dFirst}:$${colL(6 + k)}$${dLast}`, vals: m.deptSummary.map(d => d.a.st[s]) } })) });
  r = Math.max(chartDeptTop + Math.max(12, m.deptSummary.length * 2 + 4) + 2, 26);

  // ③ Tra cứu nhân sự
  const s = r;
  secBar(s, `B${s}:Q${s}`, '③  TRA CỨU THEO NHÂN SỰ & ĐÁNH GIÁ CHẤT LƯỢNG  (chọn ở các ô vàng — danh sách thả xuống)', 'FF7030A0');
  const selRow = s + 1, SEL = `$D$${selRow}`;
  put(ws, `B${selRow}:C${selRow}`, '▶ Chọn nhân sự:', { font: { name: FN, size: 11, bold: true, color: { argb: C.navy } }, alignment: { vertical: 'middle' } });
  put(ws, `D${selRow}:F${selRow}`, ALL, { font: { name: FN, size: 12, bold: true, color: { argb: 'FF7030A0' } }, fill: solidFill(C.yellow), border: box('FFBF8F00'), alignment: center });
  // ô lọc "Đánh giá chất lượng" (áp dụng cho bảng chi tiết ⑦)
  const QSELROW = selRow, QSEL = `$J$${QSELROW}`;
  put(ws, `G${selRow}:I${selRow}`, '▶ Lọc đánh giá chất lượng:', { font: { name: FN, size: 11, bold: true, color: { argb: 'FFC55A11' } }, alignment: { horizontal: 'right', vertical: 'middle' } });
  put(ws, `J${selRow}:K${selRow}`, ALL, { font: { name: FN, size: 12, bold: true, color: { argb: 'FFC55A11' } }, fill: solidFill(C.yellow), border: box('FFBF8F00'), alignment: center, numFmt: '@' });
  put(ws, `L${selRow}:Q${selRow}`, '← bảng chi tiết ⑦ lọc theo cả 2 ô vàng', { font: { name: FN, size: 9, italic: true, color: { argb: 'FF7F7F7F' } }, alignment: { vertical: 'middle', wrapText: true } });
  const QOPT = [ALL, '+3%', '+2%', '+1%', '0%', '-1%', '-2%', '-3%', '(Trống)'];
  ws.getCell('X1').value = 'DS mức đánh giá';
  QOPT.forEach((o, i) => { ws.getCell(`X${i + 2}`).value = o; ws.getCell(`X${i + 2}`).numFmt = '@'; });
  ws.getColumn('X').hidden = true;
  ws.getCell(`J${selRow}`).dataValidation = { type: 'list', allowBlank: false, showErrorMessage: true, errorTitle: 'Giá trị không hợp lệ', error: 'Chọn một mức trong danh sách.', formulae: [`$X$2:$X$${QOPT.length + 1}`] };
  // V3 (ẩn): điều kiện đánh giá; "*" = tất cả, "" = ô trống; chuẩn hoá khi người dùng gõ số (VD 3%)
  ws.getCell('V3').value = F(`IF(OR(${QSEL}="",${QSEL}=${qs(ALL)}),"*",IF(${QSEL}="(Trống)","",IF(ISNUMBER(${QSEL}),IF(${QSEL}>0,"+","")&ROUND(${QSEL}*100,0)&"%",${QSEL})))`, '*');
  ws.getRow(selRow).height = 26;
  // danh sách chọn (cột ẩn S:U)
  const listNames = [ALL].concat(m.people.map(p => p.name));
  put(ws, 'S1', 'DS nhân sự'); put(ws, 'T1', 'Phòng ban'); put(ws, 'U1', 'Số dự án');
  listNames.forEach((n, i) => {
    ws.getCell(`S${i + 2}`).value = n;
    if (i === 0) { ws.getCell('T2').value = 'Toàn công ty'; ws.getCell('U2').value = m.projects.length; }
    else { const p = m.people[i - 1]; ws.getCell(`T${i + 2}`).value = p.deptLabel; ws.getCell(`U${i + 2}`).value = p.nProjects; }
  });
  const listEnd = listNames.length + 1;
  ws.getCell(`D${selRow}`).dataValidation = { type: 'list', allowBlank: false, showErrorMessage: true, errorTitle: 'Tên không hợp lệ', error: 'Vui lòng chọn tên trong danh sách.', formulae: [`$S$2:$S$${listEnd}`] };
  ws.getCell('V1').value = F(`IF(OR(${SEL}="",${SEL}=${qs(ALL)}),"*",${SEL})`, '*');
  const CR = '$V$1';
  const infoRow = s + 2;
  const KV = { font: { name: FN, size: 10, bold: true, color: { argb: C.navy } }, alignment: { vertical: 'middle' } };
  put(ws, `B${infoRow}:C${infoRow}`, 'Phòng ban / nguồn:', KV);
  put(ws, `D${infoRow}:K${infoRow}`, F(`IFERROR(VLOOKUP(${SEL},$S$2:$U$${listEnd},2,0),"")`, 'Toàn công ty'), { font: { name: FN, size: 10 }, alignment: { vertical: 'middle' } });
  put(ws, `L${infoRow}:M${infoRow}`, 'Số dự án tham gia:', KV);
  put(ws, `N${infoRow}`, F(`IFERROR(VLOOKUP(${SEL},$S$2:$U$${listEnd},3,0),0)`, m.projects.length), { font: { name: FN, size: 10, bold: true }, alignment: { horizontal: 'left', vertical: 'middle' } });
  const mh = s + 4;
  put(ws, `B${mh}:C${mh}`, 'Chỉ số', HD); put(ws, `D${mh}`, 'Giá trị', HD); put(ws, `F${mh}:G${mh}`, 'Đánh giá chất lượng', HD); put(ws, `H${mh}`, 'Số CV', HD);
  const cB = col => DR(col);
  const cnt = extra => `COUNTIFS(${cB('B')},${CR}${extra || ''})`;
  const lvNo = A.total - A.lv[1] - A.lv[2] - A.lv[3];
  const leftRows = [
    ['Tổng số công việc', F(cnt(), A.total), '#,##0'],
    ['Hoàn thành', F(cnt(`,${DD},"Hoàn thành"`), A.st['Hoàn thành']), '#,##0'],
    ['Hoàn thành muộn', F(cnt(`,${DD},"Hoàn thành muộn"`), A.st['Hoàn thành muộn']), '#,##0'],
    ['Đang làm', F(cnt(`,${DD},"Đang làm"`), A.st['Đang làm']), '#,##0'],
    ['Quá hạn', F(cnt(`,${DD},"Quá hạn"`), A.st['Quá hạn']), '#,##0'],
    ['Tỷ lệ hoàn thành', F(`IFERROR((D${mh + 2}+D${mh + 3})/D${mh + 1},0)`, A.rate), '0.0%'],
    ['Mức 1 - Cơ bản', F(cnt(`,${cB('F')},1`), A.lv[1]), '#,##0'],
    ['Mức 2 - Trung bình', F(cnt(`,${cB('F')},2`), A.lv[2]), '#,##0'],
    ['Mức 3 - Phức tạp', F(cnt(`,${cB('F')},3`), A.lv[3]), '#,##0'],
    ['Chưa có / ngoài thang mức độ', F(`D${mh + 1}-SUM(D${mh + 7}:D${mh + 9})`, lvNo), '#,##0']
  ];
  const rightRows = QBUCKETS.map(b => [b === '0%' ? '0% (gồm ô trống)' : b,
    F(cnt(`,${cB('G')},${qs(b)}`) + (b === '0%' ? `+${cnt(`,${cB('G')},""`)}` : ''), A.q[b]), '#,##0'])
    .concat([
      ['Chất lượng bổ sung TB', F(`IFERROR(AVERAGEIFS(${cB('H')},${cB('B')},${CR}),0)`, A.avg), '+0.00%;-0.00%;0.00%'],
      ['Điểm đánh giá cuối cùng', F(`1+H${mh + 8}`, A.score), '0.0%'],
      ['Tỷ lệ đúng hạn (HT / HT+muộn)', F(`IFERROR(D${mh + 2}/(D${mh + 2}+D${mh + 3}),0)`, A.done ? A.st['Hoàn thành'] / A.done : 0), '0.0%']
    ]);
  leftRows.forEach(([l, v, fmt], i) => {
    const rr = mh + 1 + i;
    const hi = i === 0 || i === 5;
    put(ws, `B${rr}:C${rr}`, l, { ...TDL, font: { name: FN, size: 10, bold: hi } });
    put(ws, `D${rr}`, v, { ...TD, numFmt: fmt, font: { name: FN, size: 10, bold: true, color: { argb: [C.navy, C.ok, C.late, C.doing, C.over][i] || 'FF262626' } } });
  });
  rightRows.forEach(([l, v, fmt], i) => {
    const rr = mh + 1 + i;
    const hi = i >= 7;
    put(ws, `F${rr}:G${rr}`, l, { ...TDL, font: { name: FN, size: 10, bold: hi } });
    put(ws, `H${rr}`, v, { ...TD, numFmt: fmt, font: { name: FN, size: 10, bold: true, color: { argb: i === 8 ? 'FFBF8F00' : 'FF262626' } } });
  });
  for (let i = 0; i <= 10; i++) ws.getRow(mh + i).height = 19;
  charts.push({ kind: 'bar', dir: 'col', grouping: 'clustered', title: 'Trạng thái công việc - nhân sự được chọn', from: [10, mh], to: [17, mh + 13],
    series: [{ name: 'Số CV', color: '4F81BD', pointColors: ['3AA76D', 'F2A93B', '4F81BD', 'E05353'],
      cat: { ref: `Mastersheet!$B$${mh + 2}:$B$${mh + 5}`, vals: STATUSES }, val: { ref: `Mastersheet!$D$${mh + 2}:$D$${mh + 5}`, vals: STATUSES.map(x => A.st[x]) } }] });
  const q2 = mh + 13;
  charts.push({ kind: 'bar', dir: 'col', grouping: 'clustered', title: 'Phân bố đánh giá chất lượng - nhân sự được chọn', from: [2, q2], to: [9, q2 + 13],
    series: [{ name: 'Số CV', color: '4F81BD', pointColors: ['1B7F4B', '3AA76D', '8FD19E', 'BFBFBF', 'F7B267', 'F08A4B', 'D64545'],
      cat: { ref: `Mastersheet!$F$${mh + 1}:$F$${mh + 7}`, vals: QBUCKETS }, val: { ref: `Mastersheet!$H$${mh + 1}:$H$${mh + 7}`, vals: QBUCKETS.map(b => A.q[b]) } }] });
  charts.push({ kind: 'bar', dir: 'col', grouping: 'clustered', title: 'Mức độ công việc - nhân sự được chọn', from: [10, q2 + 1], to: [17, q2 + 13],
    series: [{ name: 'Số CV', color: '4F81BD', pointColors: ['9DC3E6', '4F81BD', '1F3864'],
      cat: { ref: `Mastersheet!$B$${mh + 7}:$B$${mh + 9}`, vals: ['Mức 1 - Cơ bản', 'Mức 2 - Trung bình', 'Mức 3 - Phức tạp'] }, val: { ref: `Mastersheet!$D$${mh + 7}:$D$${mh + 9}`, vals: [A.lv[1], A.lv[2], A.lv[3]] } }] });
  r = q2 + 15;

  // ④ Top 5 & ⑤ Top 3
  const t = r;
  secBar(t, `B${t}:L${t}`, '④  TOP 5 NHÂN SỰ XỬ LÝ NHIỀU VIỆC NHẤT', 'FF2E75B6');
  const topHead = (rr, cols) => cols.forEach(([range, txt]) => put(ws, range.replace(/#/g, rr), txt, HD));
  topHead(t + 1, [['B#', 'Hạng'], ['C#:E#', 'Nhân sự'], ['F#:G#', 'Phòng ban / nguồn'], ['H#', 'Tổng CV'], ['I#', 'Hoàn thành'], ['J#', 'Đang làm'], ['K#', 'Quá hạn'], ['L#', 'Điểm']]);
  ws.getRow(t + 1).height = 26;
  const rankFill = ['FFFFD966', 'FFD9D9D9', 'FFF4B183', C.pale, C.pale];
  m.top5.forEach((p, i) => {
    const rr = t + 2 + i, n = qs(p.name);
    put(ws, `B${rr}`, i + 1, { ...TD, fill: solidFill(rankFill[i]), font: { name: FN, size: 11, bold: true } });
    put(ws, `C${rr}:E${rr}`, p.name, { ...TDL, font: { name: FN, size: 10, bold: true } });
    put(ws, `F${rr}:G${rr}`, p.deptLabel, TDL);
    put(ws, `H${rr}`, F(`COUNTIFS(${cB('B')},${n})`, p.a.total), { ...TD, font: { name: FN, size: 11, bold: true, color: { argb: C.navy } } });
    put(ws, `I${rr}`, F(`COUNTIFS(${cB('B')},${n},${DD},"Hoàn thành")+COUNTIFS(${cB('B')},${n},${DD},"Hoàn thành muộn")`, p.a.done), TD);
    put(ws, `J${rr}`, F(`COUNTIFS(${cB('B')},${n},${DD},"Đang làm")`, p.a.st['Đang làm']), TD);
    put(ws, `K${rr}`, F(`COUNTIFS(${cB('B')},${n},${DD},"Quá hạn")`, p.a.st['Quá hạn']), TD);
    put(ws, `L${rr}`, F(`1+IFERROR(AVERAGEIFS(${cB('H')},${cB('B')},${n}),0)`, p.a.score), { ...TD, numFmt: '0.0%' });
    ws.getRow(rr).height = 20;
  });
  if (m.top5.length) charts.push({ kind: 'bar', dir: 'bar', grouping: 'clustered', title: 'Top 5 - số công việc xử lý', from: [13, t], to: [17, t + 14],
    series: [{ name: 'Tổng CV', color: '2E75B6', cat: { ref: `Mastersheet!$C$${t + 2}:$C$${t + 1 + m.top5.length}`, vals: m.top5.map(p => p.name) },
      val: { ref: `Mastersheet!$H$${t + 2}:$H$${t + 1 + m.top5.length}`, vals: m.top5.map(p => p.a.total) } }] });
  const t3 = t + 8;
  secBar(t3, `B${t3}:L${t3}`, `⑤  TOP 3 ĐIỂM ĐÁNH GIÁ TỔNG HỢP CAO NHẤT  (tối thiểu ${m.opts.minTasks} CV)`, 'FFBF8F00');
  topHead(t3 + 1, [['B#', 'Hạng'], ['C#:E#', 'Nhân sự'], ['F#:G#', 'Phòng ban / nguồn'], ['H#', 'Tổng CV'], ['I#', 'CV tính điểm'], ['J#', 'CL bổ sung TB'], ['K#:L#', 'Điểm cuối cùng']]);
  ws.getRow(t3 + 1).height = 26;
  m.top3.forEach((p, i) => {
    const rr = t3 + 2 + i, n = qs(p.name);
    put(ws, `B${rr}`, i + 1, { ...TD, fill: solidFill(rankFill[i]), font: { name: FN, size: 11, bold: true } });
    put(ws, `C${rr}:E${rr}`, p.name, { ...TDL, font: { name: FN, size: 10, bold: true } });
    put(ws, `F${rr}:G${rr}`, p.deptLabel, TDL);
    put(ws, `H${rr}`, F(`COUNTIFS(${cB('B')},${n})`, p.a.total), TD);
    put(ws, `I${rr}`, F(`COUNTIFS(${cB('B')},${n},${cB('I')},"")`, p.a.qCnt), TD);
    put(ws, `J${rr}`, F(`IFERROR(AVERAGEIFS(${cB('H')},${cB('B')},${n}),0)`, p.a.avg), { ...TD, numFmt: '+0.00%;-0.00%;0.00%' });
    put(ws, `K${rr}:L${rr}`, F(`1+J${rr}`, p.a.score), { ...TD, numFmt: '0.0%', font: { name: FN, size: 12, bold: true, color: { argb: 'FFBF8F00' } } });
    ws.getRow(rr).height = 20;
  });
  r = t + 16;

  // ⑥ Insight
  secBar(r, `B${r}:Q${r}`, '⑥  INSIGHT TỰ ĐỘNG', 'FFC00000');
  m.insights.forEach(x => {
    r++;
    put(ws, `B${r}:Q${r}`, `${x.i}  ${x.t.replace(/<[^>]+>/g, '')}`, { font: { name: FN, size: 10 }, alignment: { vertical: 'middle', wrapText: true, indent: 1 }, border: { bottom: thin('FFE7E6E6') } });
    ws.getRow(r).height = 30;
  });
  r += 2;

  // ⑦ Chi tiết công việc theo bộ lọc mục ③ (nhân sự × đánh giá chất lượng)
  secBar(r, `B${r}:Q${r}`, '⑦  CHI TIẾT CÔNG VIỆC THEO BỘ LỌC Ở MỤC ③  (nhân sự × đánh giá chất lượng)', 'FF7030A0');
  r++;
  const bR = cB('B'), gR = cB('G');
  // điều kiện mảng: đúng nhân sự (hoặc tất cả) VÀ đúng mức đánh giá ("0%" gồm cả ô trống theo quy tắc)
  const COND = `((${bR}=${SEL})+(${CR}="*"))*(($V$3="*")+(${gR}=$V$3)+(($V$3="0%")*(${gR}="")))`;
  ws.getCell('V4').value = F(`SUMPRODUCT(--(${COND}>0))`, m.data.length);
  put(ws, `B${r}:Q${r}`, F(`IF(AND(${CR}="*",$V$3="*"),"Toàn bộ "&$V$4&" công việc (chưa lọc) - chọn nhân sự và/hoặc mức đánh giá ở mục ③ để lọc.","Có "&$V$4&" công việc"&IF(${CR}<>"*"," của "&${SEL},"")&IF($V$3<>"*"," có đánh giá "&IF($V$3="","(trống)",IF($V$3="0%","0% (gồm ô trống)",$V$3)),""))`,
    `Toàn bộ ${m.data.length} công việc (chưa lọc) - chọn nhân sự và/hoặc mức đánh giá ở mục ③ để lọc.`), { font: { name: FN, size: 10, bold: true, italic: true, color: { argb: 'FF7030A0' } }, alignment: { vertical: 'middle' } });
  r++;
  topHead(r, [['B#', 'STT'], ['C#:D#', 'Nhân viên'], ['E#:F#', 'Phòng ban'], ['G#:K#', 'Tên công việc'], ['L#:M#', 'Trạng thái'], ['N#', 'Mức độ'], ['O#', 'Đánh giá CL'], ['P#:Q#', 'Ghi chú']]);
  ws.getRow(r).height = 24;
  const maxRows = Math.min(1500, Math.max(10, m.data.length));
  const detFirst = r + 1;
  for (let k = 1; k <= maxRows; k++) {
    const rr = r + k, W = `$W${rr}`, d = m.data[k - 1];
    ws.getCell(`W${rr}`).value = F(`IFERROR(_xlfn.AGGREGATE(15,6,ROW(${bR})/(${COND}>0),${k}),"")`, d ? k + 1 : '');
    const idx = (col, v) => F(`IF(${W}="","",INDEX(Data!$${col}:$${col},${W})&"")`, d ? (v == null ? '' : String(v)) : '');
    ws.getRow(rr).height = 24;
    put(ws, `B${rr}`, F(`IF(${W}="","",${k})`, d ? k : ''), TD);
    put(ws, `C${rr}:D${rr}`, idx('B', d && d.name), TDL);
    put(ws, `E${rr}:F${rr}`, idx('A', d && d.dept), TDL);
    put(ws, `G${rr}:K${rr}`, idx('C', d && d.task), TDL);
    put(ws, `L${rr}:M${rr}`, idx('D', d && d.status), TD);
    put(ws, `N${rr}`, idx('E', d && d.lvRaw), TD);
    put(ws, `O${rr}`, idx('G', d && d.qRaw), TD);
    put(ws, `P${rr}:Q${rr}`, idx('I', d && d.note), { ...TDL, font: { name: FN, size: 8, color: { argb: 'FF7F7F7F' } } });
  }
  const detLast = r + maxRows;
  const stRule = (txt, font, bg) => ({ type: 'expression', formulae: [`$L${detFirst}=${qs(txt)}`], style: { font: { color: { argb: font }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: bg } } } });
  ws.addConditionalFormatting({ ref: `L${detFirst}:M${detLast}`, rules: [
    stRule('Quá hạn', 'FFC00000', 'FFFDE2E2'), stRule('Hoàn thành muộn', 'FF9C5700', 'FFFFEB9C'),
    stRule('Đang làm', 'FF1F4E79', 'FFDDEBF7'), stRule('Hoàn thành', 'FF006100', 'FFE2F0D9')] });
  const qRule = (prefix, font, bg) => ({ type: 'expression', formulae: [`LEFT($O${detFirst},1)="${prefix}"`], style: { font: { color: { argb: font }, bold: true }, fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: bg } } } });
  ws.addConditionalFormatting({ ref: `O${detFirst}:O${detLast}`, rules: [qRule('+', 'FF006100', 'FFE2F0D9'), qRule('-', 'FFC00000', 'FFFDE2E2')] });
  ws.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printArea: `A1:Q${detLast}`,
    margins: { left: 0.3, right: 0.3, top: 0.4, bottom: 0.4, header: 0.2, footer: 0.2 } };
  return { ws, charts };
}

/* ---------- dựng workbook ---------- */
async function buildWorkbook(m) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Tool tổng hợp KPI'; wb.created = new Date();
  wb.calcProperties = { fullCalcOnLoad: true };
  const N = Math.max(2, m.data.length + 1), M = Math.max(2, m.proj.length + 1);
  const ctx = {
    N, M, period: m.opts.period,
    DR: c => `Data!$${c}$2:$${c}$${N}`,
    PR: c => `'Data_Dự án'!$${c}$2:$${c}$${M}`,
    hasProj: m.proj.length > 0
  };
  const used = new Set(['mastersheet', 'nguyên tắc đánh giá', 'data', 'data_dự án', 'tổng hợp dự án']);
  const { charts } = masterSheet(wb, m, ctx);
  rulesSheet(wb, 'Nguyên tắc đánh giá');
  const sheetList = [{ name: 'Mastersheet', desc: 'Dashboard tổng quan, lọc theo nhân sự, biểu đồ, Top 5/Top 3, insight' },
    { name: 'Nguyên tắc đánh giá', desc: 'Quy tắc tính điểm' }];
  for (const d of m.depts) {
    const sn = safeSheetName(d.name, used);
    deptSheet(wb, sn, d, ctx);
    sheetList.push({ name: sn, desc: `Phòng ban ${d.name}: ${d.persons.length} nhân sự, ${d.rows.length} CV` + (d.persons.some(p => p.proj.length) ? ' (+ cộng việc Dự án cho người trùng tên)' : '') });
  }
  if (m.proj.length) {
    projectSheet(wb, 'Tổng hợp Dự án', m, ctx);
    sheetList.push({ name: 'Tổng hợp Dự án', desc: `${m.projPersons.length} nhân sự trên ${m.projects.length} dự án (form "Kỹ thuật (2)")` });
  }
  dataSheet(wb, m); sheetList.push({ name: 'Data', desc: `${m.data.length} dòng = file phòng ban + file dự án` });
  projDataSheet(wb, m); sheetList.push({ name: 'Data_Dự án', desc: `${m.proj.length} dòng, đã tách "Trường dữ liệu tùy chỉnh"` });
  wb.views = [{ x: 0, y: 0, width: 24000, height: 14000, firstSheet: 0, activeTab: 0, visibility: 'visible' }];
  const buf = await wb.xlsx.writeBuffer();
  const out = await injectCharts(buf, 'Mastersheet', charts);
  return { blob: out, sheetList };
}
