/* ================= CHARTS: chèn biểu đồ Excel (DrawingML) vào sheet ================= */
const xesc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const strCache = a => `<c:strCache><c:ptCount val="${a.length}"/>${a.map((v, i) => `<c:pt idx="${i}"><c:v>${xesc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
const numCache = a => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${a.length}"/>${a.map((v, i) => `<c:pt idx="${i}"><c:v>${+v || 0}</c:v></c:pt>`).join('')}</c:numCache>`;
const txPr = (sz = 900, b = 0, col = '404040') => `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${b}"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="vi-VN"/></a:p></c:txPr>`;
const spSolid = c => `<c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill></c:spPr>`;
const chartTitle = t => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1100" b="1"/></a:pPr><a:r><a:rPr lang="vi-VN" sz="1100" b="1"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:rPr><a:t>${xesc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
const dLbls = (pos, fmt = '#,##0;;;') => `<c:dLbls><c:numFmt formatCode="${xesc(fmt)}" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(900, 1, pos === 'ctr' ? 'FFFFFF' : '262626')}${pos ? `<c:dLblPos val="${pos}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
const legendXml = () => `<c:legend><c:legendPos val="b"/><c:overlay val="0"/>${txPr(900)}</c:legend>`;
const catRef = c => `<c:cat><c:strRef><c:f>${xesc(c.ref)}</c:f>${strCache(c.vals)}</c:strRef></c:cat>`;
const valRef = v => `<c:val><c:numRef><c:f>${xesc(v.ref)}</c:f>${numCache(v.vals)}</c:numRef></c:val>`;

function barChartBody(s) {
  const stacked = s.grouping === 'stacked', horiz = s.dir === 'bar';
  const sers = s.series.map((se, i) =>
    `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx><c:v>${xesc(se.name)}</c:v></c:tx>${spSolid(se.color)}<c:invertIfNegative val="0"/>` +
    (se.pointColors || []).map((pc, j) => `<c:dPt><c:idx val="${j}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/>${spSolid(pc)}</c:dPt>`).join('') +
    dLbls(stacked ? 'ctr' : 'outEnd') + catRef(se.cat) + valRef(se.val) + `</c:ser>`).join('');
  return `<c:barChart><c:barDir val="${s.dir}"/><c:grouping val="${s.grouping}"/><c:varyColors val="0"/>${sers}` +
    `<c:gapWidth val="${stacked ? 50 : 70}"/>${stacked ? '<c:overlap val="100"/>' : ''}<c:axId val="50010"/><c:axId val="50020"/></c:barChart>` +
    `<c:catAx><c:axId val="50010"/><c:scaling><c:orientation val="${horiz ? 'maxMin' : 'minMax'}"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'l' : 'b'}"/>` +
    `<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    `<c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>${txPr(900, 0, '404040')}` +
    `<c:crossAx val="50020"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="50020"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${horiz ? 'b' : 'l'}"/>` +
    `<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="EDEDED"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>` +
    `<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>` +
    `<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>${txPr(800, 0, '8C8C8C')}<c:crossAx val="50010"/><c:crosses val="${horiz ? 'max' : 'autoZero'}"/><c:crossBetween val="between"/></c:valAx>`;
}
function doughnutBody(s) {
  const pts = s.colors.map((c, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${c}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('');
  return `<c:doughnutChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Số CV</c:v></c:tx>${pts}` +
    `<c:dLbls><c:numFmt formatCode="#,##0;;;" sourceLinked="0"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${txPr(1000, 1, 'FFFFFF')}` +
    `<c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/><c:showLeaderLines val="0"/></c:dLbls>` +
    catRef(s.cat) + valRef(s.val) + `</c:ser><c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
}
function chartXml(s) {
  const body = s.kind === 'doughnut' ? doughnutBody(s) : barChartBody(s);
  const legend = (s.kind === 'doughnut' || s.legend) ? legendXml() : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="${NS_C}" xmlns:a="${NS_A}" xmlns:r="${NS_R}">` +
    `<c:roundedCorners val="0"/><c:chart>${chartTitle(s.title)}<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/>${body}` +
    `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>${legend}<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr>` +
    `${txPr(900)}</c:chartSpace>`;
}
function drawingXml(charts) {
  const anchors = charts.map((s, i) =>
    `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${s.from[0] - 1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${s.from[1] - 1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${s.to[0]}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${s.to[1]}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="Chart ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="${NS_A}">${anchors}</xdr:wsDr>`;
}

async function injectCharts(buf, sheetName, charts) {
  const zip = await JSZip.loadAsync(buf);
  const P = new DOMParser();
  const wbDoc = P.parseFromString(await zip.file('xl/workbook.xml').async('string'), 'application/xml');
  const sheetEl = [...wbDoc.getElementsByTagName('sheet')].find(e => e.getAttribute('name') === sheetName);
  const rid = sheetEl.getAttribute('r:id') || sheetEl.getAttributeNS(NS_R, 'id');
  const relDoc = P.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'application/xml');
  const rel = [...relDoc.getElementsByTagName('Relationship')].find(e => e.getAttribute('Id') === rid);
  let target = rel.getAttribute('Target').replace(/^\/?xl\//, '').replace(/^\//, '');
  const sheetPath = 'xl/' + target;                               // xl/worksheets/sheet1.xml
  const sheetFile = sheetPath.split('/').pop();
  const sheetRelsPath = `xl/worksheets/_rels/${sheetFile}.rels`;

  let di = 1; while (zip.file(`xl/drawings/drawing${di}.xml`)) di++;
  let ci = 1; while (zip.file(`xl/charts/chart${ci}.xml`)) ci++;
  const drawPath = `xl/drawings/drawing${di}.xml`;
  const ctAdd = [`<Override PartName="/${drawPath}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`];
  const drawRels = [];
  charts.forEach((s, i) => {
    const n = ci + i;
    zip.file(`xl/charts/chart${n}.xml`, chartXml(s));
    ctAdd.push(`<Override PartName="/xl/charts/chart${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
    drawRels.push(`<Relationship Id="rId${i + 1}" Type="${NS_R}/chart" Target="../charts/chart${n}.xml"/>`);
  });
  zip.file(drawPath, drawingXml(charts));
  zip.file(`xl/drawings/_rels/drawing${di}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawRels.join('')}</Relationships>`);

  // rels của sheet
  const drawRelId = 'rIdKpiDrawing1';
  const drawRel = `<Relationship Id="${drawRelId}" Type="${NS_R}/drawing" Target="../drawings/drawing${di}.xml"/>`;
  if (zip.file(sheetRelsPath)) {
    const x = await zip.file(sheetRelsPath).async('string');
    zip.file(sheetRelsPath, x.replace('</Relationships>', drawRel + '</Relationships>'));
  } else {
    zip.file(sheetRelsPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${drawRel}</Relationships>`);
  }
  // thẻ <drawing> trong sheet (đúng thứ tự schema)
  let sx = await zip.file(sheetPath).async('string');
  if (!/xmlns:r=/.test(sx.slice(0, 600))) sx = sx.replace('<worksheet ', `<worksheet xmlns:r="${NS_R}" `);
  const tag = `<drawing r:id="${drawRelId}"/>`;
  const after = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst', '</worksheet>'];
  let pos = -1;
  for (const a of after) { const p = sx.indexOf(a); if (p >= 0 && (pos < 0 || p < pos)) pos = p; }
  sx = sx.slice(0, pos) + tag + sx.slice(pos);
  zip.file(sheetPath, sx);
  // content types
  let ct = await zip.file('[Content_Types].xml').async('string');
  ct = ct.replace('</Types>', ctAdd.join('') + '</Types>');
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
