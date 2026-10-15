// Spike: สร้าง XML ของกราฟ (DrawingML chart) และ drawing สำหรับแทรกลง .xlsx
// ไม่ใช่โค้ดของแอป — ใช้ทดลองความเป็นไปได้เท่านั้น

export const NS = {
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  xdr: 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
};

export const REL = {
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
};

export const CT = {
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
};

export function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** อ้างช่วงในสูตรกราฟ: ใส่ ' รอบชื่อชีตเสมอ (ชื่อไทย/มีช่องว่าง) และ ' ในชื่อเป็น '' */
export function ref(sheetName, range) {
  return `'${sheetName.replace(/'/g, "''")}'!${range}`;
}

function strCache(values) {
  return `<c:strCache><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`)
    .join('')}</c:strCache>`;
}

function numCache(values, formatCode = 'General') {
  return `<c:numCache><c:formatCode>${esc(formatCode)}</c:formatCode><c:ptCount val="${values.length}"/>${values
    .map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)
    .join('')}</c:numCache>`;
}

/**
 * series: { nameRef, name, catRef, categories, valRef, values }
 * cache (strCache/numCache) ใส่ไว้ให้โปรแกรมที่ไม่คำนวณใหม่ยังวาดกราฟได้ทันที
 */
function seriesXml(type, s) {
  const tx = `<c:tx><c:strRef><c:f>${esc(s.nameRef)}</c:f>${strCache([s.name])}</c:strRef></c:tx>`;
  const cat = `<c:cat><c:strRef><c:f>${esc(s.catRef)}</c:f>${strCache(s.categories)}</c:strRef></c:cat>`;
  const val = `<c:val><c:numRef><c:f>${esc(s.valRef)}</c:f>${numCache(s.values)}</c:numRef></c:val>`;
  const head = `<c:idx val="0"/><c:order val="0"/>${tx}`;
  if (type === 'bar') return `<c:ser>${head}<c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
  if (type === 'line') return `<c:ser>${head}<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>${cat}${val}<c:smooth val="0"/></c:ser>`;
  if (type === 'pie') return `<c:ser>${head}${cat}${val}</c:ser>`;
  throw new Error('unknown chart type ' + type);
}

const AX_CAT = 500000001;
const AX_VAL = 500000002;

function axesXml() {
  return (
    `<c:catAx><c:axId val="${AX_CAT}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="nextTo"/><c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${AX_VAL}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="#,##0" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="nextTo"/><c:crossAx val="${AX_CAT}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  );
}

export function chartXml(type, title, series) {
  let plot;
  if (type === 'bar') {
    plot =
      `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${seriesXml('bar', series)}` +
      `<c:gapWidth val="150"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:barChart>${axesXml()}`;
  } else if (type === 'line') {
    plot =
      `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${seriesXml('line', series)}` +
      `<c:marker val="1"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:lineChart>${axesXml()}`;
  } else {
    plot = `<c:pieChart><c:varyColors val="1"/>${seriesXml('pie', series)}<c:firstSliceAng val="0"/></c:pieChart>`;
  }
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="${NS.c}" xmlns:a="${NS.a}" xmlns:r="${NS.r}">` +
    `<c:roundedCorners val="0"/>` +
    `<c:chart>` +
    `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` +
    `<c:autoTitleDeleted val="0"/>` +
    `<c:plotArea><c:layout/>${plot}</c:plotArea>` +
    `<c:legend><c:legendPos val="${type === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend>` +
    `<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/>` +
    `</c:chart>` +
    `</c:chartSpace>`
  );
}

/** anchor: { fromCol, fromRow, toCol, toRow } (เริ่มที่ 0) */
export function drawingXml(anchor, chartRelId, name = 'Chart 1') {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<xdr:wsDr xmlns:xdr="${NS.xdr}" xmlns:a="${NS.a}">` +
    `<xdr:twoCellAnchor editAs="oneCell">` +
    `<xdr:from><xdr:col>${anchor.fromCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.fromRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${anchor.toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${anchor.toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
    `<a:graphic><a:graphicData uri="${NS.c}"><c:chart xmlns:c="${NS.c}" xmlns:r="${NS.r}" r:id="${chartRelId}"/></a:graphicData></a:graphic>` +
    `</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>` +
    `</xdr:wsDr>`
  );
}

export function relsXml(rels) {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${esc(r.target)}"/>`).join('') +
    `</Relationships>`
  );
}
