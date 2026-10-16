/**
 * กราฟใน .xlsx (5D): ExcelJS 4.4 เขียนกราฟไม่ได้และทิ้งกราฟตอนอ่าน
 * จึงแทรกส่วนกราฟลง zip เองหลัง ExcelJS เขียนไฟล์ และอ่านกราฟจาก zip เองตอนเปิด
 * (ผลทดลองใน scripts/chart-spike/ — Google Sheets แสดงกราฟแท่ง/เส้น/วงกลมได้)
 */
import JSZip from 'jszip';
import type { Sheet, SheetChart, ChartType } from './model';
import { ChartData, chartToTwoCellAnchor, twoCellAnchorToChart, CellAnchorEmu, EMU_PER_PX, MIN_CHART_WIDTH, MIN_CHART_HEIGHT, pxToAnchor, newChartId } from './chart';
import { colToStr, parseA1, parseRange } from './address';
import { sameSheetName } from './sheetName';
import { XmlNode, XML_NS, parseXml, child, childrenOf, descendants, textOf, escapeXml } from './xmlTree';

const REL_TYPE = {
  drawing: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
  chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  chartsheet: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet',
};
const CONTENT_TYPE = {
  chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
  drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
};

// ---------------------------------------------------------------------------
// ตำแหน่งไฟล์ในแพ็กเกจ
// ---------------------------------------------------------------------------

interface Rel {
  id: string;
  type: string;
  target: string;
}

function relsPathFor(partPath: string): string {
  const i = partPath.lastIndexOf('/');
  return `${partPath.slice(0, i)}/_rels/${partPath.slice(i + 1)}.rels`;
}

/** แปลง Target ที่สัมพัทธ์กับไฟล์ต้นทาง เป็น path ในแพ็กเกจ */
function resolveTarget(fromPart: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = fromPart.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

async function readText(zip: JSZip, path: string): Promise<string | undefined> {
  return zip.file(path)?.async('string');
}

async function readRels(zip: JSZip, partPath: string): Promise<Rel[]> {
  const xml = await readText(zip, relsPathFor(partPath));
  if (!xml) return [];
  return childrenOf(parseXml(xml), XML_NS.pkgRels, 'Relationship').map((r) => ({
    id: r.attrs.Id,
    type: r.attrs.Type,
    target: resolveTarget(partPath, r.attrs.Target),
  }));
}

function relsXml(rels: Array<{ id: string; type: string; target: string }>): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    rels.map((r) => `<Relationship Id="${r.id}" Type="${r.type}" Target="${escapeXml(r.target)}"/>`).join('') +
    '</Relationships>'
  );
}

/** ชื่อชีต → path ของไฟล์ชีต (ตาม xl/workbook.xml) และรายการ chartsheet */
async function sheetParts(zip: JSZip): Promise<{ byName: Array<{ name: string; path: string }>; chartsheets: number }> {
  const wbXml = await readText(zip, 'xl/workbook.xml');
  if (!wbXml) return { byName: [], chartsheets: 0 };
  const rels = await readRels(zip, 'xl/workbook.xml');
  const sheetsNode = child(parseXml(wbXml), XML_NS.main, 'sheets');
  const byName: Array<{ name: string; path: string }> = [];
  let chartsheets = 0;
  for (const s of childrenOf(sheetsNode, XML_NS.main, 'sheet')) {
    const rel = rels.find((r) => r.id === s.nsAttrs[`${XML_NS.r}|id`]);
    if (!rel) continue;
    if (rel.type === REL_TYPE.chartsheet) chartsheets++;
    else byName.push({ name: s.attrs.name, path: rel.target });
  }
  return { byName, chartsheets };
}

// ---------------------------------------------------------------------------
// เขียน
// ---------------------------------------------------------------------------

/** อ้างช่วงในสูตรของกราฟ: ใส่ '...' รอบชื่อชีตเสมอ (SPEC 5.18) */
export function chartRef(sheetName: string, c1: number, r1: number, c2?: number, r2?: number): string {
  const sheet = `'${sheetName.replace(/'/g, "''")}'`;
  const a = `$${colToStr(c1)}$${r1 + 1}`;
  if (c2 === undefined || r2 === undefined || (c1 === c2 && r1 === r2)) return `${sheet}!${a}`;
  return `${sheet}!${a}:$${colToStr(c2)}$${r2 + 1}`;
}

function strCache(values: string[]): string {
  return (
    `<c:strCache><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => `<c:pt idx="${i}"><c:v>${escapeXml(v)}</c:v></c:pt>`).join('') +
    '</c:strCache>'
  );
}

function numCache(values: (number | null)[]): string {
  return (
    `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>` +
    values.map((v, i) => (v === null ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('') +
    '</c:numCache>'
  );
}

const AX_CAT = 500000001;
const AX_VAL = 500000002;

const AXES_XML =
  `<c:catAx><c:axId val="${AX_CAT}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
  `<c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
  `<c:tickLblPos val="nextTo"/><c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
  `<c:valAx><c:axId val="${AX_VAL}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
  `<c:axPos val="l"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
  `<c:tickLblPos val="nextTo"/><c:crossAx val="${AX_CAT}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;

/**
 * XML ของกราฟหนึ่งกราฟ — ไม่เขียนสี/สไตล์ (spPr, txPr ฯลฯ) ให้โปรแกรมที่เปิดใช้ค่าเริ่มต้นของตัวเอง
 * data ต้องมาจาก chart.range (ไม่ข้ามแถวที่ซ่อน — ตัวกรองไม่ได้อยู่ในไฟล์)
 */
export function buildChartXml(sheetName: string, chart: SheetChart, data: ChartData): string {
  const range = chart.range!;
  const firstDataRow = data.hasHeader ? range.r1 + 1 : range.r1;
  const series = chart.type === 'pie' ? data.series.slice(0, 1) : data.series;

  const serXml = series
    .map((s, i) => {
      const tx = data.hasHeader
        ? `<c:tx><c:strRef><c:f>${escapeXml(chartRef(sheetName, s.col, range.r1))}</c:f>${strCache([s.name])}</c:strRef></c:tx>`
        : '';
      const cat = data.hasCategoryColumn
        ? `<c:cat><c:strRef><c:f>${escapeXml(chartRef(sheetName, range.c1, firstDataRow, range.c1, range.r2))}</c:f>${strCache(data.categories)}</c:strRef></c:cat>`
        : '';
      const val = `<c:val><c:numRef><c:f>${escapeXml(chartRef(sheetName, s.col, firstDataRow, s.col, range.r2))}</c:f>${numCache(s.values)}</c:numRef></c:val>`;
      const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
      if (chart.type === 'bar') return `<c:ser>${head}<c:invertIfNegative val="0"/>${cat}${val}</c:ser>`;
      if (chart.type === 'line') return `<c:ser>${head}${cat}${val}<c:smooth val="0"/></c:ser>`;
      return `<c:ser>${head}${cat}${val}</c:ser>`;
    })
    .join('');

  let plot: string;
  if (chart.type === 'bar') {
    plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXml}<c:gapWidth val="150"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:barChart>${AXES_XML}`;
  } else if (chart.type === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXml}<c:marker val="1"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:lineChart>${AXES_XML}`;
  } else {
    plot = `<c:pieChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/></c:pieChart>`;
  }

  const title = chart.title.trim()
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapeXml(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<c:chartSpace xmlns:c="${XML_NS.c}" xmlns:a="${XML_NS.a}" xmlns:r="${XML_NS.r}">` +
    '<c:roundedCorners val="0"/>' +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>` +
    `<c:legend><c:legendPos val="${chart.type === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend>` +
    '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
    '</c:chartSpace>'
  );
}

function anchorXml(from: CellAnchorEmu, to: CellAnchorEmu, frameId: number, relId: string, name: string): string {
  const pos = (tag: string, p: CellAnchorEmu) =>
    `<xdr:${tag}><xdr:col>${p.col}</xdr:col><xdr:colOff>${p.colOff}</xdr:colOff><xdr:row>${p.row}</xdr:row><xdr:rowOff>${p.rowOff}</xdr:rowOff></xdr:${tag}>`;
  return (
    '<xdr:twoCellAnchor editAs="oneCell">' +
    pos('from', from) +
    pos('to', to) +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${frameId}" name="${escapeXml(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>' +
    `<a:graphic><a:graphicData uri="${XML_NS.c}"><c:chart xmlns:c="${XML_NS.c}" xmlns:r="${XML_NS.r}" r:id="${relId}"/></a:graphicData></a:graphic>` +
    '</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>'
  );
}

// ลำดับ element ใน <worksheet> (CT_Worksheet): <drawing> อยู่ก่อนรายการเหล่านี้
const AFTER_DRAWING = ['legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];

function insertDrawingTag(sheetXml: string, relId: string): string {
  const tag = `<drawing r:id="${relId}"/>`;
  for (const el of AFTER_DRAWING) {
    const idx = sheetXml.search(new RegExp(`<${el}[\\s/>]`));
    if (idx >= 0) return sheetXml.slice(0, idx) + tag + sheetXml.slice(idx);
  }
  return sheetXml.replace('</worksheet>', `${tag}</worksheet>`);
}

function nextRelId(rels: Rel[]): string {
  const nums = rels.map((r) => Number(/^rId(\d+)$/.exec(r.id)?.[1] ?? 0));
  return `rId${Math.max(0, ...nums) + 1}`;
}

/**
 * แทรกกราฟทุกกราฟของสมุดงานลงไฟล์ .xlsx ที่ ExcelJS เขียนแล้ว
 * dataFor: ข้อมูลกราฟจากช่วง (ทุกแถว) — กราฟที่ช่วงถูกลบหมด (range = null) ไม่ถูกเขียน
 */
export async function injectChartsIntoXlsx(
  bytes: Uint8Array,
  sheets: Sheet[],
  dataFor: (sheetIndex: number, chart: SheetChart) => ChartData
): Promise<Uint8Array> {
  const jobs = sheets
    .map((sheet, sheetIndex) => ({ sheet, sheetIndex, charts: (sheet.charts ?? []).filter((c) => c.range) }))
    .filter((j) => j.charts.length > 0);
  if (jobs.length === 0) return bytes;

  const zip = await JSZip.loadAsync(bytes);
  const { byName } = await sheetParts(zip);
  let ct = (await readText(zip, '[Content_Types].xml'))!;
  const addOverride = (part: string, type: string) => {
    if (!ct.includes(`PartName="/${part}"`)) ct = ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  let chartNo = 0;
  let drawingNo = 0;
  while (zip.file(`xl/drawings/drawing${drawingNo + 1}.xml`)) drawingNo++;

  for (const job of jobs) {
    const part = byName.find((p) => p.name === job.sheet.name);
    if (!part) throw new Error(`ไม่พบชีต "${job.sheet.name}" ในไฟล์ที่เขียน`);
    let sheetXml = (await readText(zip, part.path))!;
    if (/<drawing\b/.test(sheetXml)) throw new Error(`ชีต "${job.sheet.name}" มี drawing อยู่แล้ว`);

    drawingNo++;
    const drawingPath = `xl/drawings/drawing${drawingNo}.xml`;
    const drawingRels: Array<{ id: string; type: string; target: string }> = [];
    const anchors: string[] = [];
    job.charts.forEach((chart, i) => {
      chartNo++;
      while (zip.file(`xl/charts/chart${chartNo}.xml`)) chartNo++;
      const chartPath = `xl/charts/chart${chartNo}.xml`;
      zip.file(chartPath, buildChartXml(job.sheet.name, chart, dataFor(job.sheetIndex, chart)));
      addOverride(chartPath, CONTENT_TYPE.chart);
      const relId = `rId${i + 1}`;
      drawingRels.push({ id: relId, type: REL_TYPE.chart, target: `../charts/chart${chartNo}.xml` });
      const { from, to } = chartToTwoCellAnchor(job.sheet, chart);
      anchors.push(anchorXml(from, to, i + 2, relId, chart.title || `Chart ${chartNo}`));
    });
    zip.file(
      drawingPath,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<xdr:wsDr xmlns:xdr="${XML_NS.xdr}" xmlns:a="${XML_NS.a}">${anchors.join('')}</xdr:wsDr>`
    );
    zip.file(relsPathFor(drawingPath), relsXml(drawingRels));
    addOverride(drawingPath, CONTENT_TYPE.drawing);

    // ชีต → drawing
    // rels เดิมที่ ExcelJS เขียน (เช่น hyperlink ภายนอกที่มี TargetMode) ให้เติมต่อท้ายแทนการเขียนใหม่ทั้งไฟล์
    const drawingRelId = nextRelId(await readRels(zip, part.path));
    const drawingTarget = `../drawings/drawing${drawingNo}.xml`;
    const relsPath = relsPathFor(part.path);
    const oldRels = await readText(zip, relsPath);
    zip.file(
      relsPath,
      oldRels
        ? oldRels.replace('</Relationships>', `<Relationship Id="${drawingRelId}" Type="${REL_TYPE.drawing}" Target="${drawingTarget}"/></Relationships>`)
        : relsXml([{ id: drawingRelId, type: REL_TYPE.drawing, target: drawingTarget }])
    );
    if (!/xmlns:r="/.test(sheetXml.slice(0, 2000))) {
      sheetXml = sheetXml.replace(/<worksheet\b/, `<worksheet xmlns:r="${XML_NS.r}"`);
    }
    zip.file(part.path, insertDrawingTag(sheetXml, drawingRelId));
  }

  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// ---------------------------------------------------------------------------
// อ่าน
// ---------------------------------------------------------------------------

/** ส่วนของกราฟที่แอปไม่เก็บ ถ้าพบ = กราฟมี "การจัดรูปแบบ" ที่จะหายเมื่อบันทึก */
const FORMATTING_ELEMENTS = new Set(['spPr', 'txPr', 'dLbls', 'dPt', 'trendline', 'errBars', 'userShapes', 'style', 'pictureOptions', 'dTable', 'upDownBars', 'dropLines', 'hiLowLines', 'view3D']);

interface ParsedRef {
  sheet: string | null;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** อ่าน 'ชื่อชีต'!$A$1:$B$5 หรือ ชีต!A1 */
export function parseChartRef(f: string): ParsedRef | null {
  const text = f.trim();
  let sheet: string | null = null;
  let rest = text;
  if (text.startsWith("'")) {
    let i = 1;
    let name = '';
    while (i < text.length) {
      if (text[i] === "'") {
        if (text[i + 1] === "'") {
          name += "'";
          i += 2;
          continue;
        }
        break;
      }
      name += text[i++];
    }
    if (text[i + 1] !== '!') return null;
    sheet = name;
    rest = text.slice(i + 2);
  } else {
    const bang = text.lastIndexOf('!');
    if (bang >= 0) {
      sheet = text.slice(0, bang);
      rest = text.slice(bang + 1);
    }
  }
  rest = rest.replace(/\$/g, '');
  if (rest.includes(':')) {
    const r = parseRange(rest);
    return r ? { sheet, r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2), r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2) } : null;
  }
  const p = parseA1(rest);
  return p ? { sheet, r1: p.r, c1: p.c, r2: p.r, c2: p.c } : null;
}

function refOf(node: XmlNode | undefined): ParsedRef | null | undefined {
  if (!node) return undefined;
  const refNode = child(node, XML_NS.c, 'strRef') ?? child(node, XML_NS.c, 'numRef');
  const f = child(refNode, XML_NS.c, 'f');
  return f ? parseChartRef(textOf(f)) : null;
}

function chartTitle(chartNode: XmlNode | undefined): string {
  const title = child(chartNode, XML_NS.c, 'title');
  if (!title) return '';
  const tx = child(title, XML_NS.c, 'tx');
  const rich = child(tx, XML_NS.c, 'rich');
  if (rich) {
    return childrenOf(rich, XML_NS.a, 'p')
      .map((p) => descendants(p, XML_NS.a, 't').map(textOf).join(''))
      .join(' ')
      .trim();
  }
  // ชื่อจากเซล: ใช้ค่า cache
  return descendants(title, XML_NS.c, 'v').map(textOf).join('').trim();
}

/**
 * แปลง XML กราฟเป็นชนิด/ชื่อ/ช่วง ถ้าเข้ารูปแบบที่แอปรองรับ (SPEC 5.18) มิฉะนั้นคืน null
 */
export function parseChartXml(xml: string, hostSheetName: string): { type: ChartType; title: string; range: NonNullable<SheetChart['range']>; formatting: boolean } | null {
  const space = parseXml(xml);
  const chartNode = child(space, XML_NS.c, 'chart');
  const plotArea = child(chartNode, XML_NS.c, 'plotArea');
  if (!plotArea) return null;
  const AXES = new Set(['layout', 'catAx', 'valAx', 'dateAx', 'serAx', 'spPr', 'extLst']);
  const plots = plotArea.children.filter((c) => c.uri === XML_NS.c && !AXES.has(c.local));
  if (plots.length !== 1) return null;
  const plot = plots[0];
  let type: ChartType;
  if (plot.local === 'barChart') {
    const dir = child(plot, XML_NS.c, 'barDir')?.attrs.val;
    if (dir !== 'col') return null; // แท่งแนวนอนไม่รองรับ
    const grouping = child(plot, XML_NS.c, 'grouping')?.attrs.val ?? 'clustered';
    if (grouping !== 'clustered') return null;
    type = 'bar';
  } else if (plot.local === 'lineChart') {
    const grouping = child(plot, XML_NS.c, 'grouping')?.attrs.val ?? 'standard';
    if (grouping !== 'standard') return null;
    type = 'line';
  } else if (plot.local === 'pieChart') {
    type = 'pie';
  } else {
    return null;
  }

  let sers = childrenOf(plot, XML_NS.c, 'ser');
  if (sers.length === 0) return null;
  if (type === 'pie') sers = sers.slice(0, 1);

  const onHost = (r: ParsedRef | null | undefined) => !r || r.sheet === null || sameSheetName(r.sheet, hostSheetName);
  const vals = sers.map((s) => refOf(child(s, XML_NS.c, 'val')));
  const cats = sers.map((s) => refOf(child(s, XML_NS.c, 'cat')));
  const names = sers.map((s) => refOf(child(s, XML_NS.c, 'tx')));
  if (vals.some((v) => !v)) return null;
  if ([...vals, ...cats, ...names].some((r) => r === null || !onHost(r))) return null;

  const v0 = vals[0]!;
  // ชุดข้อมูลแต่ละชุดเป็นคอลัมน์เดียว แถวเดียวกัน และเรียงคอลัมน์ติดกัน
  for (let i = 0; i < vals.length; i++) {
    const v = vals[i]!;
    if (v.c1 !== v.c2 || v.r1 !== v0.r1 || v.r2 !== v0.r2 || v.c1 !== v0.c1 + i) return null;
  }
  const cat = cats[0];
  if (cats.some((c) => (c ? !cat || c.c1 !== cat.c1 || c.r1 !== cat.r1 || c.r2 !== cat.r2 : !!cat))) return null;
  if (cat && (cat.c1 !== cat.c2 || cat.r1 !== v0.r1 || cat.r2 !== v0.r2 || cat.c1 !== v0.c1 - 1)) return null;
  if (!cat && vals.length > 1) return null;

  // ชื่อชุดข้อมูล: ทุกชุดอยู่แถวเหนือข้อมูลในคอลัมน์ของตัวเอง หรือไม่มีเลย
  const withName = names.filter(Boolean).length;
  if (withName !== 0 && withName !== names.length) return null;
  if (withName > 0 && names.some((n, i) => n!.r1 !== v0.r1 - 1 || n!.r2 !== n!.r1 || n!.c1 !== vals[i]!.c1 || n!.c2 !== n!.c1)) return null;

  const range = {
    r1: withName > 0 ? v0.r1 - 1 : v0.r1,
    c1: cat ? cat.c1 : v0.c1,
    r2: v0.r2,
    c2: vals[vals.length - 1]!.c1,
  };
  return { type, title: chartTitle(chartNode), range, formatting: hasFormatting(space) };
}

function hasFormatting(node: XmlNode): boolean {
  for (const c of node.children) {
    if (c.uri === XML_NS.c && FORMATTING_ELEMENTS.has(c.local)) return true;
    if (c.uri === XML_NS.c && c.local === 'rich') continue; // ข้อความชื่อกราฟ ไม่นับ
    if (hasFormatting(c)) return true;
  }
  return false;
}

function cellAnchor(node: XmlNode | undefined): CellAnchorEmu | null {
  if (!node) return null;
  const n = (local: string) => Number(textOf(child(node, XML_NS.xdr, local)) || 0);
  return { col: n('col'), colOff: n('colOff'), row: n('row'), rowOff: n('rowOff') };
}

export interface ReadChartsResult {
  /** กราฟที่อ่านได้ ตามลำดับชีตของ sheets ที่ส่งเข้ามา */
  chartsBySheet: SheetChart[][];
  /** มีกราฟที่แอปไม่รองรับ (จะหายเมื่อบันทึก) */
  unsupported: boolean;
  /** มีกราฟที่อ่านได้แต่มีการจัดรูปแบบที่แอปไม่เก็บ */
  formatting: boolean;
}

/** อ่านกราฟจากไฟล์ .xlsx (sheets: ชีตที่อ่านจากไฟล์เดียวกันแล้ว ใช้ชื่อและความกว้าง/สูง) */
export async function readChartsFromXlsx(bytes: Uint8Array, sheets: Sheet[]): Promise<ReadChartsResult> {
  const result: ReadChartsResult = { chartsBySheet: sheets.map(() => []), unsupported: false, formatting: false };
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return result;
  }
  const allCharts = Object.keys(zip.files).filter((n) => /^xl\/charts\/chart[^/]*\.xml$/.test(n));
  const { byName, chartsheets } = await sheetParts(zip);
  if (chartsheets > 0) result.unsupported = true;
  const used = new Set<string>();

  for (const part of byName) {
    const sheetIndex = sheets.findIndex((s) => s.name === part.name);
    if (sheetIndex < 0) continue;
    const sheet = sheets[sheetIndex];
    const sheetRels = await readRels(zip, part.path);
    for (const drawingRel of sheetRels.filter((r) => r.type === REL_TYPE.drawing)) {
      const drawingXml = await readText(zip, drawingRel.target);
      if (!drawingXml) continue;
      const drawingRels = await readRels(zip, drawingRel.target);
      let root: XmlNode;
      try {
        root = parseXml(drawingXml);
      } catch {
        continue;
      }
      for (const anchorNode of root.children) {
        const chartRefNode = descendants(anchorNode, XML_NS.c, 'chart')[0];
        if (!chartRefNode) continue; // รูปภาพ/รูปร่าง: ไม่ใช่กราฟ
        const rel = drawingRels.find((r) => r.id === chartRefNode.nsAttrs[`${XML_NS.r}|id`]);
        if (!rel) continue;
        used.add(rel.target);
        const xml = await readText(zip, rel.target);
        let parsed: ReturnType<typeof parseChartXml> = null;
        try {
          parsed = xml ? parseChartXml(xml, sheet.name) : null;
        } catch {
          parsed = null;
        }
        let place: Pick<SheetChart, 'anchor' | 'width' | 'height'> | null = null;
        if (anchorNode.local === 'twoCellAnchor') {
          const from = cellAnchor(child(anchorNode, XML_NS.xdr, 'from'));
          const to = cellAnchor(child(anchorNode, XML_NS.xdr, 'to'));
          if (from && to) place = twoCellAnchorToChart(sheet, from, to);
        } else if (anchorNode.local === 'oneCellAnchor') {
          const from = cellAnchor(child(anchorNode, XML_NS.xdr, 'from'));
          const ext = child(anchorNode, XML_NS.xdr, 'ext');
          if (from && ext) {
            place = {
              anchor: { row: from.row, col: from.col, dx: Math.round(from.colOff / EMU_PER_PX), dy: Math.round(from.rowOff / EMU_PER_PX) },
              width: Math.max(MIN_CHART_WIDTH, Math.round(Number(ext.attrs.cx) / EMU_PER_PX)),
              height: Math.max(MIN_CHART_HEIGHT, Math.round(Number(ext.attrs.cy) / EMU_PER_PX)),
            };
          }
        } else if (anchorNode.local === 'absoluteAnchor') {
          const pos = child(anchorNode, XML_NS.xdr, 'pos');
          const ext = child(anchorNode, XML_NS.xdr, 'ext');
          if (pos && ext) {
            place = {
              anchor: pxToAnchor(sheet, Number(pos.attrs.x) / EMU_PER_PX, Number(pos.attrs.y) / EMU_PER_PX),
              width: Math.max(MIN_CHART_WIDTH, Math.round(Number(ext.attrs.cx) / EMU_PER_PX)),
              height: Math.max(MIN_CHART_HEIGHT, Math.round(Number(ext.attrs.cy) / EMU_PER_PX)),
            };
          }
        }
        if (!parsed || !place) {
          result.unsupported = true;
          continue;
        }
        if (parsed.formatting) result.formatting = true;
        result.chartsBySheet[sheetIndex].push({ id: newChartId(), type: parsed.type, title: parsed.title, range: parsed.range, ...place });
      }
    }
  }
  // กราฟที่ไม่ได้ผูกกับชีตที่แอปอ่าน (เช่น อยู่ในชีตกราฟแยก)
  if (allCharts.some((c) => !used.has(c))) result.unsupported = true;
  return result;
}
