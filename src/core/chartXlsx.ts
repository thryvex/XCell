/**
 * กราฟใน .xlsx (5D): ExcelJS 4.4 เขียนกราฟไม่ได้และทิ้งกราฟตอนอ่าน
 * จึงแทรกส่วนกราฟลง zip เองหลัง ExcelJS เขียนไฟล์ และอ่านกราฟจาก zip เองตอนเปิด
 * (วิธีเดียวกับ spike ใน scripts/chart-spike/)
 */
import JSZip from 'jszip';
import type { Sheet, SheetChart, ChartSeries, ChartType, ChartLegend } from './model';
import {
  ChartContext,
  CellAnchorEmu,
  EMU_PER_PX,
  MIN_CHART_WIDTH,
  MIN_CHART_HEIGHT,
  chartToTwoCellAnchor,
  twoCellAnchorToChart,
  pxToAnchor,
  newChartId,
  parseChartRef,
  makeChartRef,
  refCells,
  pointValue,
  seriesName,
  paletteColor,
  ParsedChartRef,
  MAX_CHART_SERIES,
} from './chart';
import { cellToA1, colToStr } from './address';
import { storedColWidthPx, storedRowHeightPx } from './sizeUnits';
import { XmlNode, XML_NS, parseXml, child, childrenOf, descendants, textOf, escapeXml, MAX_XML_LENGTH } from './xmlTree';

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
    target: r.attrs.TargetMode === 'External' ? r.attrs.Target : resolveTarget(partPath, r.attrs.Target),
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

/** ชื่อชีต → path ของไฟล์ชีต (ตาม xl/workbook.xml) และจำนวน chartsheet */
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

/** reference ใน chart XML: ใส่ '...' รอบชื่อชีตเสมอ และใช้ $ (SPEC 5.18) */
export function xlsxChartRef(r: ParsedChartRef): string {
  const sheet = `'${r.sheet.replace(/'/g, "''")}'`;
  const a = `$${colToStr(r.c1)}$${r.r1 + 1}`;
  if (r.r1 === r.r2 && r.c1 === r.c2) return `${sheet}!${a}`;
  return `${sheet}!${a}:$${colToStr(r.c2)}$${r.r2 + 1}`;
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

function hex(color: string): string {
  const m = /^#?([0-9a-fA-F]{6})$/.exec(color.trim());
  return (m ? m[1] : paletteColor(0).slice(1)).toUpperCase();
}

const solid = (color: string) => `<a:solidFill><a:srgbClr val="${hex(color)}"/></a:solidFill>`;

const DATA_LABELS =
  '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>';

function richText(text: string, vertical = false): string {
  const bodyPr = vertical ? '<a:bodyPr rot="-5400000" vert="horz"/>' : '<a:bodyPr/>';
  return `<c:tx><c:rich>${bodyPr}<a:lstStyle/><a:p><a:r><a:t>${escapeXml(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

function titleXml(text: string, vertical = false): string {
  return text.trim() ? `<c:title>${richText(text, vertical)}<c:overlay val="0"/></c:title>` : '';
}

const AX_CAT = 500000001;
const AX_VAL = 500000002;

function axesXml(chart: SheetChart): string {
  return (
    `<c:catAx><c:axId val="${AX_CAT}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="b"/>${titleXml(chart.xTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="nextTo"/><c:crossAx val="${AX_VAL}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
    `<c:valAx><c:axId val="${AX_VAL}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/>` +
    `<c:axPos val="l"/><c:majorGridlines/>${titleXml(chart.yTitle, true)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/>` +
    `<c:tickLblPos val="nextTo"/><c:crossAx val="${AX_CAT}"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`
  );
}

/** ข้อมูลของชุดหนึ่งสำหรับเขียนลงไฟล์ (ทุกเซลในช่วง ไม่ข้ามแถวที่ซ่อน — ไฟล์ใช้ plotVisOnly เอง) */
interface SeriesForFile {
  series: ChartSeries;
  index: number;
  name: ParsedChartRef | null;
  nameText: string;
  categories: ParsedChartRef | null;
  categoryLabels: string[];
  values: ParsedChartRef;
  numbers: (number | null)[];
}

function seriesForFile(chart: SheetChart, ctx: ChartContext): SeriesForFile[] {
  const out: SeriesForFile[] = [];
  chart.series.forEach((series, index) => {
    const values = parseChartRef(series.values);
    const vSheet = values ? ctx.sheetIndexOf(values.sheet) : -1;
    if (!values || vSheet < 0) return; // reference หายแล้ว: ไม่เขียนชุดนี้
    const categories = series.categories ? parseChartRef(series.categories) : null;
    const cSheet = categories ? ctx.sheetIndexOf(categories.sheet) : -1;
    const name = 'ref' in series.name ? parseChartRef(series.name.ref) : null;
    out.push({
      series,
      index,
      name: name && ctx.sheetIndexOf(name.sheet) >= 0 ? name : null,
      nameText: seriesName(series, index, ctx),
      categories: categories && cSheet >= 0 ? categories : null,
      categoryLabels: categories && cSheet >= 0 ? refCells(categories).map((c) => ctx.label(cSheet, cellToA1(c.r, c.c))) : [],
      values,
      numbers: refCells(values).map((c) => pointValue(ctx.value(vSheet, cellToA1(c.r, c.c))).value),
    });
  });
  return chart.type === 'pie' ? out.slice(0, 1) : out;
}

/**
 * XML ของกราฟหนึ่งกราฟ หรือ null ถ้าไม่มีชุดข้อมูลที่ reference ยังอยู่
 * สีชุดข้อมูลเขียนชัด (srgbClr) ไม่พึ่งธีมของ Office; ไม่เขียนฟอนต์/สไตล์อื่น
 */
export function buildChartXml(chart: SheetChart, ctx: ChartContext): string | null {
  const list = seriesForFile(chart, ctx);
  if (list.length === 0) return null;

  const serXml = list
    .map((s, i) => {
      let tx = '';
      if (s.name) tx = `<c:tx><c:strRef><c:f>${escapeXml(xlsxChartRef(s.name))}</c:f>${strCache([s.nameText])}</c:strRef></c:tx>`;
      else if ('text' in s.series.name && s.series.name.text.trim()) tx = `<c:tx><c:v>${escapeXml(s.series.name.text)}</c:v></c:tx>`;
      const cat = s.categories
        ? `<c:cat><c:strRef><c:f>${escapeXml(xlsxChartRef(s.categories))}</c:f>${strCache(s.categoryLabels)}</c:strRef></c:cat>`
        : '';
      const val = `<c:val><c:numRef><c:f>${escapeXml(xlsxChartRef(s.values))}</c:f>${numCache(s.numbers)}</c:numRef></c:val>`;
      const head = `<c:idx val="${i}"/><c:order val="${i}"/>${tx}`;
      const labels = chart.dataLabels ? DATA_LABELS : '';
      if (chart.type === 'bar') {
        return `<c:ser>${head}<c:spPr>${solid(s.series.color)}</c:spPr><c:invertIfNegative val="0"/>${labels}${cat}${val}</c:ser>`;
      }
      if (chart.type === 'line') {
        const ln = `<a:ln w="28575" cap="rnd">${solid(s.series.color)}<a:round/></a:ln>`;
        const marker = `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr>${solid(s.series.color)}<a:ln>${solid(s.series.color)}</a:ln></c:spPr></c:marker>`;
        return `<c:ser>${head}<c:spPr>${ln}</c:spPr>${marker}${labels}${cat}${val}<c:smooth val="0"/></c:ser>`;
      }
      // วงกลม: สีแต่ละชิ้นจากจานสี
      const dPts = s.numbers
        .map((_, p) => `<c:dPt><c:idx val="${p}"/><c:bubble3D val="0"/><c:spPr>${solid(paletteColor(p))}</c:spPr></c:dPt>`)
        .join('');
      return `<c:ser>${head}${dPts}${labels}${cat}${val}</c:ser>`;
    })
    .join('');

  let plot: string;
  if (chart.type === 'bar') {
    plot = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXml}<c:gapWidth val="150"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:barChart>${axesXml(chart)}`;
  } else if (chart.type === 'line') {
    plot = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXml}<c:marker val="1"/><c:axId val="${AX_CAT}"/><c:axId val="${AX_VAL}"/></c:lineChart>${axesXml(chart)}`;
  } else {
    plot = `<c:pieChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/></c:pieChart>`;
  }

  const title = chart.title.trim() ? `${titleXml(chart.title)}<c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend =
    chart.legend === 'none' ? '' : `<c:legend><c:legendPos val="${chart.legend === 'right' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend>`;

  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<c:chartSpace xmlns:c="${XML_NS.c}" xmlns:a="${XML_NS.a}" xmlns:r="${XML_NS.r}">` +
    '<c:roundedCorners val="0"/>' +
    `<c:chart>${title}<c:plotArea><c:layout/>${plot}</c:plotArea>${legend}` +
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
 * ctx: อ่านค่าปัจจุบันสำหรับค่า cache; กราฟที่ reference หายทุกชุดไม่ถูกเขียน
 */
export async function injectChartsIntoXlsx(bytes: Uint8Array, sheets: Sheet[], ctx: ChartContext): Promise<Uint8Array> {
  const jobs = sheets
    .map((sheet) => ({
      sheet,
      charts: (sheet.charts ?? []).map((chart) => ({ chart, xml: buildChartXml(chart, ctx) })).filter((c) => c.xml !== null),
    }))
    .filter((j) => j.charts.length > 0);
  if (jobs.length === 0) return bytes;

  const zip = await JSZip.loadAsync(bytes);
  const { byName } = await sheetParts(zip);
  let ct = (await readText(zip, '[Content_Types].xml'))!;
  const addOverride = (part: string, type: string) => {
    if (!ct.includes(`PartName="/${part}"`)) ct = ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  // ชื่อไฟล์ในแพ็กเกจไม่ชนกับที่มีอยู่
  let chartNo = 0;
  const nextChartPath = () => {
    do chartNo++;
    while (zip.file(`xl/charts/chart${chartNo}.xml`));
    return `xl/charts/chart${chartNo}.xml`;
  };
  let drawingNo = 0;
  const nextDrawingPath = () => {
    do drawingNo++;
    while (zip.file(`xl/drawings/drawing${drawingNo}.xml`));
    return `xl/drawings/drawing${drawingNo}.xml`;
  };

  for (const job of jobs) {
    const part = byName.find((p) => p.name === job.sheet.name);
    if (!part) throw new Error(`ไม่พบชีต "${job.sheet.name}" ในไฟล์ที่เขียน`);
    let sheetXml = (await readText(zip, part.path))!;
    if (/<drawing\b/.test(sheetXml)) throw new Error(`ชีต "${job.sheet.name}" มี drawing อยู่แล้ว`);

    // ตำแหน่งในไฟล์คิดจากขนาดคอลัมน์/แถวตามที่ถูกเขียนลงไฟล์จริง (ปัดหน่วยแล้ว)
    const storedSizes: Sheet = {
      ...job.sheet,
      colWidths: Object.fromEntries(Object.entries(job.sheet.colWidths).map(([k, v]) => [k, storedColWidthPx(v)])),
      rowHeights: Object.fromEntries(Object.entries(job.sheet.rowHeights).map(([k, v]) => [k, storedRowHeightPx(v)])),
    };
    const drawingPath = nextDrawingPath();
    const drawingRels: Array<{ id: string; type: string; target: string }> = [];
    const anchors: string[] = [];
    job.charts.forEach(({ chart, xml }, i) => {
      const chartPath = nextChartPath();
      zip.file(chartPath, xml!);
      addOverride(chartPath, CONTENT_TYPE.chart);
      const relId = `rId${i + 1}`;
      drawingRels.push({ id: relId, type: REL_TYPE.chart, target: `../charts/${chartPath.split('/').pop()}` });
      const { from, to } = chartToTwoCellAnchor(storedSizes, chart);
      anchors.push(anchorXml(from, to, i + 2, relId, chart.title || `Chart ${i + 1}`));
    });
    zip.file(
      drawingPath,
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
        `<xdr:wsDr xmlns:xdr="${XML_NS.xdr}" xmlns:a="${XML_NS.a}">${anchors.join('')}</xdr:wsDr>`
    );
    zip.file(relsPathFor(drawingPath), relsXml(drawingRels));
    addOverride(drawingPath, CONTENT_TYPE.drawing);

    // ชีต → drawing: rels เดิมที่ ExcelJS เขียน (เช่น hyperlink ภายนอก) ให้เติมต่อท้ายแทนการเขียนใหม่ทั้งไฟล์
    const drawingRelId = nextRelId(await readRels(zip, part.path));
    const drawingTarget = `../drawings/${drawingPath.split('/').pop()}`;
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

/** ส่วนที่ทำให้กราฟเกินขอบเขต 5.18 (ไม่นำเข้า) */
const OUT_OF_SCOPE_IN_SERIES = new Set(['trendline', 'errBars']);

export type ParsedChart = Omit<SheetChart, 'id' | 'anchor' | 'width' | 'height'>;

/** reference ของกราฟจากไฟล์ → reference ในโมเดล (ไม่มีชื่อชีต = ชีตที่กราฟอยู่); undefined = ไม่มี, null = ใช้ไม่ได้ */
function refOf(node: XmlNode | undefined, hostSheet: string): ParsedChartRef | null | undefined {
  if (!node) return undefined;
  const refNode = child(node, XML_NS.c, 'strRef') ?? child(node, XML_NS.c, 'numRef');
  const f = child(refNode, XML_NS.c, 'f');
  return f ? parseChartRef(textOf(f), hostSheet) : null;
}

function isLine(r: ParsedChartRef): boolean {
  return r.r1 === r.r2 || r.c1 === r.c2;
}

function richOrCachedText(title: XmlNode | undefined): string {
  if (!title) return '';
  const rich = child(child(title, XML_NS.c, 'tx'), XML_NS.c, 'rich');
  if (rich) {
    return childrenOf(rich, XML_NS.a, 'p')
      .map((p) => descendants(p, XML_NS.a, 't').map(textOf).join(''))
      .join(' ')
      .trim();
  }
  // ชื่อจากเซล: ใช้ค่า cache (ในแอปเป็นข้อความคงที่)
  return descendants(title, XML_NS.c, 'v').map(textOf).join('').trim();
}

/** สีชุดข้อมูลจาก spPr (srgbClr); สีจากธีมของ Office (schemeClr) อ่านไม่ได้ → จานสีของแอป */
function seriesColor(ser: XmlNode, type: ChartType, i: number): string {
  const spPr = child(ser, XML_NS.c, 'spPr');
  const fillOwner = type === 'line' ? child(spPr, XML_NS.a, 'ln') : spPr;
  const rgb = child(child(fillOwner, XML_NS.a, 'solidFill'), XML_NS.a, 'srgbClr')?.attrs.val;
  return rgb && /^[0-9a-fA-F]{6}$/.test(rgb) ? `#${rgb.toUpperCase()}` : paletteColor(i);
}

function showsValues(dLbls: XmlNode | undefined): boolean {
  return child(dLbls, XML_NS.c, 'showVal')?.attrs.val === '1' && child(dLbls, XML_NS.c, 'delete')?.attrs.val !== '1';
}

/**
 * แปลง XML กราฟเป็นกราฟในแอป ถ้าเข้าขอบเขต 5.18 มิฉะนั้นคืน null
 * hostSheetName: ชีตที่กราฟอยู่ (ใช้กับ reference ที่ไม่มีชื่อชีต)
 */
export function parseChartXml(xml: string, hostSheetName: string): ParsedChart | null {
  const space = parseXml(xml);
  const chartNode = child(space, XML_NS.c, 'chart');
  const plotArea = child(chartNode, XML_NS.c, 'plotArea');
  if (!plotArea) return null;
  const NOT_PLOT = new Set(['layout', 'catAx', 'valAx', 'dateAx', 'serAx', 'spPr', 'dTable', 'extLst']);
  const plots = plotArea.children.filter((c) => c.uri === XML_NS.c && !NOT_PLOT.has(c.local));
  if (plots.length !== 1) return null; // combo
  if (childrenOf(plotArea, XML_NS.c, 'valAx').length > 1) return null; // แกนรอง
  if (childrenOf(plotArea, XML_NS.c, 'serAx').length > 0 || child(chartNode, XML_NS.c, 'view3D')) return null;
  const plot = plots[0];

  let type: ChartType;
  if (plot.local === 'barChart') {
    if (child(plot, XML_NS.c, 'barDir')?.attrs.val !== 'col') return null; // แท่งแนวนอน
    if ((child(plot, XML_NS.c, 'grouping')?.attrs.val ?? 'clustered') !== 'clustered') return null;
    type = 'bar';
  } else if (plot.local === 'lineChart') {
    if ((child(plot, XML_NS.c, 'grouping')?.attrs.val ?? 'standard') !== 'standard') return null;
    type = 'line';
  } else if (plot.local === 'pieChart') {
    type = 'pie';
  } else {
    return null;
  }

  let sers = childrenOf(plot, XML_NS.c, 'ser');
  if (sers.length === 0) return null;
  if (sers.some((s) => s.children.some((c) => c.uri === XML_NS.c && OUT_OF_SCOPE_IN_SERIES.has(c.local)))) return null;
  if (type === 'pie') sers = sers.slice(0, 1);
  if (sers.length > MAX_CHART_SERIES) return null;

  const series: ChartSeries[] = [];
  for (const [i, ser] of sers.entries()) {
    const values = refOf(child(ser, XML_NS.c, 'val'), hostSheetName);
    if (!values || !isLine(values)) return null; // ค่าต้องอ้างช่วงเซลแถวเดียว/คอลัมน์เดียว
    const catNode = child(ser, XML_NS.c, 'cat');
    const categories = refOf(catNode, hostSheetName);
    if (catNode && (!categories || !isLine(categories))) return null; // ป้ายแบบหลายระดับหรือค่าคงที่
    const txNode = child(ser, XML_NS.c, 'tx');
    let name: ChartSeries['name'] = { text: '' };
    if (child(txNode, XML_NS.c, 'strRef')) {
      const nameRef = refOf(txNode, hostSheetName);
      if (!nameRef) return null;
      name = { ref: makeChartRef(nameRef.sheet, nameRef.r1, nameRef.c1) };
    } else if (txNode) {
      name = { text: textOf(child(txNode, XML_NS.c, 'v')) };
    }
    const s: ChartSeries = {
      name,
      values: makeChartRef(values.sheet, values.r1, values.c1, values.r2, values.c2),
      color: seriesColor(ser, type, i),
    };
    if (categories) s.categories = makeChartRef(categories.sheet, categories.r1, categories.c1, categories.r2, categories.c2);
    series.push(s);
  }

  const legendNode = child(chartNode, XML_NS.c, 'legend');
  let legend: ChartLegend = 'none';
  if (legendNode) {
    const pos = child(legendNode, XML_NS.c, 'legendPos')?.attrs.val ?? 'r';
    legend = pos === 'b' || pos === 't' ? 'bottom' : 'right';
  }
  const catAx = child(plotArea, XML_NS.c, 'catAx') ?? child(plotArea, XML_NS.c, 'dateAx');
  const valAx = child(plotArea, XML_NS.c, 'valAx');
  const dataLabels =
    showsValues(child(plot, XML_NS.c, 'dLbls')) || sers.some((s) => showsValues(child(s, XML_NS.c, 'dLbls')));

  return {
    type,
    title: richOrCachedText(child(chartNode, XML_NS.c, 'title')),
    series,
    legend,
    xTitle: type === 'pie' ? '' : richOrCachedText(child(catAx, XML_NS.c, 'title')),
    yTitle: type === 'pie' ? '' : richOrCachedText(child(valAx, XML_NS.c, 'title')),
    dataLabels,
  };
}

function cellAnchor(node: XmlNode | undefined): CellAnchorEmu | null {
  if (!node) return null;
  const n = (local: string) => Number(textOf(child(node, XML_NS.xdr, local)) || 0);
  return { col: n('col'), colOff: n('colOff'), row: n('row'), rowOff: n('rowOff') };
}

function placement(sheet: Sheet, anchorNode: XmlNode): Pick<SheetChart, 'anchor' | 'width' | 'height'> | null {
  const size = (ext: XmlNode) => ({
    width: Math.max(MIN_CHART_WIDTH, Math.round(Number(ext.attrs.cx) / EMU_PER_PX) || 0),
    height: Math.max(MIN_CHART_HEIGHT, Math.round(Number(ext.attrs.cy) / EMU_PER_PX) || 0),
  });
  if (anchorNode.local === 'twoCellAnchor') {
    const from = cellAnchor(child(anchorNode, XML_NS.xdr, 'from'));
    const to = cellAnchor(child(anchorNode, XML_NS.xdr, 'to'));
    return from && to ? twoCellAnchorToChart(sheet, from, to) : null;
  }
  if (anchorNode.local === 'oneCellAnchor') {
    const from = cellAnchor(child(anchorNode, XML_NS.xdr, 'from'));
    const ext = child(anchorNode, XML_NS.xdr, 'ext');
    if (!from || !ext) return null;
    return {
      anchor: { row: from.row, col: from.col, dx: Math.round(from.colOff / EMU_PER_PX), dy: Math.round(from.rowOff / EMU_PER_PX) },
      ...size(ext),
    };
  }
  if (anchorNode.local === 'absoluteAnchor') {
    const pos = child(anchorNode, XML_NS.xdr, 'pos');
    const ext = child(anchorNode, XML_NS.xdr, 'ext');
    if (!pos || !ext) return null;
    return { anchor: pxToAnchor(sheet, Number(pos.attrs.x) / EMU_PER_PX, Number(pos.attrs.y) / EMU_PER_PX), ...size(ext) };
  }
  return null;
}

export interface ReadChartsResult {
  /** กราฟที่อ่านได้ ตามลำดับชีตของ sheets ที่ส่งเข้ามา */
  chartsBySheet: SheetChart[][];
  /** มีกราฟที่นำเข้าไม่ได้ (จะหายเมื่อบันทึก) */
  unsupported: boolean;
}

/** ไฟล์ส่วนกราฟใหญ่เกินนี้ (ไบต์ หลังแตก zip) ไม่อ่าน */
const MAX_PART_BYTES = MAX_XML_LENGTH;

async function readPart(zip: JSZip, path: string): Promise<string | undefined> {
  const file = zip.file(path);
  if (!file) return undefined;
  const size = (file as unknown as { _data?: { uncompressedSize?: number } })._data?.uncompressedSize;
  if (size !== undefined && size > MAX_PART_BYTES) return undefined;
  const text = await file.async('string');
  return text.length > MAX_PART_BYTES ? undefined : text;
}

/** อ่านกราฟจากไฟล์ .xlsx (sheets: ชีตที่อ่านจากไฟล์เดียวกันแล้ว ใช้ชื่อและความกว้าง/สูง) */
export async function readChartsFromXlsx(bytes: Uint8Array, sheets: Sheet[]): Promise<ReadChartsResult> {
  const result: ReadChartsResult = { chartsBySheet: sheets.map(() => []), unsupported: false };
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(bytes);
  } catch {
    return result;
  }
  // ใช้รายชื่อไฟล์ใน zip ตรวจว่ามีกราฟ (ExcelJS ทิ้งกราฟตั้งแต่ตอนอ่าน)
  const allCharts = Object.keys(zip.files).filter((n) => /^xl\/charts\/chart[^/]*\.xml$/.test(n));
  let parts: Awaited<ReturnType<typeof sheetParts>>;
  try {
    parts = await sheetParts(zip);
  } catch {
    return { ...result, unsupported: allCharts.length > 0 };
  }
  if (parts.chartsheets > 0) result.unsupported = true;
  const used = new Set<string>();

  for (const part of parts.byName) {
    const sheetIndex = sheets.findIndex((s) => s.name === part.name);
    if (sheetIndex < 0) continue;
    const sheet = sheets[sheetIndex];
    let sheetRels: Rel[];
    try {
      sheetRels = await readRels(zip, part.path);
    } catch {
      continue;
    }
    for (const drawingRel of sheetRels.filter((r) => r.type === REL_TYPE.drawing)) {
      let root: XmlNode;
      let drawingRels: Rel[];
      try {
        const drawingXml = await readPart(zip, drawingRel.target);
        if (!drawingXml) continue;
        root = parseXml(drawingXml);
        drawingRels = await readRels(zip, drawingRel.target);
      } catch {
        continue;
      }
      for (const anchorNode of root.children) {
        const chartRefNode = descendants(anchorNode, XML_NS.c, 'chart')[0];
        if (!chartRefNode) continue; // รูปภาพ/รูปร่าง: ไม่ใช่กราฟ
        const rel = drawingRels.find((r) => r.id === chartRefNode.nsAttrs[`${XML_NS.r}|id`]);
        if (!rel) continue;
        used.add(rel.target);
        let parsed: ParsedChart | null = null;
        try {
          const xml = await readPart(zip, rel.target);
          parsed = xml ? parseChartXml(xml, sheet.name) : null;
        } catch {
          parsed = null;
        }
        const place = placement(sheet, anchorNode);
        if (!parsed || !place) {
          result.unsupported = true;
          continue;
        }
        result.chartsBySheet[sheetIndex].push({ id: newChartId(), ...parsed, ...place });
      }
    }
  }
  // กราฟที่ไม่ได้ผูกกับชีตที่แอปอ่าน (เช่น อยู่ในชีตกราฟแยก)
  if (allCharts.some((c) => !used.has(c))) result.unsupported = true;
  return result;
}
