import type { Sheet, SheetChart, ChartSeries, ChartType, CellValue, NumFmt, Workbook } from './model';
import { cellToA1, parseA1, parseRange, rangeToA1 } from './address';
import { tokenize } from './tokenizer';
import { formatSheetReference, sameSheetName } from './sheetName';
import { shiftFormulaOnInsert, shiftFormulaOnDelete, renameSheetInFormula, invalidateSheetInFormula } from './refshift';
import { isCellError } from './functions';
import { th } from '../i18n/th';
import { Evaluator } from './evaluator';
import { formatDate, formatDisplayValue } from './numfmt';

/** ขนาดเริ่มต้นของคอลัมน์/แถวในตาราง (px ที่ซูม 100%) — Grid ใช้ค่าชุดเดียวกัน */
export const DEFAULT_COL_WIDTH_PX = 110;
export const DEFAULT_ROW_HEIGHT_PX = 28;

/** ขีดจำกัด (SPEC 5.18) */
export const MAX_CHART_SERIES = 10;
export const MAX_CHART_POINTS = 1000;
/** ชุดเดียวไม่เกินเท่านี้จุด → แนะนำกราฟวงกลม */
export const PIE_SUGGEST_MAX_POINTS = 8;

export const DEFAULT_CHART_WIDTH = 480;
export const DEFAULT_CHART_HEIGHT = 300;
export const MIN_CHART_WIDTH = 160;
export const MIN_CHART_HEIGHT = 120;

/** จานสีชุดข้อมูล 8 สี อ่านออกทั้งพื้นสว่างและพื้นมืด (ใช้ทั้งในแอปและในไฟล์) */
export const CHART_PALETTE = ['#4E79A7', '#F28E2B', '#E15759', '#76B7B2', '#59A14F', '#EDC948', '#B07AA1', '#FF9DA7'];

export function paletteColor(i: number): string {
  return CHART_PALETTE[((i % CHART_PALETTE.length) + CHART_PALETTE.length) % CHART_PALETTE.length];
}

let chartIdCounter = 0;
export function newChartId(): string {
  chartIdCounter++;
  return `c${Date.now().toString(36)}-${chartIdCounter}`;
}

export function newChart(type: ChartType, series: ChartSeries[], anchor: SheetChart['anchor']): SheetChart {
  return {
    id: newChartId(),
    type,
    title: '',
    series,
    anchor,
    width: DEFAULT_CHART_WIDTH,
    height: DEFAULT_CHART_HEIGHT,
    legend: type === 'pie' ? 'right' : 'bottom',
    xTitle: '',
    yTitle: '',
    dataLabels: false,
  };
}

// ---------------------------------------------------------------------------
// reference ของกราฟ: "'ชีต'!A1:B5" (แบบสูตรไม่มี "=" ระบุชื่อชีตเสมอ)
// ---------------------------------------------------------------------------

export interface ParsedChartRef {
  sheet: string;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

/** สร้าง reference ตามกติกาการใส่ '...' ของ 4B (ใช้แสดงใน UI และเก็บในโมเดล) */
export function makeChartRef(sheetName: string, r1: number, c1: number, r2 = r1, c2 = c1): string {
  return `${formatSheetReference(sheetName)}!${rangeToA1(r1, c1, r2, c2)}`;
}

/**
 * อ่าน reference ของกราฟ; defaultSheet ใช้เมื่อไม่มีชื่อชีต (เช่น ผู้ใช้พิมพ์ A1:B5)
 * คืน null ถ้าไม่ใช่ reference เดียวที่ถูกต้อง หรือเป็น #REF!
 */
export function parseChartRef(ref: string, defaultSheet?: string): ParsedChartRef | null {
  const text = ref.trim().replace(/^=/, '');
  if (!text) return null;
  const tokens = tokenize(`=${text}`).filter((t) => t.type !== 'WHITESPACE');
  if (tokens.length !== 1) return null;
  const t = tokens[0];
  const sheet = t.sheet ?? defaultSheet;
  if (sheet === undefined) return null;
  if (t.type === 'CELL_REF') {
    const p = parseA1(t.value);
    return p ? { sheet, r1: p.r, c1: p.c, r2: p.r, c2: p.c } : null;
  }
  if (t.type === 'RANGE') {
    const r = parseRange(t.value);
    if (!r) return null;
    return { sheet, r1: Math.min(r.r1, r.r2), c1: Math.min(r.c1, r.c2), r2: Math.max(r.r1, r.r2), c2: Math.max(r.c1, r.c2) };
  }
  return null;
}

/** reference ถูกลบไปแล้ว (#REF!) หรืออ่านไม่ได้ */
export function isLostRef(ref: string | undefined): boolean {
  return ref !== undefined && parseChartRef(ref) === null;
}

/** เซลในช่วงหนึ่งมิติตามลำดับ (แถวเดียว: ซ้ายไปขวา, คอลัมน์เดียว: บนลงล่าง); ช่วงสองมิติอ่านทีละแถว */
export function refCells(r: ParsedChartRef): Array<{ r: number; c: number }> {
  const out: Array<{ r: number; c: number }> = [];
  if (r.c1 === r.c2) for (let row = r.r1; row <= r.r2; row++) out.push({ r: row, c: r.c1 });
  else for (let row = r.r1; row <= r.r2; row++) for (let col = r.c1; col <= r.c2; col++) out.push({ r: row, c: col });
  return out;
}

function refLength(r: ParsedChartRef): number {
  return (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
}

// ---------------------------------------------------------------------------
// ค่าที่ใช้วาด
// ---------------------------------------------------------------------------

/** ตัวช่วยอ่านค่าจากสมุดงาน (ui ใช้ Evaluator ของ store, ไฟล์ใช้ Evaluator ใหม่) */
export interface ChartContext {
  /** index ของชีตจากชื่อ (ไม่สนตัวพิมพ์) หรือ -1 */
  sheetIndexOf(name: string): number;
  /** ค่าที่คำนวณแล้ว */
  value(sheetIndex: number, key: string): CellValue | undefined;
  /** ข้อความที่แสดงในเซล (รูปแบบวันที่/ปี ตามการตั้งค่า) */
  label(sheetIndex: number, key: string): string;
  fmt(sheetIndex: number, key: string): NumFmt | undefined;
  /** แถวที่ซ่อนด้วยตัวกรอง */
  hiddenRows(sheetIndex: number): ReadonlySet<number>;
}

export type ChartStatus = 'ok' | 'lost' | 'empty' | 'tooManySeries' | 'tooManyPoints';

export interface ResolvedSeries {
  /** index ใน chart.series */
  index: number;
  name: string;
  color: string;
  /** null = ไม่มีค่า (ช่องว่าง/error) */
  values: (number | null)[];
  hasError: boolean;
}

export interface ResolvedChart {
  status: ChartStatus;
  categories: string[];
  series: ResolvedSeries[];
  /** มี error ในช่องค่า (แสดงธงเตือน) */
  hasError: boolean;
  /** รูปแบบตัวเลขของเซลค่าแรก (ใช้กับแกนและตัวเลขบนกราฟ) */
  numFmt: NumFmt | undefined;
}

/** ค่าในช่องค่า: ตัวเลข = ค่านั้น, ข้อความ/TRUE/FALSE = 0, ว่าง = null, error = null + ธง */
export function pointValue(v: CellValue | undefined): { value: number | null; error: boolean } {
  if (v === undefined || v === '') return { value: null, error: false };
  if (typeof v === 'number') return Number.isFinite(v) ? { value: v, error: false } : { value: null, error: true };
  if (isCellError(v)) return { value: null, error: true };
  return { value: 0, error: false };
}

export function defaultSeriesName(i: number): string {
  return th.chart.defaultSeriesName.replace('{n}', String(i + 1));
}

/**
 * คำนวณค่าที่ใช้วาดกราฟ
 * skipHidden: ข้ามแถวที่ซ่อนด้วยตัวกรอง (ในแอป); ค่า cache ในไฟล์ใช้ false (ตำแหน่ง idx ตรงกับเซล)
 */
export function resolveChartData(chart: SheetChart, ctx: ChartContext, skipHidden = true): ResolvedChart {
  const empty = (status: ChartStatus): ResolvedChart => ({ status, categories: [], series: [], hasError: false, numFmt: undefined });
  if (chart.series.length > MAX_CHART_SERIES) return empty('tooManySeries');

  type Live = { index: number; s: ChartSeries; ref: ParsedChartRef; sheet: number };
  const live: Live[] = [];
  chart.series.forEach((s, index) => {
    const ref = parseChartRef(s.values);
    const sheet = ref ? ctx.sheetIndexOf(ref.sheet) : -1;
    if (ref && sheet >= 0) live.push({ index, s, ref, sheet });
  });
  // reference ของทุกชุดหาย (ช่วงถูกลบหมด หรือชีตที่อ้างถูกลบ) = "ข้อมูลหาย"
  if (live.length === 0) return empty('lost');
  if (live.some((l) => refLength(l.ref) > MAX_CHART_POINTS)) return empty('tooManyPoints');

  const series = chart.type === 'pie' ? live.slice(0, 1) : live;
  const catRef = series[0].s.categories ? parseChartRef(series[0].s.categories) : null;
  const catSheet = catRef ? ctx.sheetIndexOf(catRef.sheet) : -1;
  const catCells = catRef && catSheet >= 0 ? refCells(catRef) : null;

  const pointCount = Math.max(...series.map((l) => refLength(l.ref)));
  // จุดที่แสดง: ตัดสินจากแถวของเซลป้ายชื่อ (หรือเซลค่าของชุดแรกถ้าไม่มีป้าย)
  const axisCells = catCells ?? refCells(series[0].ref);
  const axisSheet = catCells ? catSheet : series[0].sheet;
  const axisHidden = ctx.hiddenRows(axisSheet);
  const visible: number[] = [];
  for (let i = 0; i < pointCount; i++) {
    const cell = axisCells[i];
    if (skipHidden && cell && axisHidden.has(cell.r)) continue;
    visible.push(i);
  }

  const categories = visible.map((i) => {
    const cell = catCells?.[i];
    return cell ? ctx.label(catSheet, cellToA1(cell.r, cell.c)) : String(i + 1);
  });

  let hasError = false;
  let numFmt: NumFmt | undefined;
  const out: ResolvedSeries[] = [];
  for (const l of series) {
    const cells = refCells(l.ref);
    const hidden = ctx.hiddenRows(l.sheet);
    // ชุดที่ทุกเซลค่าอยู่ในแถวที่ซ่อน (เช่น ชุดตามแถว) ไม่แสดง
    if (skipHidden && cells.every((c) => hidden.has(c.r))) continue;
    let seriesError = false;
    const values = visible.map((i) => {
      const cell = cells[i];
      if (!cell) return null;
      const key = cellToA1(cell.r, cell.c);
      const p = pointValue(ctx.value(l.sheet, key));
      if (p.error) seriesError = true;
      if (numFmt === undefined && p.value !== null) numFmt = ctx.fmt(l.sheet, key);
      return p.value;
    });
    hasError ||= seriesError;
    out.push({ index: l.index, name: seriesName(l.s, l.index, ctx), color: l.s.color, values, hasError: seriesError });
  }

  const anyValue = out.some((s) => s.values.some((v) => v !== null));
  return { status: anyValue ? 'ok' : 'empty', categories, series: out, hasError, numFmt };
}

export function seriesName(s: ChartSeries, i: number, ctx: ChartContext): string {
  if ('text' in s.name) return s.name.text.trim() || defaultSeriesName(i);
  const ref = parseChartRef(s.name.ref);
  const sheet = ref ? ctx.sheetIndexOf(ref.sheet) : -1;
  if (!ref || sheet < 0) return defaultSeriesName(i);
  return ctx.label(sheet, cellToA1(ref.r1, ref.c1)) || defaultSeriesName(i);
}

// ---------------------------------------------------------------------------
// สร้างกราฟจากช่วงที่เลือก (ค่าเริ่มต้นอัตโนมัติ)
// ---------------------------------------------------------------------------

export interface ChartBuildOptions {
  headerRow?: boolean;
  labelCol?: boolean;
  byRow?: boolean;
}

export type ChartBuildResult =
  | { ok: true; series: ChartSeries[]; headerRow: boolean; labelCol: boolean; byRow: boolean; suggestedType: ChartType }
  | { ok: false; error: 'noNumbers' | 'tooManySeries' };

function isTextValue(v: CellValue | undefined): boolean {
  return typeof v === 'string' && v !== '' && !isCellError(v);
}

/**
 * แนะนำชุดข้อมูลจากช่วง (SPEC 5.18):
 * - แถวแรกเป็นหัวข้อ: เซลที่ไม่ว่างในแถวแรกเป็นข้อความทั้งหมด และแถวถัดไปมีตัวเลข
 * - คอลัมน์แรกเป็นป้ายชื่อ: เซลที่ไม่ว่างเป็นข้อความหรือวันที่ทั้งหมด
 * - ชุดข้อมูลจากแต่ละคอลัมน์ (หรือแถว ถ้าสลับ) ที่มีตัวเลข
 * opts ใช้บังคับค่าแทนการตรวจอัตโนมัติ
 */
export function buildChartFromRange(
  ctx: ChartContext,
  sheetName: string,
  range: { r1: number; c1: number; r2: number; c2: number },
  opts: ChartBuildOptions = {}
): ChartBuildResult {
  const si = ctx.sheetIndexOf(sheetName);
  const { r1, c1, r2, c2 } = range;
  const val = (r: number, c: number) => ctx.value(si, cellToA1(r, c));
  const isNum = (r: number, c: number) => typeof val(r, c) === 'number';
  const isDate = (r: number, c: number) => isNum(r, c) && ctx.fmt(si, cellToA1(r, c)) === 'date';
  const isLabel = (r: number, c: number) => {
    const v = val(r, c);
    return v === undefined || v === '' || isTextValue(v) || isDate(r, c);
  };
  const nonEmpty = (r: number, c: number) => {
    const v = val(r, c);
    return v !== undefined && v !== '';
  };

  const byRow = opts.byRow ?? false;
  // ทิศทางตาม byRow: "เส้นหัว" = แถวแรก (หรือคอลัมน์แรกถ้าสลับ), "เส้นป้าย" = คอลัมน์แรก (หรือแถวแรก)
  const autoHeaderRow = (() => {
    if (r2 <= r1) return false;
    let anyText = false;
    for (let c = c1; c <= c2; c++) {
      const v = val(r1, c);
      if (v === undefined || v === '') continue;
      if (!isTextValue(v)) return false;
      anyText = true;
    }
    if (!anyText) return false;
    for (let c = c1; c <= c2; c++) if (isNum(r1 + 1, c) && !isDate(r1 + 1, c)) return true;
    return false;
  })();
  const headerRow = opts.headerRow ?? autoHeaderRow;
  const autoLabelCol = (() => {
    if (c2 <= c1) return false;
    let any = false;
    for (let r = headerRow ? r1 + 1 : r1; r <= r2; r++) {
      if (!isLabel(r, c1)) return false;
      if (nonEmpty(r, c1)) any = true;
    }
    return any;
  })();
  const labelCol = opts.labelCol ?? autoLabelCol;

  const dataR1 = headerRow ? r1 + 1 : r1;
  const dataC1 = labelCol ? c1 + 1 : c1;
  if (dataR1 > r2 || dataC1 > c2) return { ok: false, error: 'noNumbers' };

  const series: ChartSeries[] = [];
  if (!byRow) {
    for (let c = dataC1; c <= c2; c++) {
      let hasNum = false;
      for (let r = dataR1; r <= r2; r++) if (isNum(r, c)) hasNum = true;
      if (!hasNum) continue;
      series.push({
        name: headerRow ? { ref: makeChartRef(sheetName, r1, c) } : { text: '' },
        categories: labelCol ? makeChartRef(sheetName, dataR1, c1, r2, c1) : undefined,
        values: makeChartRef(sheetName, dataR1, c, r2, c),
        color: paletteColor(series.length),
      });
    }
  } else {
    for (let r = dataR1; r <= r2; r++) {
      let hasNum = false;
      for (let c = dataC1; c <= c2; c++) if (isNum(r, c)) hasNum = true;
      if (!hasNum) continue;
      series.push({
        name: labelCol ? { ref: makeChartRef(sheetName, r, c1) } : { text: '' },
        categories: headerRow ? makeChartRef(sheetName, r1, dataC1, r1, c2) : undefined,
        values: makeChartRef(sheetName, r, dataC1, r, c2),
        color: paletteColor(series.length),
      });
    }
  }
  if (series.length === 0) return { ok: false, error: 'noNumbers' };
  if (series.length > MAX_CHART_SERIES) return { ok: false, error: 'tooManySeries' };
  for (const s of series) if (s.categories === undefined) delete s.categories;

  // ชนิดแนะนำ: ป้ายเป็นวันที่ → เส้น, ชุดเดียวไม่เกิน 8 จุด → วงกลม, นอกนั้น → แท่ง
  const labelsAreDates = (() => {
    if (!byRow && labelCol) {
      let any = false;
      for (let r = dataR1; r <= r2; r++) {
        if (!nonEmpty(r, c1)) continue;
        if (!isDate(r, c1)) return false;
        any = true;
      }
      return any;
    }
    if (byRow && headerRow) {
      let any = false;
      for (let c = dataC1; c <= c2; c++) {
        if (!nonEmpty(r1, c)) continue;
        if (!isDate(r1, c)) return false;
        any = true;
      }
      return any;
    }
    return false;
  })();
  const points = byRow ? c2 - dataC1 + 1 : r2 - dataR1 + 1;
  const suggestedType: ChartType = labelsAreDates ? 'line' : series.length === 1 && points <= PIE_SUGGEST_MAX_POINTS ? 'pie' : 'bar';
  return { ok: true, series, headerRow, labelCol, byRow, suggestedType };
}

// ---------------------------------------------------------------------------
// reference และตำแหน่งเมื่อแทรก/ลบแถว-คอลัมน์ เปลี่ยนชื่อ/ลบ/ทำสำเนาชีต (ใช้ refshift เดิม)
// ---------------------------------------------------------------------------

/** แปลง reference ทุกตัวของกราฟด้วยฟังก์ชันของ refshift (ทำงานกับข้อความสูตรที่ขึ้นต้นด้วย "=") */
function mapChartRefs(chart: SheetChart, fn: (formula: string) => string): void {
  const apply = (ref: string) => fn(`=${ref}`).slice(1);
  for (const s of chart.series) {
    s.values = apply(s.values);
    if (s.categories !== undefined) s.categories = apply(s.categories);
    if ('ref' in s.name) s.name = { ref: apply(s.name.ref) };
  }
}

function allCharts(wb: Workbook): Array<{ chart: SheetChart; host: Sheet; hostIndex: number }> {
  const out: Array<{ chart: SheetChart; host: Sheet; hostIndex: number }> = [];
  wb.sheets.forEach((host, hostIndex) => {
    for (const chart of host.charts ?? []) out.push({ chart, host, hostIndex });
  });
  return out;
}

/** แทรกแถว/คอลัมน์ในชีต sheetIndex: ช่วงที่อ้างชีตนั้น (จากกราฟทุกชีต) และตำแหน่งของกราฟบนชีตนั้นเลื่อนตาม */
export function shiftChartsOnInsert(wb: Workbook, sheetIndex: number, axis: 'row' | 'col', at: number, count: number): void {
  const target = wb.sheets[sheetIndex];
  if (!target) return;
  for (const { chart, host, hostIndex } of allCharts(wb)) {
    const scope = { formulaSheet: host.name, targetSheet: target.name };
    mapChartRefs(chart, (f) => shiftFormulaOnInsert(f, axis, at, count, scope));
    if (hostIndex !== sheetIndex) continue;
    if (axis === 'row' && chart.anchor.row >= at) chart.anchor.row += count;
    if (axis === 'col' && chart.anchor.col >= at) chart.anchor.col += count;
  }
}

/** ลบแถว/คอลัมน์ในชีต sheetIndex: ช่วงหดหรือกลายเป็น #REF! (ข้อมูลหาย); กราฟที่ยึดกับแถวที่ถูกลบย้ายไปแถวที่มาแทน */
export function shiftChartsOnDelete(wb: Workbook, sheetIndex: number, axis: 'row' | 'col', at: number, count: number): void {
  const target = wb.sheets[sheetIndex];
  if (!target) return;
  const end = at + count - 1;
  for (const { chart, host, hostIndex } of allCharts(wb)) {
    const scope = { formulaSheet: host.name, targetSheet: target.name };
    mapChartRefs(chart, (f) => shiftFormulaOnDelete(f, axis, at, count, scope));
    if (hostIndex !== sheetIndex) continue;
    const key = axis === 'row' ? 'row' : 'col';
    const off = axis === 'row' ? 'dy' : 'dx';
    if (chart.anchor[key] > end) chart.anchor[key] -= count;
    else if (chart.anchor[key] >= at) {
      chart.anchor[key] = at;
      chart.anchor[off] = 0;
    }
  }
}

/** เปลี่ยนชื่อชีต: reference ของกราฟทุกชีตที่อ้างชื่อเดิมเปลี่ยนตาม */
export function renameSheetInCharts(wb: Workbook, oldName: string, newName: string): void {
  for (const { chart } of allCharts(wb)) mapChartRefs(chart, (f) => renameSheetInFormula(f, oldName, newName));
}

/** ชีตถูกลบ (เรียกหลังเอาชีตออกแล้ว): reference ที่ชี้ชีตนั้นเป็น #REF! — กราฟไม่ถูกลบ */
export function invalidateSheetInCharts(wb: Workbook, deletedName: string): void {
  for (const { chart } of allCharts(wb)) mapChartRefs(chart, (f) => invalidateSheetInFormula(f, deletedName));
}

/**
 * ทำสำเนาชีต (copy คือสำเนาที่คัดลอกกราฟมาแล้ว): กราฟได้ id ใหม่ และ reference ที่ชี้ชีตต้นฉบับชี้สำเนา
 * (เทียบ 4B: reference ที่ไม่มีชื่อชีตชี้สำเนาเอง — กราฟระบุชื่อชีตเสมอจึงเทียบกับชีตต้นฉบับ)
 */
export function retargetDuplicatedCharts(copy: Sheet, sourceName: string): void {
  for (const chart of copy.charts ?? []) {
    chart.id = newChartId();
    mapChartRefs(chart, (f) => renameSheetInFormula(f, sourceName, copy.name));
  }
}

/** กราฟอ้างถึงชีตนี้หรือไม่ (ใช้ตรวจก่อนลบชีต) */
export function chartReferencesSheet(chart: SheetChart, sheetName: string): boolean {
  return chart.series.some((s) =>
    [s.values, s.categories, 'ref' in s.name ? s.name.ref : undefined].some((r) => {
      const p = r ? parseChartRef(r) : null;
      return p !== null && sameSheetName(p.sheet, sheetName);
    })
  );
}

// ---------------------------------------------------------------------------
// ตำแหน่ง (px ที่ซูม 100% นับจากมุมบนซ้ายของเซล A1) <-> เซลที่ยึด
// ---------------------------------------------------------------------------

export function colWidthPx(sheet: Sheet, c: number): number {
  return sheet.colWidths[c] || DEFAULT_COL_WIDTH_PX;
}

export function rowHeightPx(sheet: Sheet, r: number, hidden?: ReadonlySet<number>): number {
  if (hidden?.has(r)) return 0;
  return sheet.rowHeights[r] || DEFAULT_ROW_HEIGHT_PX;
}

export function anchorToPx(sheet: Sheet, anchor: SheetChart['anchor'], hidden?: ReadonlySet<number>) {
  let x = 0;
  for (let c = 0; c < anchor.col; c++) x += colWidthPx(sheet, c);
  let y = 0;
  for (let r = 0; r < anchor.row; r++) y += rowHeightPx(sheet, r, hidden);
  return { x: x + anchor.dx, y: y + anchor.dy };
}

export function pxToAnchor(sheet: Sheet, x: number, y: number, hidden?: ReadonlySet<number>): SheetChart['anchor'] {
  const px = Math.max(0, x);
  const py = Math.max(0, y);
  let col = 0;
  let acc = 0;
  while (acc + colWidthPx(sheet, col) <= px) acc += colWidthPx(sheet, col++);
  const dx = Math.round(px - acc);
  let row = 0;
  acc = 0;
  // แถวที่ซ่อน (สูง 0) ไม่เป็นที่ยึด
  while (acc + rowHeightPx(sheet, row, hidden) <= py) acc += rowHeightPx(sheet, row++, hidden);
  return { row, col, dx, dy: Math.round(py - acc) };
}

// ---------------------------------------------------------------------------
// two-cell anchor ใน .xlsx (EMU) <-> anchor + ขนาด
// ---------------------------------------------------------------------------

export const EMU_PER_PX = 9525;

export interface CellAnchorEmu {
  col: number;
  colOff: number;
  row: number;
  rowOff: number;
}

/** ตำแหน่งในไฟล์ไม่ขึ้นกับตัวกรองในแอป จึงคิดทุกแถวตามความสูงจริง */
export function chartToTwoCellAnchor(sheet: Sheet, chart: SheetChart): { from: CellAnchorEmu; to: CellAnchorEmu } {
  const from: CellAnchorEmu = {
    col: chart.anchor.col,
    colOff: Math.round(chart.anchor.dx * EMU_PER_PX),
    row: chart.anchor.row,
    rowOff: Math.round(chart.anchor.dy * EMU_PER_PX),
  };
  let col = chart.anchor.col;
  let remainX = chart.anchor.dx + chart.width;
  while (remainX >= colWidthPx(sheet, col)) remainX -= colWidthPx(sheet, col++);
  let row = chart.anchor.row;
  let remainY = chart.anchor.dy + chart.height;
  while (remainY >= rowHeightPx(sheet, row)) remainY -= rowHeightPx(sheet, row++);
  return {
    from,
    to: { col, colOff: Math.round(remainX * EMU_PER_PX), row, rowOff: Math.round(remainY * EMU_PER_PX) },
  };
}

export function twoCellAnchorToChart(
  sheet: Sheet,
  from: CellAnchorEmu,
  to: CellAnchorEmu
): Pick<SheetChart, 'anchor' | 'width' | 'height'> {
  const anchor = { row: from.row, col: from.col, dx: Math.round(from.colOff / EMU_PER_PX), dy: Math.round(from.rowOff / EMU_PER_PX) };
  const a = anchorToPx(sheet, anchor);
  const b = anchorToPx(sheet, { row: to.row, col: to.col, dx: Math.round(to.colOff / EMU_PER_PX), dy: Math.round(to.rowOff / EMU_PER_PX) });
  return {
    anchor,
    width: Math.max(MIN_CHART_WIDTH, b.x - a.x),
    height: Math.max(MIN_CHART_HEIGHT, b.y - a.y),
  };
}

// ---------------------------------------------------------------------------
// แกนตัวเลข: ขีดแบ่งที่อ่านง่าย (1, 2, 5 × 10^n)
// ---------------------------------------------------------------------------

export function niceTicks(min: number, max: number, maxTicks = 6): { min: number; max: number; step: number; ticks: number[] } {
  let lo = Math.min(0, min);
  let hi = Math.max(0, max);
  if (lo === hi) hi = lo + 1;
  const raw = (hi - lo) / Math.max(1, maxTicks - 1);
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 5, 10].map((m) => m * mag).find((s) => s >= raw * (1 - 1e-9)) ?? 10 * mag;
  lo = Math.floor(lo / step + 1e-9) * step;
  hi = Math.ceil(hi / step - 1e-9) * step;
  const ticks: number[] = [];
  for (let i = 0; lo + i * step <= hi + step / 2; i++) ticks.push(Number((lo + i * step).toPrecision(12)));
  return { min: Number(lo.toPrecision(12)), max: Number(hi.toPrecision(12)), step: Number(step.toPrecision(12)), ticks };
}

// ---------------------------------------------------------------------------
// ChartContext จากสมุดงาน
// ---------------------------------------------------------------------------

/**
 * forFile: ป้ายวันที่ใน cache ของไฟล์ใช้ปี ค.ศ. เสมอ (พ.ศ. เป็นการแสดงผลในแอปเท่านั้น ตาม 16.9)
 */
export function workbookChartContext(wb: Workbook, evaluator: Evaluator, forFile = false): ChartContext {
  const sheetIndexOf = (name: string) => wb.sheets.findIndex((s) => sameSheetName(s.name, name));
  const hiddenCache = new Map<number, ReadonlySet<number>>();
  return {
    sheetIndexOf,
    value: (si, key) => (wb.sheets[si] ? evaluator.getCellValue(key, si) : undefined),
    fmt: (si, key) => wb.sheets[si]?.cells[key]?.fmt,
    label: (si, key) => {
      const sheet = wb.sheets[si];
      if (!sheet) return '';
      const v = evaluator.getCellValue(key, si);
      if (v === undefined || v === '') return '';
      const fmt = sheet.cells[key]?.fmt;
      if (forFile && fmt === 'date' && typeof v === 'number') return formatDate(v, 'ce');
      return formatDisplayValue(v, fmt);
    },
    hiddenRows: (si) => {
      let set = hiddenCache.get(si);
      if (!set) {
        set = new Set(wb.sheets[si]?.hiddenRows ?? []);
        hiddenCache.set(si, set);
      }
      return set;
    },
  };
}

// ---------------------------------------------------------------------------
// Undo/Redo: snapshot กราฟของทุกชีต (ตาม id ของชีต)
// ---------------------------------------------------------------------------

export type ChartsSnapshot = Record<string, SheetChart[]>;

export function snapshotCharts(wb: Workbook): ChartsSnapshot {
  const snap: ChartsSnapshot = {};
  for (const s of wb.sheets) if (s.charts && s.charts.length > 0) snap[s.id ?? s.name] = JSON.parse(JSON.stringify(s.charts));
  return snap;
}

export function restoreCharts(wb: Workbook, snap: ChartsSnapshot): void {
  for (const s of wb.sheets) {
    const list = snap[s.id ?? s.name];
    if (list) s.charts = JSON.parse(JSON.stringify(list));
    else delete s.charts;
  }
}
