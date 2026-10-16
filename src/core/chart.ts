import type { Sheet, SheetChart, CellValue } from './model';
import { cellToA1 } from './address';

/** ขนาดเริ่มต้นของคอลัมน์/แถวในตาราง (px ที่ซูม 100%) — Grid ใช้ค่าชุดเดียวกัน */
export const DEFAULT_COL_WIDTH_PX = 110;
export const DEFAULT_ROW_HEIGHT_PX = 28;

/** ชุดข้อมูลสูงสุดต่อกราฟ (SPEC 5.18) */
export const MAX_CHART_SERIES = 6;
export const DEFAULT_CHART_WIDTH = 480;
export const DEFAULT_CHART_HEIGHT = 300;
export const MIN_CHART_WIDTH = 160;
export const MIN_CHART_HEIGHT = 120;

export type ChartRange = NonNullable<SheetChart['range']>;

export interface ChartSeries {
  name: string;
  /** ค่าตามลำดับหมวด; null = เซลว่าง/ไม่ใช่ตัวเลข (ไม่มีค่า) */
  values: (number | null)[];
  /** คอลัมน์ของชุดข้อมูลในชีต (0-based) */
  col: number;
}

export interface ChartData {
  categories: string[];
  series: ChartSeries[];
  /** แถวแรกของช่วงเป็นหัวตาราง (ชื่อชุดข้อมูล) */
  hasHeader: boolean;
  /** มีคอลัมน์หมวด (ช่วงกว้างกว่า 1 คอลัมน์) */
  hasCategoryColumn: boolean;
  /** แถวข้อมูลจริง (ไม่รวมหัว) ที่แสดงในกราฟ — ข้ามแถวที่ซ่อนด้วยตัวกรอง */
  rows: number[];
}

let chartIdCounter = 0;
export function newChartId(): string {
  chartIdCounter++;
  return `c${Date.now().toString(36)}-${chartIdCounter}`;
}

function isText(v: CellValue | undefined): v is string {
  return typeof v === 'string' && v !== '' && !v.startsWith('#');
}

/**
 * อ่านข้อมูลกราฟจากช่วง: คอลัมน์แรก = หมวด, คอลัมน์ถัดไป = ชุดข้อมูล (สูงสุด 6)
 * ถ้าช่วงกว้าง 1 คอลัมน์ ใช้คอลัมน์นั้นเป็นชุดข้อมูลและใช้ลำดับ 1, 2, 3… เป็นหมวด
 * getValue: ค่าที่คำนวณแล้วของเซล, label: ข้อความที่แสดงของเซล (ใช้เป็นชื่อหมวด)
 */
export function getChartData(
  range: ChartRange | null,
  getValue: (key: string) => CellValue | undefined,
  label: (key: string) => string,
  hiddenRows: ReadonlySet<number> = new Set()
): ChartData {
  if (!range) return { categories: [], series: [], hasHeader: false, hasCategoryColumn: false, rows: [] };
  const { r1, c1, r2, c2 } = range;
  const hasCategoryColumn = c2 > c1;
  const firstSeriesCol = hasCategoryColumn ? c1 + 1 : c1;
  const lastSeriesCol = Math.min(c2, firstSeriesCol + MAX_CHART_SERIES - 1);

  let hasHeader = false;
  if (r2 > r1) {
    for (let c = firstSeriesCol; c <= lastSeriesCol; c++) {
      if (isText(getValue(cellToA1(r1, c)))) hasHeader = true;
    }
  }
  const dataStart = hasHeader ? r1 + 1 : r1;
  const rows: number[] = [];
  for (let r = dataStart; r <= r2; r++) if (!hiddenRows.has(r)) rows.push(r);

  const categories = rows.map((r, i) => (hasCategoryColumn ? label(cellToA1(r, c1)) : String(i + 1)));
  const series: ChartSeries[] = [];
  for (let c = firstSeriesCol; c <= lastSeriesCol; c++) {
    const head = hasHeader ? getValue(cellToA1(r1, c)) : undefined;
    series.push({
      name: hasHeader && head !== undefined && head !== '' ? String(head) : `ชุดข้อมูล ${series.length + 1}`,
      col: c,
      values: rows.map((r) => {
        const v = getValue(cellToA1(r, c));
        return typeof v === 'number' && Number.isFinite(v) ? v : null;
      }),
    });
  }
  return { categories, series, hasHeader, hasCategoryColumn, rows };
}

/** มีตัวเลขให้ทำกราฟอย่างน้อยหนึ่งค่า */
export function hasNumericData(data: ChartData): boolean {
  return data.series.some((s) => s.values.some((v) => v !== null));
}

/** ตำแหน่งเริ่มต้นของกราฟใหม่: ขวาของช่วงข้อมูลเว้นหนึ่งคอลัมน์ แถวเดียวกับบรรทัดแรก */
export function defaultChartAnchor(range: ChartRange): SheetChart['anchor'] {
  return { row: range.r1, col: range.c2 + 2, dx: 0, dy: 0 };
}

// ---------------------------------------------------------------------------
// เลื่อนช่วงข้อมูล/ตำแหน่งกราฟเมื่อแทรก/ลบแถว-คอลัมน์ (แบบเดียวกับช่วงในสูตร)
// ---------------------------------------------------------------------------

function shiftSpanOnInsert(a: number, b: number, at: number, count: number): [number, number] {
  if (a >= at) return [a + count, b + count];
  if (b >= at) return [a, b + count];
  return [a, b];
}

/** ช่วง [a,b] หลังลบ [at, at+count-1]; null = ถูกลบทั้งหมด */
function shiftSpanOnDelete(a: number, b: number, at: number, count: number): [number, number] | null {
  const end = at + count - 1;
  if (b < at) return [a, b];
  if (a > end) return [a - count, b - count];
  if (a >= at && b <= end) return null;
  const na = a < at ? a : at;
  const nb = b > end ? b - count : at - 1;
  return [na, nb];
}

export function shiftChartsOnInsert(sheet: Sheet, axis: 'row' | 'col', at: number, count: number) {
  for (const ch of sheet.charts ?? []) {
    if (ch.range) {
      if (axis === 'row') [ch.range.r1, ch.range.r2] = shiftSpanOnInsert(ch.range.r1, ch.range.r2, at, count);
      else [ch.range.c1, ch.range.c2] = shiftSpanOnInsert(ch.range.c1, ch.range.c2, at, count);
    }
    if (axis === 'row' && ch.anchor.row >= at) ch.anchor.row += count;
    if (axis === 'col' && ch.anchor.col >= at) ch.anchor.col += count;
  }
}

export function shiftChartsOnDelete(sheet: Sheet, axis: 'row' | 'col', at: number, count: number) {
  const end = at + count - 1;
  for (const ch of sheet.charts ?? []) {
    if (ch.range) {
      const span =
        axis === 'row'
          ? shiftSpanOnDelete(ch.range.r1, ch.range.r2, at, count)
          : shiftSpanOnDelete(ch.range.c1, ch.range.c2, at, count);
      if (!span) ch.range = null;
      else if (axis === 'row') [ch.range.r1, ch.range.r2] = span;
      else [ch.range.c1, ch.range.c2] = span;
    }
    const key = axis === 'row' ? 'row' : 'col';
    const off = axis === 'row' ? 'dy' : 'dx';
    if (ch.anchor[key] > end) ch.anchor[key] -= count;
    else if (ch.anchor[key] >= at) {
      // แถว/คอลัมน์ที่ยึดไว้ถูกลบ: ยึดกับแถว/คอลัมน์ที่มาแทนที่
      ch.anchor[key] = at;
      ch.anchor[off] = 0;
    }
  }
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
// แกนตัวเลข: ขีดแบ่งที่อ่านง่าย (1, 2, 2.5, 5 × 10^n)
// ---------------------------------------------------------------------------

export function niceTicks(min: number, max: number, maxTicks = 6): { min: number; max: number; step: number; ticks: number[] } {
  let lo = Math.min(0, min);
  let hi = Math.max(0, max);
  if (lo === hi) hi = lo + 1;
  const raw = (hi - lo) / Math.max(1, maxTicks - 1);
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  lo = Math.floor(lo / step) * step;
  hi = Math.ceil(hi / step) * step;
  const ticks: number[] = [];
  for (let v = lo; v <= hi + step / 2; v += step) ticks.push(Number(v.toFixed(10)));
  return { min: lo, max: hi, step, ticks };
}
