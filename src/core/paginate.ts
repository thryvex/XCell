import { Sheet, Workbook, PrintSettings, PaperSize, PrintMargins } from './model';
import { parseA1 } from './address';
import { CellRange } from './sort';

/** 1 px ของตาราง (96 dpi) = 25.4/96 มม. */
export const PX_TO_MM = 25.4 / 96;
export const DEFAULT_COL_WIDTH_PX = 110;
export const DEFAULT_ROW_HEIGHT_PX = 28;

export const PAPER_SIZES_MM: Record<PaperSize, { w: number; h: number }> = {
  A4: { w: 210, h: 297 },
  Letter: { w: 215.9, h: 279.4 },
  A5: { w: 148, h: 210 },
};

/** ระยะขอบแบบเดียวกับค่าสำเร็จรูปของ Excel (นิ้ว → มม.) */
export const MARGINS_MM: Record<PrintMargins, { top: number; right: number; bottom: number; left: number }> = {
  normal: { top: 19.05, right: 17.78, bottom: 19.05, left: 17.78 },
  narrow: { top: 19.05, right: 6.35, bottom: 19.05, left: 6.35 },
  wide: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 },
};

export const HEADER_MM = 8; // หัวกระดาษ (ชื่อไฟล์ + วันที่)
export const FOOTER_MM = 8; // ท้ายกระดาษ (เลขหน้า)
export const MIN_SCALE = 0.3;
export const MAX_SCALE = 2;
export const MANY_PAGES_WARNING = 200;

export interface PageGeometry {
  pageW: number;
  pageH: number;
  margins: { top: number; right: number; bottom: number; left: number };
  headerH: number;
  footerH: number;
  contentW: number; // พื้นที่ตารางหลังหักขอบ
  contentH: number; // พื้นที่ตารางหลังหักขอบและหัว/ท้ายกระดาษ
}

export function pageGeometry(settings: PrintSettings): PageGeometry {
  const paper = PAPER_SIZES_MM[settings.paper];
  const landscape = settings.orientation === 'landscape';
  const pageW = landscape ? paper.h : paper.w;
  const pageH = landscape ? paper.w : paper.h;
  const margins = MARGINS_MM[settings.margins];
  const headerH = settings.header ? HEADER_MM : 0;
  const footerH = settings.pageNumbers ? FOOTER_MM : 0;
  return {
    pageW,
    pageH,
    margins,
    headerH,
    footerH,
    contentW: pageW - margins.left - margins.right,
    contentH: pageH - margins.top - margins.bottom - headerH - footerH,
  };
}

export function colWidthMm(sheet: Sheet, c: number): number {
  return (sheet.colWidths[c] ?? DEFAULT_COL_WIDTH_PX) * PX_TO_MM;
}

export function rowHeightMm(sheet: Sheet, r: number): number {
  return (sheet.rowHeights[r] ?? DEFAULT_ROW_HEIGHT_PX) * PX_TO_MM;
}

/** ช่วงที่พิมพ์เริ่มต้น: A1 ถึงเซลสุดท้ายที่มีข้อมูลหรือสไตล์ (รวมเซลผสาน) หรือ null ถ้าว่าง */
export function printExtent(sheet: Sheet): CellRange | null {
  let maxR = -1;
  let maxC = -1;
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (cell.raw === '' && !cell.style) continue;
    const p = parseA1(key);
    if (!p) continue;
    maxR = Math.max(maxR, p.r);
    maxC = Math.max(maxC, p.c);
  }
  if (maxR < 0) return null;
  for (const m of sheet.merges) {
    if (m.r1 <= maxR && m.c1 <= maxC) {
      maxR = Math.max(maxR, m.r2);
      maxC = Math.max(maxC, m.c2);
    }
  }
  return { r1: 0, c1: 0, r2: maxR, c2: maxC };
}

export interface PrintPage {
  sheetIndex: number;
  scale: number; // มาตราส่วนที่ใช้ (1 = 100%)
  titleRows: number[]; // แถวตรึงที่พิมพ์ซ้ำด้านบน
  rows: number[];
  titleCols: number[]; // คอลัมน์ตรึงที่พิมพ์ซ้ำด้านซ้าย
  cols: number[];
}

export interface Pagination {
  pages: PrintPage[];
  scale: number;
  /** มีกลุ่มเซลผสานที่สูง/กว้างเกินหนึ่งหน้า จึงต้องตัดคร่อมหน้า */
  mergeSplit: boolean;
}

/**
 * แบ่งรายการแถว (หรือคอลัมน์) เป็นชุดๆ ตามพื้นที่ที่มี
 * - spans: ช่วงเซลผสาน [a, b] ห้ามตัดคร่อม ย้ายจุดแบ่งไปก่อนกลุ่ม (ถ้ากลุ่มยาวกว่าหน้าเดียว ยอมตัด)
 * - อย่างน้อยหนึ่งรายการต่อชุดเสมอ (แถวที่สูงกว่าหน้าก็พิมพ์เดี่ยวๆ)
 */
export function splitIntoChunks(
  items: number[],
  size: (item: number) => number,
  available: number,
  spans: Array<[number, number]>
): { chunks: number[][]; split: boolean } {
  const chunks: number[][] = [];
  let split = false;
  let i = 0;
  const EPS = 1e-6;

  while (i < items.length) {
    const start = i;
    let used = 0;
    let j = i;
    while (j < items.length && (j === start || used + size(items[j]) <= available + EPS)) {
      used += size(items[j]);
      j++;
    }

    let breakAt = j;
    if (breakAt < items.length) {
      // ย้ายจุดแบ่งไม่ให้ตัดกลุ่มเซลผสาน
      for (let guard = 0; guard < items.length; guard++) {
        const prev = items[breakAt - 1];
        const next = items[breakAt];
        const crossing = spans.find(([a, b]) => a <= prev && b >= next);
        if (!crossing) break;
        let p = start;
        while (p < breakAt && items[p] < crossing[0]) p++;
        if (p <= start) {
          split = true; // กลุ่มยาวกว่าหนึ่งหน้า: ยอมตัด
          break;
        }
        breakAt = p;
      }
    }

    chunks.push(items.slice(start, breakAt));
    i = breakAt;
  }

  return { chunks, split };
}

function range(a: number, b: number): number[] {
  const out: number[] = [];
  for (let x = a; x <= b; x++) out.push(x);
  return out;
}

/**
 * แบ่งหน้าของชีตเดียว: ลงล่างก่อนแล้วไปทางขวา
 * ใช้ความกว้าง/สูงจริงของตาราง (มม.) ตามมาตราส่วน หักขอบและหัว/ท้ายกระดาษ ข้ามแถวที่ซ่อนด้วยตัวกรอง
 */
export function paginateSheet(
  sheet: Sheet,
  printRange: CellRange | null,
  settings: PrintSettings,
  sheetIndex = 0
): Pagination {
  if (!printRange) return { pages: [], scale: 1, mergeSplit: false };
  const geo = pageGeometry(settings);
  const hidden = new Set(sheet.hiddenRows ?? []);

  let titleRows =
    settings.repeatFrozen && sheet.frozenRows > 0
      ? range(0, sheet.frozenRows - 1).filter((r) => !hidden.has(r))
      : [];
  let titleCols = settings.repeatFrozen && sheet.frozenCols > 0 ? range(0, sheet.frozenCols - 1) : [];
  const titleRowSet = new Set(titleRows);
  const titleColSet = new Set(titleCols);

  let bodyRows = range(printRange.r1, printRange.r2).filter((r) => !hidden.has(r) && !titleRowSet.has(r));
  let bodyCols = range(printRange.c1, printRange.c2).filter((c) => !titleColSet.has(c));
  // ช่วงที่พิมพ์อยู่ในส่วนที่ตรึงทั้งหมด: พิมพ์ส่วนนั้นเป็นเนื้อหาปกติ
  if (bodyRows.length === 0 && titleRows.length > 0) {
    bodyRows = titleRows.filter((r) => r >= printRange.r1 && r <= printRange.r2);
    titleRows = [];
  }
  if (bodyCols.length === 0 && titleCols.length > 0) {
    bodyCols = titleCols.filter((c) => c >= printRange.c1 && c <= printRange.c2);
    titleCols = [];
  }
  if (bodyRows.length === 0 || bodyCols.length === 0) return { pages: [], scale: 1, mergeSplit: false };

  let scale = 1;
  if (settings.scaleMode === 'custom') {
    scale = Math.min(MAX_SCALE, Math.max(MIN_SCALE, settings.scalePercent / 100));
  } else if (settings.scaleMode === 'fitWidth') {
    const totalW = [...titleCols, ...bodyCols].reduce((s, c) => s + colWidthMm(sheet, c), 0);
    scale = Math.min(1, geo.contentW / totalW);
  }

  const rowH = (r: number) => rowHeightMm(sheet, r) * scale;
  const colW = (c: number) => colWidthMm(sheet, c) * scale;

  // แถว/คอลัมน์ที่ตรึงกินพื้นที่จนไม่เหลือที่ให้เนื้อหา → ไม่พิมพ์ซ้ำ
  let availH = geo.contentH - titleRows.reduce((s, r) => s + rowH(r), 0);
  if (titleRows.length > 0 && availH < Math.min(...bodyRows.map(rowH))) {
    titleRows = [];
    availH = geo.contentH;
  }
  let availW = geo.contentW - titleCols.reduce((s, c) => s + colW(c), 0);
  if (titleCols.length > 0 && availW < Math.min(...bodyCols.map(colW))) {
    titleCols = [];
    availW = geo.contentW;
  }

  const colsInRange = (m: { c1: number; c2: number }) => m.c2 >= printRange.c1 && m.c1 <= printRange.c2;
  const rowsInRange = (m: { r1: number; r2: number }) => m.r2 >= printRange.r1 && m.r1 <= printRange.r2;
  const rowSpans = sheet.merges.filter((m) => colsInRange(m) && m.r2 > m.r1).map((m) => [m.r1, m.r2] as [number, number]);
  const colSpans = sheet.merges.filter((m) => rowsInRange(m) && m.c2 > m.c1).map((m) => [m.c1, m.c2] as [number, number]);

  const rowSplit = splitIntoChunks(bodyRows, rowH, availH, rowSpans);
  const colSplit =
    settings.scaleMode === 'fitWidth'
      ? { chunks: [bodyCols], split: false }
      : splitIntoChunks(bodyCols, colW, availW, colSpans);

  const pages: PrintPage[] = [];
  for (const cols of colSplit.chunks) {
    for (const rows of rowSplit.chunks) {
      pages.push({ sheetIndex, scale, titleRows, rows, titleCols, cols });
    }
  }
  return { pages, scale, mergeSplit: rowSplit.split || colSplit.split };
}

export interface PrintJob {
  pages: PrintPage[];
  mergeSplit: boolean;
}

/**
 * รวมหน้าที่จะพิมพ์ตามตัวเลือก: ชีตปัจจุบัน / ทุกชีต (ไม่รวมชีตซ่อน) / ช่วงที่เลือก
 */
export function buildPrintJob(
  wb: Workbook,
  activeSheetIndex: number,
  selection: CellRange,
  settings: PrintSettings
): PrintJob {
  const results: Pagination[] = [];
  if (settings.what === 'all') {
    wb.sheets.forEach((sheet, i) => {
      if (sheet.hidden) return;
      results.push(paginateSheet(sheet, printExtent(sheet), settings, i));
    });
  } else {
    const sheet = wb.sheets[activeSheetIndex];
    const isMulti = selection.r1 !== selection.r2 || selection.c1 !== selection.c2;
    const target = settings.what === 'selection' && isMulti ? selection : printExtent(sheet);
    results.push(paginateSheet(sheet, target, settings, activeSheetIndex));
  }
  return {
    pages: results.flatMap((r) => r.pages),
    mergeSplit: results.some((r) => r.mergeSplit),
  };
}
