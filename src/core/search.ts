import { Workbook, Sheet } from './model';
import { Evaluator } from './evaluator';
import { parseA1 } from './address';
import { formatDisplayValue } from './numfmt';

export interface SearchOptions {
  matchCase: boolean; // ตรงตัวพิมพ์เล็ก/ใหญ่
  wholeCell: boolean; // ให้ตรงทั้งเซล
  includeFormulas: boolean; // ค้นในสูตรด้วย
  allSheets: boolean; // ค้นทุกชีต
}

export const DEFAULT_SEARCH_OPTIONS: SearchOptions = {
  matchCase: false,
  wholeCell: false,
  includeFormulas: false,
  allSheets: false,
};

export interface SearchMatch {
  sheetIndex: number;
  r: number;
  c: number;
}

export function hiddenRowSet(sheet: Sheet): Set<number> {
  return new Set(sheet.hiddenRows ?? []);
}

function textMatches(text: string, query: string, opts: SearchOptions): boolean {
  const a = opts.matchCase ? text : text.toLocaleLowerCase();
  const b = opts.matchCase ? query : query.toLocaleLowerCase();
  return opts.wholeCell ? a === b : a.includes(b);
}

/**
 * ค้นหาเซลที่ตรงกับคำค้น เรียงตามลำดับชีต → แถว → คอลัมน์
 * - ค่าเริ่มต้นค้นใน "ค่าที่แสดง" (เช่น ฿1,234.00, 01/10/2026)
 * - includeFormulas: ค้นในข้อความสูตรด้วย (ตรงอย่างใดอย่างหนึ่งก็นับ)
 * - ไม่นับเซลในแถวที่ถูกตัวกรองซ่อน และไม่ค้นในชีตซ่อน
 */
export function findMatches(
  wb: Workbook,
  evaluator: Evaluator,
  query: string,
  opts: SearchOptions,
  activeSheetIndex: number
): SearchMatch[] {
  if (!query) return [];

  const sheetIndices = opts.allSheets
    ? wb.sheets.map((_, i) => i).filter((i) => !wb.sheets[i].hidden)
    : [activeSheetIndex];

  const matches: SearchMatch[] = [];
  for (const sheetIndex of sheetIndices) {
    const sheet = wb.sheets[sheetIndex];
    if (!sheet) continue;
    const hidden = hiddenRowSet(sheet);
    const sheetMatches: SearchMatch[] = [];

    for (const [key, cell] of Object.entries(sheet.cells)) {
      if (!cell.raw) continue;
      const p = parseA1(key);
      if (!p || hidden.has(p.r)) continue;

      const display = formatDisplayValue(evaluator.getCellValue(key, sheetIndex), cell.fmt);
      const formulaHit =
        opts.includeFormulas && cell.raw.startsWith('=') && textMatches(cell.raw, query, opts);
      if (formulaHit || textMatches(display, query, opts)) {
        sheetMatches.push({ sheetIndex, r: p.r, c: p.c });
      }
    }

    sheetMatches.sort((a, b) => a.r - b.r || a.c - b.c);
    matches.push(...sheetMatches);
  }

  return matches;
}

function comparePosition(a: SearchMatch, b: SearchMatch, sheetOrder: (i: number) => number): number {
  return sheetOrder(a.sheetIndex) - sheetOrder(b.sheetIndex) || a.r - b.r || a.c - b.c;
}

/**
 * ผลถัดไป (direction 1) หรือก่อนหน้า (-1) จากตำแหน่งปัจจุบัน วนรอบเมื่อสุดรายการ
 * inclusive: นับเซลปัจจุบันด้วย (ใช้ตอนพิมพ์คำค้นครั้งแรก)
 */
export function nextMatchIndex(
  matches: SearchMatch[],
  from: SearchMatch,
  direction: 1 | -1,
  inclusive = false
): number {
  if (matches.length === 0) return -1;
  const order = (i: number) => i;

  if (direction === 1) {
    const idx = matches.findIndex((m) => {
      const cmp = comparePosition(m, from, order);
      return inclusive ? cmp >= 0 : cmp > 0;
    });
    return idx >= 0 ? idx : 0;
  }

  for (let i = matches.length - 1; i >= 0; i--) {
    const cmp = comparePosition(matches[i], from, order);
    if (inclusive ? cmp <= 0 : cmp < 0) return i;
  }
  return matches.length - 1;
}
