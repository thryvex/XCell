import { Sheet, Cell, CellValue, NumFmt } from './model';
import { parseDateString } from './numfmt';
import { Evaluator } from './evaluator';
import { cellToA1, parseRange } from './address';
import { shiftFormula } from './refshift';
import { tokenize } from './tokenizer';
import { sameSheetName } from './sheetName';
import { isCellError } from './functions';
import { CellPatch } from './history';

export interface CellRange {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

export type SortDirection = 'asc' | 'desc';

export interface SortKey {
  col: number; // index คอลัมน์จริง (0-based)
  dir: SortDirection;
}

function isFilled(sheet: Sheet, r: number, c: number): boolean {
  if (r < 0 || c < 0) return false;
  const cell = sheet.cells[cellToA1(r, c)];
  return !!cell && cell.raw !== '' && cell.raw !== undefined;
}

/**
 * เดาช่วงข้อมูลรอบเซล (r, c): ขยายกรอบตราบใดที่แถว/คอลัมน์ที่ติดขอบ (รวมแนวทแยง) มีข้อมูล
 * เหมือน "ช่วงข้อมูลปัจจุบัน" ของ Google Sheets / Excel
 */
export function detectDataRegion(sheet: Sheet, r: number, c: number): CellRange {
  let r1 = r;
  let r2 = r;
  let c1 = c;
  let c2 = c;

  const rowHasData = (row: number, from: number, to: number) => {
    for (let col = Math.max(0, from); col <= to; col++) if (isFilled(sheet, row, col)) return true;
    return false;
  };
  const colHasData = (col: number, from: number, to: number) => {
    for (let row = Math.max(0, from); row <= to; row++) if (isFilled(sheet, row, col)) return true;
    return false;
  };

  let changed = true;
  while (changed) {
    changed = false;
    if (r1 > 0 && rowHasData(r1 - 1, c1 - 1, c2 + 1)) {
      r1--;
      changed = true;
    }
    if (rowHasData(r2 + 1, c1 - 1, c2 + 1)) {
      r2++;
      changed = true;
    }
    if (c1 > 0 && colHasData(c1 - 1, r1 - 1, r2 + 1)) {
      c1--;
      changed = true;
    }
    if (colHasData(c2 + 1, r1 - 1, r2 + 1)) {
      c2++;
      changed = true;
    }
  }

  return { r1, c1, r2, c2 };
}

export function rangeHasData(sheet: Sheet, range: CellRange): boolean {
  for (let r = range.r1; r <= range.r2; r++) {
    for (let c = range.c1; c <= range.c2; c++) {
      if (isFilled(sheet, r, c)) return true;
    }
  }
  return false;
}

/**
 * แถวแรกเป็นหัวตารางหรือไม่: ค่าทุกเซลที่มีข้อมูลในแถวแรกเป็นข้อความ
 * และแถวถัดไปมีตัวเลข/วันที่อย่างน้อยหนึ่งเซล
 */
export function detectHeaderRow(
  sheet: Sheet,
  evaluator: Evaluator,
  sheetIndex: number,
  range: CellRange
): boolean {
  if (range.r2 <= range.r1) return false;

  let headerCount = 0;
  for (let c = range.c1; c <= range.c2; c++) {
    const v = evaluator.getCellValue(cellToA1(range.r1, c), sheetIndex);
    if (v === '') continue;
    if (typeof v !== 'string' || isCellError(v)) return false;
    headerCount++;
  }
  if (headerCount === 0) return false;

  for (let c = range.c1; c <= range.c2; c++) {
    const v = evaluator.getCellValue(cellToA1(range.r1 + 1, c), sheetIndex);
    if (typeof v === 'number') return true;
  }
  return false;
}

export function rangeHasMerge(sheet: Sheet, range: CellRange): boolean {
  return sheet.merges.some(
    (m) => !(m.r2 < range.r1 || m.r1 > range.r2 || m.c2 < range.c1 || m.c1 > range.c2)
  );
}

/**
 * สูตรในแถวข้อมูลของช่วงอ้างอิงแถวอื่นในช่วงเดียวกันหรือไม่ (เช่น คงเหลือสะสม =F2+D3-E3)
 */
export function formulasReferenceOtherRows(sheet: Sheet, range: CellRange, hasHeader: boolean): boolean {
  const dataR1 = hasHeader ? range.r1 + 1 : range.r1;

  for (let r = dataR1; r <= range.r2; r++) {
    for (let c = range.c1; c <= range.c2; c++) {
      const raw = sheet.cells[cellToA1(r, c)]?.raw;
      if (!raw || !raw.startsWith('=')) continue;

      for (const t of tokenize(raw)) {
        if (t.type !== 'CELL_REF' && t.type !== 'RANGE') continue;
        if (t.sheet !== undefined && !sameSheetName(t.sheet, sheet.name)) continue;
        const rng = parseRange(t.value);
        if (!rng) continue;
        const colsOverlap = rng.c2 >= range.c1 && rng.c1 <= range.c2;
        const lo = Math.max(rng.r1, dataR1);
        const hi = Math.min(rng.r2, range.r2);
        if (!colsOverlap || lo > hi) continue;
        // แถวที่อ้างในช่วงข้อมูล มีแถวที่ไม่ใช่แถวตัวเองหรือไม่
        if (!(lo === r && hi === r)) return true;
      }
    }
  }
  return false;
}

const thaiCollator = new Intl.Collator('th', { sensitivity: 'accent', numeric: false });

/** หมวดของค่า: ตัวเลข/วันที่ → ข้อความ → TRUE/FALSE → error → ค่าว่าง */
function category(v: CellValue): number {
  if (v === '' || v === null || v === undefined) return 4;
  if (typeof v === 'number') return 0;
  if (typeof v === 'boolean') return 2;
  if (isCellError(v)) return 3;
  return 1;
}

/**
 * เปรียบเทียบค่าสำหรับเรียงลำดับ: ค่าว่างอยู่ท้ายสุดเสมอทั้ง A→Z และ Z→A
 * ข้อความเรียงด้วย Intl.Collator ภาษาไทย ไม่สนตัวพิมพ์เล็ก/ใหญ่
 */
export function compareSortValues(a: CellValue, b: CellValue, dir: SortDirection): number {
  const ca = category(a);
  const cb = category(b);
  if (ca === 4 || cb === 4) return ca === cb ? 0 : ca === 4 ? 1 : -1;

  let base = ca - cb;
  if (base === 0) {
    if (ca === 0) base = (a as number) - (b as number);
    else if (ca === 1) base = thaiCollator.compare(String(a), String(b));
    else if (ca === 2) base = Number(a) - Number(b);
    else base = String(a).localeCompare(String(b));
  }
  return dir === 'asc' ? base : -base;
}

/**
 * เรียงแถวข้อมูลของช่วง (ไม่รวมหัวตาราง) ตาม keys แบบ stable
 * ย้ายทั้งแถวของช่วง (ค่า สูตร สไตล์ รูปแบบตัวเลข); สูตรย้ายแบบคัดลอกวาง (relative เลื่อนตามแถวใหม่)
 * แก้ไขชีตโดยตรง และคืน patch ของเซลที่เปลี่ยน
 */
export function sortRange(
  sheet: Sheet,
  evaluator: Evaluator,
  sheetIndex: number,
  range: CellRange,
  keys: SortKey[],
  hasHeader: boolean,
  orderOut?: number[] // (ไม่บังคับ) รับลำดับแถวเดิมของแต่ละตำแหน่งใหม่
): CellPatch[] {
  const dataR1 = hasHeader ? range.r1 + 1 : range.r1;
  if (dataR1 >= range.r2 || keys.length === 0) return [];

  const rows: { row: number; values: CellValue[] }[] = [];
  for (let r = dataR1; r <= range.r2; r++) {
    rows.push({
      row: r,
      values: keys.map((k) => {
        const key = cellToA1(r, k.col);
        return sortableValue(evaluator.getCellValue(key, sheetIndex), sheet.cells[key]?.fmt);
      }),
    });
  }

  rows.sort((x, y) => {
    for (let i = 0; i < keys.length; i++) {
      const cmp = compareSortValues(x.values[i], y.values[i], keys[i].dir);
      if (cmp !== 0) return cmp;
    }
    return 0; // Array.prototype.sort เป็น stable
  });

  // เก็บเซลเดิมทั้งช่วงก่อนเขียนทับ
  const before: Record<string, Cell | undefined> = {};
  for (let r = dataR1; r <= range.r2; r++) {
    for (let c = range.c1; c <= range.c2; c++) {
      const key = cellToA1(r, c);
      before[key] = sheet.cells[key] ? { ...sheet.cells[key] } : undefined;
    }
  }

  const patches: CellPatch[] = [];
  if (orderOut) orderOut.push(...rows.map((x) => x.row));
  rows.forEach(({ row: oldRow }, i) => {
    const newRow = dataR1 + i;
    for (let c = range.c1; c <= range.c2; c++) {
      const key = cellToA1(newRow, c);
      const src = before[cellToA1(oldRow, c)];
      let next: Cell | undefined;
      if (src) {
        next = { ...src };
        if (src.raw && src.raw.startsWith('=') && newRow !== oldRow) {
          next.raw = shiftFormula(src.raw, newRow - oldRow, 0);
        }
      }

      const prev = before[key];
      if (next) sheet.cells[key] = next;
      else delete sheet.cells[key];

      if (JSON.stringify(prev) !== JSON.stringify(next)) {
        patches.push({ key, before: prev, after: next ? { ...next } : undefined });
      }
    }
  });

  return patches;
}

/** แปลงช่วงเป็นข้อความ เช่น A1:F20 */
export function rangeLabel(range: CellRange): string {
  const a = cellToA1(range.r1, range.c1);
  const b = cellToA1(range.r2, range.c2);
  return a === b ? a : `${a}:${b}`;
}

/**
 * วันที่ที่ผู้ใช้พิมพ์ (เช่น "5/10/2026") ยังเก็บเป็นข้อความ: ถ้าเซลจัดรูปแบบเป็นวันที่
 * ให้เรียง/กรองตามค่าวันที่ (serial) แทนการเรียงแบบข้อความ
 */
export function sortableValue(value: CellValue, fmt: NumFmt | undefined): CellValue {
  if (typeof value === 'string' && fmt === 'date') {
    const serial = parseDateString(value);
    if (serial !== null) return serial;
  }
  return value;
}
