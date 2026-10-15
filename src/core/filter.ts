import { Sheet, SheetFilter, ColumnFilter, FilterCondition, CellValue } from './model';
import { Evaluator } from './evaluator';
import { cellToA1 } from './address';
import { formatDisplayValue, parseDateString } from './numfmt';
import { CellRange } from './sort';

export function createFilter(range: CellRange): SheetFilter {
  return { r1: range.r1, c1: range.c1, r2: range.r2, c2: range.c2, columns: {} };
}

export function hasActiveCriteria(filter: SheetFilter | undefined): boolean {
  return !!filter && Object.keys(filter.columns).length > 0;
}

/** จำนวนแถวข้อมูลใต้หัวตาราง */
export function filterDataRowCount(filter: SheetFilter): number {
  return Math.max(0, filter.r2 - filter.r1);
}

/** แปลงข้อความในเงื่อนไขเป็นตัวเลข (รองรับ 1,234.5 และวันที่ 1/10/2026) หรือ null */
export function parseConditionNumber(text: string): number | null {
  const trimmed = text.trim();
  if (trimmed === '') return null;
  const cleaned = trimmed.replace(/,/g, '');
  if (!isNaN(Number(cleaned))) return Number(cleaned);
  return parseDateString(trimmed);
}

/**
 * ตรวจเงื่อนไขกับค่าหนึ่งเซล
 * - เงื่อนไขข้อความเทียบกับ "ค่าที่แสดง" แบบไม่สนตัวพิมพ์เล็ก/ใหญ่
 * - มากกว่า / น้อยกว่า / ระหว่าง ใช้กับตัวเลขและวันที่เท่านั้น (ข้อความ/ว่าง ไม่ผ่าน)
 */
export function conditionMatches(cond: FilterCondition, value: CellValue, display: string): boolean {
  const text = display.toLocaleLowerCase();
  switch (cond.type) {
    case 'empty':
      return display === '';
    case 'notEmpty':
      return display !== '';
    case 'textContains':
      return text.includes(cond.value.toLocaleLowerCase());
    case 'textNotContains':
      return !text.includes(cond.value.toLocaleLowerCase());
    case 'startsWith':
      return text.startsWith(cond.value.toLocaleLowerCase());
    case 'equals': {
      const target = parseConditionNumber(cond.value);
      if (target !== null && typeof value === 'number') return value === target;
      return text === cond.value.trim().toLocaleLowerCase();
    }
    case 'greaterThan': {
      const target = parseConditionNumber(cond.value);
      return target !== null && typeof value === 'number' && value > target;
    }
    case 'lessThan': {
      const target = parseConditionNumber(cond.value);
      return target !== null && typeof value === 'number' && value < target;
    }
    case 'between': {
      const a = parseConditionNumber(cond.min);
      const b = parseConditionNumber(cond.max);
      if (a === null || b === null || typeof value !== 'number') return false;
      return value >= Math.min(a, b) && value <= Math.max(a, b);
    }
  }
}

export function columnPasses(cf: ColumnFilter, value: CellValue, display: string): boolean {
  if (cf.kind === 'values') return !cf.hiddenValues.includes(display);
  return conditionMatches(cf.condition, value, display);
}

function cellInfo(sheet: Sheet, evaluator: Evaluator, sheetIndex: number, r: number, c: number) {
  const key = cellToA1(r, c);
  const value = evaluator.getCellValue(key, sheetIndex);
  return { value, display: formatDisplayValue(value, sheet.cells[key]?.fmt) };
}

/** คำนวณแถวที่ต้องซ่อนจากตัวกรองของชีต (ทุกคอลัมน์ต้องผ่านแบบ "และ") */
export function computeHiddenRows(sheet: Sheet, evaluator: Evaluator, sheetIndex: number): number[] {
  const filter = sheet.filter;
  if (!filter) return [];
  const entries = Object.entries(filter.columns).map(([col, cf]) => [Number(col), cf] as const);
  if (entries.length === 0) return [];

  const hidden: number[] = [];
  for (let r = filter.r1 + 1; r <= filter.r2; r++) {
    for (const [col, cf] of entries) {
      const { value, display } = cellInfo(sheet, evaluator, sheetIndex, r, col);
      if (!columnPasses(cf, value, display)) {
        hidden.push(r);
        break;
      }
    }
  }
  return hidden;
}

export interface UniqueValue {
  display: string;
  count: number;
}

/** ค่าไม่ซ้ำ (ค่าที่แสดง) ของคอลัมน์ในช่วงข้อมูล พร้อมจำนวนครั้ง เรียงแบบภาษาไทย ค่าว่างไว้ท้าย */
export function uniqueColumnValues(
  sheet: Sheet,
  evaluator: Evaluator,
  sheetIndex: number,
  col: number
): UniqueValue[] {
  const filter = sheet.filter;
  if (!filter) return [];
  const counts = new Map<string, { count: number; value: CellValue }>();
  for (let r = filter.r1 + 1; r <= filter.r2; r++) {
    const { value, display } = cellInfo(sheet, evaluator, sheetIndex, r, col);
    const entry = counts.get(display);
    if (entry) entry.count++;
    else counts.set(display, { count: 1, value });
  }

  const collator = new Intl.Collator('th', { sensitivity: 'accent', numeric: true });
  return Array.from(counts.entries())
    .sort(([da, a], [db, b]) => {
      if (da === '' || db === '') return da === db ? 0 : da === '' ? 1 : -1;
      if (typeof a.value === 'number' && typeof b.value === 'number') return a.value - b.value;
      if (typeof a.value === 'number') return -1;
      if (typeof b.value === 'number') return 1;
      return collator.compare(da, db);
    })
    .map(([display, { count }]) => ({ display, count }));
}

// --- ปรับตัวกรองและแถวที่ซ่อนเมื่อแทรก/ลบแถว-คอลัมน์ ---

export function adjustFilterOnInsertRows(sheet: Sheet, at: number, count: number) {
  if (sheet.hiddenRows) {
    sheet.hiddenRows = sheet.hiddenRows.map((r) => (r >= at ? r + count : r));
  }
  const f = sheet.filter;
  if (!f) return;
  if (at <= f.r1) {
    f.r1 += count;
    f.r2 += count;
  } else if (at <= f.r2) {
    f.r2 += count;
  }
}

export function adjustFilterOnDeleteRows(sheet: Sheet, at: number, count: number) {
  const end = at + count - 1;
  const f = sheet.filter;

  if (f && f.r1 >= at && f.r1 <= end) {
    // แถวหัวตารางถูกลบ: เอาตัวกรองออกและแสดงทุกแถว (กันแถวที่ซ่อนค้างผิดที่)
    delete sheet.filter;
    delete sheet.hiddenRows;
    return;
  }

  if (sheet.hiddenRows) {
    sheet.hiddenRows = sheet.hiddenRows
      .filter((r) => r < at || r > end)
      .map((r) => (r > end ? r - count : r));
    if (sheet.hiddenRows.length === 0) delete sheet.hiddenRows;
  }

  if (!f) return;
  if (at > f.r2) return;
  if (end < f.r1) {
    f.r1 -= count;
    f.r2 -= count;
    return;
  }
  // ลบแถวข้อมูลบางส่วนภายในช่วง
  const removedInside = Math.min(end, f.r2) - Math.max(at, f.r1 + 1) + 1;
  f.r2 -= Math.max(0, removedInside);
}

export function adjustFilterOnInsertCols(sheet: Sheet, at: number, count: number) {
  const f = sheet.filter;
  if (!f) return;
  const shiftCols = (from: number) => {
    const next: Record<number, ColumnFilter> = {};
    for (const [k, v] of Object.entries(f.columns)) {
      const c = Number(k);
      next[c >= from ? c + count : c] = v;
    }
    f.columns = next;
  };
  if (at <= f.c1) {
    f.c1 += count;
    f.c2 += count;
    shiftCols(at);
  } else if (at <= f.c2) {
    f.c2 += count;
    shiftCols(at);
  }
}

/**
 * คืน true ถ้าคอลัมน์ที่ถูกลบมีเงื่อนไขกรองอยู่ (ผู้เรียกควรคำนวณแถวที่ซ่อนใหม่)
 */
export function adjustFilterOnDeleteCols(sheet: Sheet, at: number, count: number): boolean {
  const f = sheet.filter;
  if (!f) return false;
  const end = at + count - 1;

  if (at <= f.c1 && end >= f.c2) {
    delete sheet.filter;
    delete sheet.hiddenRows;
    return false;
  }

  let removedCriteria = false;
  const next: Record<number, ColumnFilter> = {};
  for (const [k, v] of Object.entries(f.columns)) {
    const c = Number(k);
    if (c >= at && c <= end) {
      removedCriteria = true;
      continue;
    }
    next[c > end ? c - count : c] = v;
  }
  f.columns = next;

  if (at <= f.c2) {
    // จำนวนคอลัมน์ที่ถูกลบซึ่งอยู่ก่อนหรือที่ตำแหน่ง x
    const deletedUpTo = (x: number) => Math.max(0, Math.min(end, x) - at + 1);
    const c1Deleted = f.c1 >= at && f.c1 <= end;
    f.c1 = c1Deleted ? at : f.c1 - deletedUpTo(f.c1 - 1);
    f.c2 = f.c2 - deletedUpTo(f.c2);
  }
  return removedCriteria;
}
