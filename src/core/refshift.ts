import { tokenize, Token } from './tokenizer';
import { parseA1, parseRange, cellToA1, RangeCoord } from './address';
import { formatSheetReference, sameSheetName } from './sheetName';

/**
 * ขอบเขตของการแทรก/ลบแถว-คอลัมน์ข้ามชีต:
 * - formulaSheet: ชีตที่สูตรนี้อยู่ (reference ที่ไม่มีชื่อชีตหมายถึงชีตนี้)
 * - targetSheet: ชีตที่ถูกแทรก/ลบแถว-คอลัมน์
 * ปรับเฉพาะ reference ที่ชี้ไปยัง targetSheet เท่านั้น
 * ถ้าไม่ส่ง scope จะปรับทุก reference (พฤติกรรมเดิมแบบชีตเดียว)
 */
export interface RefScope {
  formulaSheet: string;
  targetSheet: string;
}

type RefToken = Token & { type: 'CELL_REF' | 'RANGE' };

/** ส่วนนำหน้าชื่อชีตตามที่ผู้ใช้เขียนไว้ เช่น "'ชีต 2'!" หรือ "" */
function sheetPrefix(t: RefToken): string {
  if (t.sheet === undefined) return '';
  return t.raw.slice(0, t.raw.lastIndexOf('!') + 1);
}

/** ส่วนที่อยู่ตามที่ผู้ใช้เขียนไว้ เช่น "a1:$B$5" */
function rawAddress(t: RefToken): string {
  return t.sheet === undefined ? t.raw : t.raw.slice(t.raw.lastIndexOf('!') + 1);
}

/**
 * เขียนสูตรใหม่โดยแปลงทีละ reference (ผ่าน tokenizer ไม่ใช้ regex แทนข้อความ)
 * fn คืนข้อความใหม่ของ reference หรือ null เพื่อคงของเดิม
 */
function rewriteRefs(formula: string, fn: (t: RefToken) => string | null): string {
  if (!formula || !formula.startsWith('=')) return formula;

  const tokens = tokenize(formula, { includeWhitespace: true });
  const result: string[] = ['='];

  for (const t of tokens) {
    if (t.type === 'CELL_REF' || t.type === 'RANGE') {
      const replaced = fn(t as RefToken);
      result.push(replaced ?? t.raw);
    } else {
      result.push(t.raw);
    }
  }

  return result.join('');
}

function isInScope(t: RefToken, scope?: RefScope): boolean {
  if (!scope) return true;
  const refSheet = t.sheet ?? scope.formulaSheet;
  return sameSheetName(refSheet, scope.targetSheet);
}

function formatRange(rng: RangeCoord, r1: number, c1: number, r2: number, c2: number): string {
  const start = cellToA1(r1, c1, rng.startColAbs, rng.startRowAbs);
  const end = cellToA1(r2, c2, rng.endColAbs, rng.endRowAbs);
  return `${start}:${end}`;
}

/**
 * Shifts formula references when copying/dragging/filling by (dr, dc)
 * Absolute references ($) do not move along that dimension.
 * ชื่อชีตในสูตรคงเดิม; reference ที่ไม่มีชื่อชีตยังคงหมายถึงชีตที่สูตรไปอยู่
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  return rewriteRefs(formula, (t) => {
    const prefix = sheetPrefix(t);

    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (!p) return null;

      const newR = p.rowAbs ? p.r : p.r + dr;
      const newC = p.colAbs ? p.c : p.c + dc;

      if (newR < 0 || newC < 0) return '#REF!';
      return prefix + cellToA1(newR, newC, p.colAbs, p.rowAbs);
    }

    const rng = parseRange(t.value);
    if (!rng) return null;

    const newR1 = rng.startRowAbs ? rng.r1 : rng.r1 + dr;
    const newC1 = rng.startColAbs ? rng.c1 : rng.c1 + dc;
    const newR2 = rng.endRowAbs ? rng.r2 : rng.r2 + dr;
    const newC2 = rng.endColAbs ? rng.c2 : rng.c2 + dc;

    if (newR1 < 0 || newC1 < 0 || newR2 < 0 || newC2 < 0) return '#REF!';
    return prefix + formatRange(rng, newR1, newC1, newR2, newC2);
  });
}

/**
 * Adjusts formula references when inserting row(s) or col(s)
 */
export function shiftFormulaOnInsert(
  formula: string,
  axis: 'row' | 'col',
  index: number,
  count = 1,
  scope?: RefScope
): string {
  return rewriteRefs(formula, (t) => {
    if (!isInScope(t, scope)) return null;
    const prefix = sheetPrefix(t);

    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (!p) return null;

      let newR = p.r;
      let newC = p.c;

      if (axis === 'row' && p.r >= index) {
        newR += count;
      } else if (axis === 'col' && p.c >= index) {
        newC += count;
      }

      return prefix + cellToA1(newR, newC, p.colAbs, p.rowAbs);
    }

    const rng = parseRange(t.value);
    if (!rng) return null;

    let { r1, c1, r2, c2 } = rng;

    if (axis === 'row') {
      if (r1 >= index) {
        r1 += count;
        r2 += count;
      } else if (r2 >= index) {
        // Span across insert point: range expands
        r2 += count;
      }
    } else {
      if (c1 >= index) {
        c1 += count;
        c2 += count;
      } else if (c2 >= index) {
        // Span across insert point: range expands
        c2 += count;
      }
    }

    return prefix + formatRange(rng, r1, c1, r2, c2);
  });
}

/**
 * Adjusts formula references when deleting row(s) or col(s)
 */
export function shiftFormulaOnDelete(
  formula: string,
  axis: 'row' | 'col',
  index: number,
  count = 1,
  scope?: RefScope
): string {
  const deleteEnd = index + count - 1;

  return rewriteRefs(formula, (t) => {
    if (!isInScope(t, scope)) return null;
    const prefix = sheetPrefix(t);

    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (!p) return null;

      if (axis === 'row') {
        if (p.r >= index && p.r <= deleteEnd) return '#REF!';
        const newR = p.r > deleteEnd ? p.r - count : p.r;
        return prefix + cellToA1(newR, p.c, p.colAbs, p.rowAbs);
      }

      if (p.c >= index && p.c <= deleteEnd) return '#REF!';
      const newC = p.c > deleteEnd ? p.c - count : p.c;
      return prefix + cellToA1(p.r, newC, p.colAbs, p.rowAbs);
    }

    const rng = parseRange(t.value);
    if (!rng) return null;

    let { r1, c1, r2, c2 } = rng;

    if (axis === 'row') {
      if (r1 >= index && r2 <= deleteEnd) {
        // Whole range deleted
        return '#REF!';
      }

      if (r1 > deleteEnd) {
        r1 -= count;
        r2 -= count;
      } else if (r1 >= index && r1 <= deleteEnd) {
        // Top of range deleted, contract top
        r1 = index;
        r2 -= count;
      } else if (r2 >= index && r2 <= deleteEnd) {
        // Bottom of range deleted, contract bottom
        r2 = index - 1;
      } else if (r1 < index && r2 > deleteEnd) {
        // Deletion inside range, range contracts
        r2 -= count;
      }
    } else {
      if (c1 >= index && c2 <= deleteEnd) {
        // Whole range deleted
        return '#REF!';
      }

      if (c1 > deleteEnd) {
        c1 -= count;
        c2 -= count;
      } else if (c1 >= index && c1 <= deleteEnd) {
        c1 = index;
        c2 -= count;
      } else if (c2 >= index && c2 <= deleteEnd) {
        c2 = index - 1;
      } else if (c1 < index && c2 > deleteEnd) {
        c2 -= count;
      }
    }

    return prefix + formatRange(rng, r1, c1, r2, c2);
  });
}

/**
 * เปลี่ยนชื่อชีตในสูตร: reference ที่อ้าง oldName (ไม่สนตัวพิมพ์) เปลี่ยนเป็น newName
 * และใส่/ถอดเครื่องหมายคำพูดตามกติกาของ formatSheetReference
 */
export function renameSheetInFormula(formula: string, oldName: string, newName: string): string {
  return rewriteRefs(formula, (t) => {
    if (t.sheet === undefined || !sameSheetName(t.sheet, oldName)) return null;
    return `${formatSheetReference(newName)}!${rawAddress(t)}`;
  });
}

/**
 * ชีตถูกลบ: reference ที่ชี้ไปยังชีตนั้นกลายเป็น #REF!
 */
export function invalidateSheetInFormula(formula: string, sheetName: string): string {
  return rewriteRefs(formula, (t) => {
    if (t.sheet === undefined || !sameSheetName(t.sheet, sheetName)) return null;
    return '#REF!';
  });
}

/**
 * เขียนชื่อชีตในสูตรให้อยู่ในรูปมาตรฐาน (ใช้ตอนบันทึก .xlsx):
 * ใส่เครื่องหมายคำพูดเสมอ ยกเว้นชื่อที่เป็น A-Z / ตัวเลข / ขีดล่างล้วน
 */
export function normalizeSheetRefsInFormula(formula: string): string {
  return rewriteRefs(formula, (t) => {
    if (t.sheet === undefined) return null;
    return `${formatSheetReference(t.sheet)}!${rawAddress(t)}`;
  });
}
