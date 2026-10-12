import { tokenize } from './tokenizer';
import { parseA1, parseRange, cellToA1 } from './address';

/**
 * Shifts formula references when copying/dragging/filling by (dr, dc)
 * Absolute references ($) do not move along that dimension.
 */
export function shiftFormula(formula: string, dr: number, dc: number): string {
  if (!formula || !formula.startsWith('=')) return formula;

  const tokens = tokenize(formula, { includeWhitespace: true });
  const result: string[] = ['='];

  for (const t of tokens) {
    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (!p) {
        result.push(t.value);
        continue;
      }

      const newR = p.rowAbs ? p.r : p.r + dr;
      const newC = p.colAbs ? p.c : p.c + dc;

      if (newR < 0 || newC < 0) {
        result.push('#REF!');
      } else {
        result.push(cellToA1(newR, newC, p.colAbs, p.rowAbs));
      }
    } else if (t.type === 'RANGE') {
      const rng = parseRange(t.value);
      if (!rng) {
        result.push(t.value);
        continue;
      }

      const newR1 = rng.startRowAbs ? rng.r1 : rng.r1 + dr;
      const newC1 = rng.startColAbs ? rng.c1 : rng.c1 + dc;
      const newR2 = rng.endRowAbs ? rng.r2 : rng.r2 + dr;
      const newC2 = rng.endColAbs ? rng.c2 : rng.c2 + dc;

      if (newR1 < 0 || newC1 < 0 || newR2 < 0 || newC2 < 0) {
        result.push('#REF!');
      } else {
        const start = cellToA1(newR1, newC1, rng.startColAbs, rng.startRowAbs);
        const end = cellToA1(newR2, newC2, rng.endColAbs, rng.endRowAbs);
        result.push(`${start}:${end}`);
      }
    } else {
      result.push(t.raw);
    }
  }

  return result.join('');
}

/**
 * Adjusts formula references when inserting row(s) or col(s)
 */
export function shiftFormulaOnInsert(
  formula: string,
  axis: 'row' | 'col',
  index: number,
  count = 1
): string {
  if (!formula || !formula.startsWith('=')) return formula;

  const tokens = tokenize(formula, { includeWhitespace: true });
  const result: string[] = ['='];

  for (const t of tokens) {
    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (!p) {
        result.push(t.value);
        continue;
      }

      let newR = p.r;
      let newC = p.c;

      if (axis === 'row' && p.r >= index) {
        newR += count;
      } else if (axis === 'col' && p.c >= index) {
        newC += count;
      }

      result.push(cellToA1(newR, newC, p.colAbs, p.rowAbs));
    } else if (t.type === 'RANGE') {
      const rng = parseRange(t.value);
      if (!rng) {
        result.push(t.value);
        continue;
      }

      let { r1, c1, r2, c2 } = rng;

      if (axis === 'row') {
        if (r1 >= index) {
          r1 += count;
          r2 += count;
        } else if (r2 >= index) {
          // Span across insert point: range expands
          r2 += count;
        }
      } else if (axis === 'col') {
        if (c1 >= index) {
          c1 += count;
          c2 += count;
        } else if (c2 >= index) {
          // Span across insert point: range expands
          c2 += count;
        }
      }

      const start = cellToA1(r1, c1, rng.startColAbs, rng.startRowAbs);
      const end = cellToA1(r2, c2, rng.endColAbs, rng.endRowAbs);
      result.push(`${start}:${end}`);
    } else {
      result.push(t.raw);
    }
  }

  return result.join('');
}

/**
 * Adjusts formula references when deleting row(s) or col(s)
 */
export function shiftFormulaOnDelete(
  formula: string,
  axis: 'row' | 'col',
  index: number,
  count = 1
): string {
  if (!formula || !formula.startsWith('=')) return formula;

  const tokens = tokenize(formula, { includeWhitespace: true });
  const result: string[] = ['='];
  const deleteEnd = index + count - 1;

  for (const t of tokens) {
    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (!p) {
        result.push(t.raw);
        continue;
      }

      if (axis === 'row') {
        if (p.r >= index && p.r <= deleteEnd) {
          result.push('#REF!');
        } else if (p.r > deleteEnd) {
          result.push(cellToA1(p.r - count, p.c, p.colAbs, p.rowAbs));
        } else {
          result.push(cellToA1(p.r, p.c, p.colAbs, p.rowAbs));
        }
      } else {
        if (p.c >= index && p.c <= deleteEnd) {
          result.push('#REF!');
        } else if (p.c > deleteEnd) {
          result.push(cellToA1(p.r, p.c - count, p.colAbs, p.rowAbs));
        } else {
          result.push(cellToA1(p.r, p.c, p.colAbs, p.rowAbs));
        }
      }
    } else if (t.type === 'RANGE') {
      const rng = parseRange(t.value);
      if (!rng) {
        result.push(t.raw);
        continue;
      }

      let { r1, c1, r2, c2 } = rng;

      if (axis === 'row') {
        if (r1 >= index && r2 <= deleteEnd) {
          // Whole range deleted
          result.push('#REF!');
          continue;
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
          result.push('#REF!');
          continue;
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

      const start = cellToA1(r1, c1, rng.startColAbs, rng.startRowAbs);
      const end = cellToA1(r2, c2, rng.endColAbs, rng.endRowAbs);
      result.push(`${start}:${end}`);
    } else {
      result.push(t.raw);
    }
  }

  return result.join('');
}
