import { Sheet, Cell, MergeRange } from './model';
import { parseA1, cellToA1 } from './address';
import { shiftFormulaOnInsert, shiftFormulaOnDelete } from './refshift';

export function insertRow(sheet: Sheet, atRow: number, count = 1): Sheet {
  const newCells: Record<string, Cell> = {};

  // 1. Shift cell positions and update formulas
  for (const [key, cell] of Object.entries(sheet.cells)) {
    const p = parseA1(key);
    if (!p) continue;

    const newR = p.r >= atRow ? p.r + count : p.r;
    const newKey = cellToA1(newR, p.c);

    let raw = cell.raw;
    if (raw && raw.startsWith('=')) {
      raw = shiftFormulaOnInsert(raw, 'row', atRow, count);
    }

    newCells[newKey] = {
      ...cell,
      raw,
    };
  }

  // 2. Adjust merges
  const newMerges: MergeRange[] = [];
  for (const m of sheet.merges) {
    let { r1, c1, r2, c2 } = m;
    if (r1 >= atRow) {
      r1 += count;
      r2 += count;
    } else if (r2 >= atRow) {
      r2 += count;
    }
    newMerges.push({ r1, c1, r2, c2 });
  }

  // 3. Shift row heights
  const newHeights: Record<number, number> = {};
  for (const [rStr, h] of Object.entries(sheet.rowHeights)) {
    const r = parseInt(rStr, 10);
    const newR = r >= atRow ? r + count : r;
    newHeights[newR] = h;
  }

  sheet.cells = newCells;
  sheet.merges = newMerges;
  sheet.rowHeights = newHeights;

  return sheet;
}

export function deleteRow(sheet: Sheet, atRow: number, count = 1): Sheet {
  const newCells: Record<string, Cell> = {};
  const deleteEnd = atRow + count - 1;

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const p = parseA1(key);
    if (!p) continue;

    if (p.r >= atRow && p.r <= deleteEnd) {
      // Cell is inside deleted row
      continue;
    }

    const newR = p.r > deleteEnd ? p.r - count : p.r;
    const newKey = cellToA1(newR, p.c);

    let raw = cell.raw;
    if (raw && raw.startsWith('=')) {
      raw = shiftFormulaOnDelete(raw, 'row', atRow, count);
    }

    newCells[newKey] = {
      ...cell,
      raw,
    };
  }

  const newMerges: MergeRange[] = [];
  for (const m of sheet.merges) {
    let { r1, c1, r2, c2 } = m;
    if (r1 >= atRow && r2 <= deleteEnd) {
      // Whole merge deleted
      continue;
    }

    if (r1 > deleteEnd) {
      r1 -= count;
      r2 -= count;
    } else if (r1 >= atRow && r1 <= deleteEnd) {
      r1 = atRow;
      r2 -= count;
    } else if (r2 >= atRow && r2 <= deleteEnd) {
      r2 = atRow - 1;
    } else if (r1 < atRow && r2 > deleteEnd) {
      r2 -= count;
    }

    if (r1 <= r2 && c1 <= c2 && (r1 !== r2 || c1 !== c2)) {
      newMerges.push({ r1, c1, r2, c2 });
    }
  }

  sheet.cells = newCells;
  sheet.merges = newMerges;
  return sheet;
}

export function insertCol(sheet: Sheet, atCol: number, count = 1): Sheet {
  const newCells: Record<string, Cell> = {};

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const p = parseA1(key);
    if (!p) continue;

    const newC = p.c >= atCol ? p.c + count : p.c;
    const newKey = cellToA1(p.r, newC);

    let raw = cell.raw;
    if (raw && raw.startsWith('=')) {
      raw = shiftFormulaOnInsert(raw, 'col', atCol, count);
    }

    newCells[newKey] = {
      ...cell,
      raw,
    };
  }

  const newMerges: MergeRange[] = [];
  for (const m of sheet.merges) {
    let { r1, c1, r2, c2 } = m;
    if (c1 >= atCol) {
      c1 += count;
      c2 += count;
    } else if (c2 >= atCol) {
      c2 += count;
    }
    newMerges.push({ r1, c1, r2, c2 });
  }

  sheet.cells = newCells;
  sheet.merges = newMerges;
  return sheet;
}

export function deleteCol(sheet: Sheet, atCol: number, count = 1): Sheet {
  const newCells: Record<string, Cell> = {};
  const deleteEnd = atCol + count - 1;

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const p = parseA1(key);
    if (!p) continue;

    if (p.c >= atCol && p.c <= deleteEnd) {
      continue;
    }

    const newC = p.c > deleteEnd ? p.c - count : p.c;
    const newKey = cellToA1(p.r, newC);

    let raw = cell.raw;
    if (raw && raw.startsWith('=')) {
      raw = shiftFormulaOnDelete(raw, 'col', atCol, count);
    }

    newCells[newKey] = {
      ...cell,
      raw,
    };
  }

  const newMerges: MergeRange[] = [];
  for (const m of sheet.merges) {
    let { r1, c1, r2, c2 } = m;
    if (c1 >= atCol && c2 <= deleteEnd) {
      continue;
    }

    if (c1 > deleteEnd) {
      c1 -= count;
      c2 -= count;
    } else if (c1 >= atCol && c1 <= deleteEnd) {
      c1 = atCol;
      c2 -= count;
    } else if (c2 >= atCol && c2 <= deleteEnd) {
      c2 = atCol - 1;
    } else if (c1 < atCol && c2 > deleteEnd) {
      c2 -= count;
    }

    if (r1 <= r2 && c1 <= c2 && (r1 !== r2 || c1 !== c2)) {
      newMerges.push({ r1, c1, r2, c2 });
    }
  }

  sheet.cells = newCells;
  sheet.merges = newMerges;
  return sheet;
}

/**
 * Counts how many formulas in the sheet will be broken into #REF! if the row is deleted.
 * Cells residing inside the deleted row are excluded.
 */
export function countFormulasReferencingRow(sheet: Sheet, atRow: number, count = 1): number {
  let affectedCount = 0;
  const deleteEnd = atRow + count - 1;

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const p = parseA1(key);
    if (!p) continue;
    if (p.r >= atRow && p.r <= deleteEnd) continue;

    if (cell.raw && cell.raw.startsWith('=')) {
      const shifted = shiftFormulaOnDelete(cell.raw, 'row', atRow, count);
      if (shifted.includes('#REF!') && !cell.raw.includes('#REF!')) {
        affectedCount++;
      }
    }
  }

  return affectedCount;
}

/**
 * Counts how many formulas in the sheet will be broken into #REF! if the column is deleted.
 * Cells residing inside the deleted column are excluded.
 */
export function countFormulasReferencingCol(sheet: Sheet, atCol: number, count = 1): number {
  let affectedCount = 0;
  const deleteEnd = atCol + count - 1;

  for (const [key, cell] of Object.entries(sheet.cells)) {
    const p = parseA1(key);
    if (!p) continue;
    if (p.c >= atCol && p.c <= deleteEnd) continue;

    if (cell.raw && cell.raw.startsWith('=')) {
      const shifted = shiftFormulaOnDelete(cell.raw, 'col', atCol, count);
      if (shifted.includes('#REF!') && !cell.raw.includes('#REF!')) {
        affectedCount++;
      }
    }
  }

  return affectedCount;
}
