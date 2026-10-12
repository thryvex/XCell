import { Sheet } from './model';
import { cellToA1 } from './address';

/**
 * Automatically detects numeric range above or to the left of active cell (for Σ button)
 */
export function detectSigmaRange(sheet: Sheet, r: number, c: number): string | null {
  // 1. Check vertically above (r - 1 down to 0)
  let topR = r - 1;
  while (topR >= 0) {
    const key = cellToA1(topR, c);
    const cell = sheet.cells[key];
    if (!cell || cell.raw === '' || cell.raw === undefined) {
      break;
    }
    const isNum = !isNaN(Number(cell.raw.replace(/,/g, '')));
    const isFormula = cell.raw.startsWith('=');
    if (!isNum && !isFormula) {
      break;
    }
    topR--;
  }

  const startR = topR + 1;
  const endR = r - 1;
  if (startR <= endR) {
    if (startR === endR) {
      return cellToA1(startR, c);
    }
    return `${cellToA1(startR, c)}:${cellToA1(endR, c)}`;
  }

  // 2. Check horizontally to the left (c - 1 down to 0)
  let leftC = c - 1;
  while (leftC >= 0) {
    const key = cellToA1(r, leftC);
    const cell = sheet.cells[key];
    if (!cell || cell.raw === '' || cell.raw === undefined) {
      break;
    }
    const isNum = !isNaN(Number(cell.raw.replace(/,/g, '')));
    const isFormula = cell.raw.startsWith('=');
    if (!isNum && !isFormula) {
      break;
    }
    leftC--;
  }

  const startC = leftC + 1;
  const endC = c - 1;
  if (startC <= endC) {
    if (startC === endC) {
      return cellToA1(r, startC);
    }
    return `${cellToA1(r, startC)}:${cellToA1(r, endC)}`;
  }

  return null;
}
