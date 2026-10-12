import { Cell, Sheet } from './model';
import { shiftFormula } from './refshift';
import { cellToA1 } from './address';
import { normalizeFloat } from './evaluator';

export interface FillResult {
  cellsToSet: Record<string, Cell>;
}

/**
 * Computes filled cells when dragging fill handle from srcRange to targetRange
 */
export function computeFill(
  sheet: Sheet,
  src: { r1: number; c1: number; r2: number; c2: number },
  target: { r1: number; c1: number; r2: number; c2: number }
): FillResult {
  const cellsToSet: Record<string, Cell> = {};

  const srcHeight = src.r2 - src.r1 + 1;
  const srcWidth = src.c2 - src.c1 + 1;

  // Determine primary direction
  const isDown = target.r2 > src.r2;
  const isUp = target.r1 < src.r1;
  const isRight = target.c2 > src.c2;
  const isLeft = target.c1 < src.c1;

  if (isDown) {
    for (let c = src.c1; c <= src.c2; c++) {
      fillVerticalSlice(sheet, c, src.r1, src.r2, src.r2 + 1, target.r2, 'down', cellsToSet);
    }
  } else if (isUp) {
    for (let c = src.c1; c <= src.c2; c++) {
      fillVerticalSlice(sheet, c, src.r1, src.r2, target.r1, src.r1 - 1, 'up', cellsToSet);
    }
  } else if (isRight) {
    for (let r = src.r1; r <= src.r2; r++) {
      fillHorizontalSlice(sheet, r, src.c1, src.c2, src.c2 + 1, target.c2, 'right', cellsToSet);
    }
  } else if (isLeft) {
    for (let r = src.r1; r <= src.r2; r++) {
      fillHorizontalSlice(sheet, r, src.c1, src.c2, target.c1, src.c1 - 1, 'left', cellsToSet);
    }
  }

  return { cellsToSet };
}

function fillVerticalSlice(
  sheet: Sheet,
  c: number,
  srcR1: number,
  srcR2: number,
  targetR1: number,
  targetR2: number,
  dir: 'down' | 'up',
  output: Record<string, Cell>
) {
  const srcLen = srcR2 - srcR1 + 1;
  const srcCells: (Cell | undefined)[] = [];
  const numbers: (number | null)[] = [];
  const dates: (number | null)[] = [];
  let allNumbers = true;
  let allDates = true;
  let hasFormula = false;

  for (let r = srcR1; r <= srcR2; r++) {
    const key = cellToA1(r, c);
    const cell = sheet.cells[key];
    srcCells.push(cell);

    if (cell?.raw?.startsWith('=')) {
      hasFormula = true;
      allNumbers = false;
      allDates = false;
    } else if (cell?.fmt === 'date' && !isNaN(Number(cell.raw))) {
      dates.push(Number(cell.raw));
      allNumbers = false;
    } else {
      allDates = false;
      const num = cell?.raw ? Number(cell.raw.replace(/,/g, '')) : NaN;
      if (!isNaN(num) && cell?.raw !== '') {
        numbers.push(num);
      } else {
        allNumbers = false;
        numbers.push(null);
      }
    }
  }

  // 1. If slice contains formulas, shift formula per cell
  if (hasFormula || (!allNumbers && !allDates && srcLen === 1 && srcCells[0]?.raw?.startsWith('='))) {
    if (dir === 'down') {
      for (let r = targetR1; r <= targetR2; r++) {
        const offset = r - targetR1;
        const srcIdx = offset % srcLen;
        const srcCell = srcCells[srcIdx];
        const srcRow = srcR1 + srcIdx;
        const dr = r - srcRow;

        if (srcCell) {
          const raw = srcCell.raw.startsWith('=') ? shiftFormula(srcCell.raw, dr, 0) : srcCell.raw;
          output[cellToA1(r, c)] = { raw, fmt: srcCell.fmt, style: srcCell.style };
        }
      }
    } else {
      for (let r = targetR2; r >= targetR1; r--) {
        const offset = targetR2 - r;
        const srcIdx = srcLen - 1 - (offset % srcLen);
        const srcCell = srcCells[srcIdx];
        const srcRow = srcR1 + srcIdx;
        const dr = r - srcRow;

        if (srcCell) {
          const raw = srcCell.raw.startsWith('=') ? shiftFormula(srcCell.raw, dr, 0) : srcCell.raw;
          output[cellToA1(r, c)] = { raw, fmt: srcCell.fmt, style: srcCell.style };
        }
      }
    }
    return;
  }

  // 2. Dates sequence (single date increments by 1 day, multiple dates use delta)
  if (allDates && dates.length > 0) {
    const delta = dates.length >= 2 ? (dates[dates.length - 1]! - dates[0]!) / (dates.length - 1) : 1;
    const baseDate = dir === 'down' ? dates[dates.length - 1]! : dates[0]!;
    const templateCell = srcCells[0];

    if (dir === 'down') {
      for (let r = targetR1; r <= targetR2; r++) {
        const step = r - srcR2;
        const nextVal = Math.round(baseDate + step * delta);
        output[cellToA1(r, c)] = {
          raw: String(nextVal),
          fmt: 'date',
          style: templateCell?.style,
        };
      }
    } else {
      for (let r = targetR2; r >= targetR1; r--) {
        const step = srcR1 - r;
        const nextVal = Math.round(baseDate - step * delta);
        output[cellToA1(r, c)] = {
          raw: String(nextVal),
          fmt: 'date',
          style: templateCell?.style,
        };
      }
    }
    return;
  }

  // 3. Numbers sequence (>= 2 cells interpolate linear progression, 1 cell repeats)
  if (allNumbers && numbers.length > 0) {
    if (numbers.length >= 2) {
      const delta = (numbers[numbers.length - 1]! - numbers[0]!) / (numbers.length - 1);
      const baseNum = dir === 'down' ? numbers[numbers.length - 1]! : numbers[0]!;
      const templateCell = srcCells[0];

      if (dir === 'down') {
        for (let r = targetR1; r <= targetR2; r++) {
          const step = r - srcR2;
          const nextVal = normalizeFloat(baseNum + step * delta);
          output[cellToA1(r, c)] = {
            raw: String(nextVal),
            fmt: templateCell?.fmt,
            style: templateCell?.style,
          };
        }
      } else {
        for (let r = targetR2; r >= targetR1; r--) {
          const step = srcR1 - r;
          const nextVal = normalizeFloat(baseNum - step * delta);
          output[cellToA1(r, c)] = {
            raw: String(nextVal),
            fmt: templateCell?.fmt,
            style: templateCell?.style,
          };
        }
      }
      return;
    } else {
      // 1 single number: repeat
      const single = srcCells[0];
      for (let r = targetR1; r <= targetR2; r++) {
        if (single) {
          output[cellToA1(r, c)] = { ...single };
        }
      }
      return;
    }
  }

  // 4. Default: cycle patterns (text, mixed)
  if (dir === 'down') {
    for (let r = targetR1; r <= targetR2; r++) {
      const offset = r - targetR1;
      const srcCell = srcCells[offset % srcLen];
      if (srcCell) {
        output[cellToA1(r, c)] = { ...srcCell };
      }
    }
  } else {
    for (let r = targetR2; r >= targetR1; r--) {
      const offset = targetR2 - r;
      const srcCell = srcCells[srcLen - 1 - (offset % srcLen)];
      if (srcCell) {
        output[cellToA1(r, c)] = { ...srcCell };
      }
    }
  }
}

function fillHorizontalSlice(
  sheet: Sheet,
  r: number,
  srcC1: number,
  srcC2: number,
  targetC1: number,
  targetC2: number,
  dir: 'right' | 'left',
  output: Record<string, Cell>
) {
  const srcLen = srcC2 - srcC1 + 1;
  const srcCells: (Cell | undefined)[] = [];
  const numbers: (number | null)[] = [];
  let allNumbers = true;
  let hasFormula = false;

  for (let c = srcC1; c <= srcC2; c++) {
    const key = cellToA1(r, c);
    const cell = sheet.cells[key];
    srcCells.push(cell);

    if (cell?.raw?.startsWith('=')) {
      hasFormula = true;
      allNumbers = false;
    } else {
      const num = cell?.raw ? Number(cell.raw.replace(/,/g, '')) : NaN;
      if (!isNaN(num) && cell?.raw !== '') {
        numbers.push(num);
      } else {
        allNumbers = false;
        numbers.push(null);
      }
    }
  }

  if (hasFormula || (!allNumbers && srcLen === 1 && srcCells[0]?.raw?.startsWith('='))) {
    if (dir === 'right') {
      for (let c = targetC1; c <= targetC2; c++) {
        const offset = c - targetC1;
        const srcIdx = offset % srcLen;
        const srcCell = srcCells[srcIdx];
        const srcCol = srcC1 + srcIdx;
        const dc = c - srcCol;

        if (srcCell) {
          const raw = srcCell.raw.startsWith('=') ? shiftFormula(srcCell.raw, 0, dc) : srcCell.raw;
          output[cellToA1(r, c)] = { raw, fmt: srcCell.fmt, style: srcCell.style };
        }
      }
    } else {
      for (let c = targetC2; c >= targetC1; c--) {
        const offset = targetC2 - c;
        const srcIdx = srcLen - 1 - (offset % srcLen);
        const srcCell = srcCells[srcIdx];
        const srcCol = srcC1 + srcIdx;
        const dc = c - srcCol;

        if (srcCell) {
          const raw = srcCell.raw.startsWith('=') ? shiftFormula(srcCell.raw, 0, dc) : srcCell.raw;
          output[cellToA1(r, c)] = { raw, fmt: srcCell.fmt, style: srcCell.style };
        }
      }
    }
    return;
  }

  if (allNumbers && numbers.length >= 2) {
    const delta = (numbers[numbers.length - 1]! - numbers[0]!) / (numbers.length - 1);
    const baseNum = dir === 'right' ? numbers[numbers.length - 1]! : numbers[0]!;
    const templateCell = srcCells[0];

    if (dir === 'right') {
      for (let c = targetC1; c <= targetC2; c++) {
        const step = c - srcC2;
        const nextVal = normalizeFloat(baseNum + step * delta);
        output[cellToA1(r, c)] = {
          raw: String(nextVal),
          fmt: templateCell?.fmt,
          style: templateCell?.style,
        };
      }
    } else {
      for (let c = targetC2; c >= targetC1; c--) {
        const step = srcC1 - c;
        const nextVal = normalizeFloat(baseNum - step * delta);
        output[cellToA1(r, c)] = {
          raw: String(nextVal),
          fmt: templateCell?.fmt,
          style: templateCell?.style,
        };
      }
    }
    return;
  }

  // Single cell or cycling
  for (let c = targetC1; c <= targetC2; c++) {
    const offset = dir === 'right' ? c - targetC1 : targetC2 - c;
    const srcCell = srcCells[offset % srcLen];
    if (srcCell) {
      output[cellToA1(r, c)] = { ...srcCell };
    }
  }
}

/**
 * Ctrl+D: Fill Down (copies top row of selection downwards)
 */
export function fillDown(
  sheet: Sheet,
  selection: { r1: number; c1: number; r2: number; c2: number }
): Record<string, Cell> {
  const cellsToSet: Record<string, Cell> = {};
  if (selection.r1 === selection.r2) return cellsToSet;

  for (let c = selection.c1; c <= selection.c2; c++) {
    const topKey = cellToA1(selection.r1, c);
    const topCell = sheet.cells[topKey];
    if (!topCell) continue;

    for (let r = selection.r1 + 1; r <= selection.r2; r++) {
      const dr = r - selection.r1;
      const raw = topCell.raw.startsWith('=') ? shiftFormula(topCell.raw, dr, 0) : topCell.raw;
      cellsToSet[cellToA1(r, c)] = {
        raw,
        fmt: topCell.fmt,
        style: topCell.style,
      };
    }
  }

  return cellsToSet;
}

/**
 * Ctrl+R: Fill Right (copies leftmost column of selection rightwards)
 */
export function fillRight(
  sheet: Sheet,
  selection: { r1: number; c1: number; r2: number; c2: number }
): Record<string, Cell> {
  const cellsToSet: Record<string, Cell> = {};
  if (selection.c1 === selection.c2) return cellsToSet;

  for (let r = selection.r1; r <= selection.r2; r++) {
    const leftKey = cellToA1(r, selection.c1);
    const leftCell = sheet.cells[leftKey];
    if (!leftCell) continue;

    for (let c = selection.c1 + 1; c <= selection.c2; c++) {
      const dc = c - selection.c1;
      const raw = leftCell.raw.startsWith('=') ? shiftFormula(leftCell.raw, 0, dc) : leftCell.raw;
      cellsToSet[cellToA1(r, c)] = {
        raw,
        fmt: leftCell.fmt,
        style: leftCell.style,
      };
    }
  }

  return cellsToSet;
}
