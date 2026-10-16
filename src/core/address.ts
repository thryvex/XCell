export interface CellCoord {
  r: number;
  c: number;
  rowAbs?: boolean;
  colAbs?: boolean;
}

export interface RangeCoord {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  startColAbs?: boolean;
  startRowAbs?: boolean;
  endColAbs?: boolean;
  endRowAbs?: boolean;
  /** อ้างทั้งคอลัมน์ เช่น A:A, $A:$C (SPEC 4.1) — r1 = 0, r2 = MAX_SHEET_ROWS - 1 */
  wholeColumns?: boolean;
}

/** จำนวนแถวสูงสุดของ .xlsx (ใช้เป็นขอบล่างของการอ้างทั้งคอลัมน์ ห้ามวนทุกแถวตรงๆ) */
export const MAX_SHEET_ROWS = 1048576;

const COLUMN_PART = /^(\$?)([A-Za-z]{1,3})$/;

export function colToStr(col: number): string {
  if (col < 0) return '';
  let str = '';
  let c = col;
  while (c >= 0) {
    str = String.fromCharCode((c % 26) + 65) + str;
    c = Math.floor(c / 26) - 1;
  }
  return str;
}

export function strToCol(str: string): number {
  if (!str) return -1;
  let col = 0;
  const s = str.toUpperCase();
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) return -1;
    col = col * 26 + (code - 64);
  }
  return col - 1;
}

export function cellToA1(r: number, c: number, colAbs = false, rowAbs = false): string {
  return `${colAbs ? '$' : ''}${colToStr(c)}${rowAbs ? '$' : ''}${r + 1}`;
}

export function parseA1(ref: string): CellCoord | null {
  if (!ref) return null;
  const m = /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(ref.trim());
  if (!m) return null;

  const colAbs = m[1] === '$';
  const c = strToCol(m[2]);
  const rowAbs = m[3] === '$';
  const r = parseInt(m[4], 10) - 1;

  if (c < 0 || r < 0) return null;
  return { r, c, rowAbs, colAbs };
}

export function parseRange(rangeStr: string): RangeCoord | null {
  if (!rangeStr) return null;
  const parts = rangeStr.trim().split(':');
  if (parts.length === 1) {
    const p = parseA1(parts[0]);
    if (!p) return null;
    return {
      r1: p.r,
      c1: p.c,
      r2: p.r,
      c2: p.c,
      startColAbs: p.colAbs,
      startRowAbs: p.rowAbs,
      endColAbs: p.colAbs,
      endRowAbs: p.rowAbs,
    };
  }

  if (parts.length === 2) {
    // ทั้งคอลัมน์: A:A, $A:$C
    const k1 = COLUMN_PART.exec(parts[0].trim());
    const k2 = COLUMN_PART.exec(parts[1].trim());
    if (k1 && k2) {
      const a = strToCol(k1[2]);
      const b = strToCol(k2[2]);
      if (a < 0 || b < 0) return null;
      const first = a <= b;
      return {
        r1: 0,
        c1: Math.min(a, b),
        r2: MAX_SHEET_ROWS - 1,
        c2: Math.max(a, b),
        startColAbs: (first ? k1 : k2)[1] === '$',
        endColAbs: (first ? k2 : k1)[1] === '$',
        wholeColumns: true,
      };
    }
    const p1 = parseA1(parts[0]);
    const p2 = parseA1(parts[1]);
    if (!p1 || !p2) return null;
    return {
      r1: Math.min(p1.r, p2.r),
      c1: Math.min(p1.c, p2.c),
      r2: Math.max(p1.r, p2.r),
      c2: Math.max(p1.c, p2.c),
      startColAbs: p1.colAbs,
      startRowAbs: p1.rowAbs,
      endColAbs: p2.colAbs,
      endRowAbs: p2.rowAbs,
    };
  }

  return null;
}

export function rangeToA1(r1: number, c1: number, r2: number, c2: number): string {
  const minR = Math.min(r1, r2);
  const minC = Math.min(c1, c2);
  const maxR = Math.max(r1, r2);
  const maxC = Math.max(c1, c2);

  if (minR === maxR && minC === maxC) {
    return cellToA1(minR, minC);
  }
  return `${cellToA1(minR, minC)}:${cellToA1(maxR, maxC)}`;
}

export function isCellInRange(r: number, c: number, range: { r1: number; c1: number; r2: number; c2: number }): boolean {
  return r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2;
}

/**
 * Cycles reference between A1 -> $A$1 -> A$1 -> $A1 -> A1 (F4 key behavior)
 */
export function cycleReference(ref: string): string {
  if (!ref) return ref;
  if (ref.includes(':')) {
    const parts = ref.split(':');
    if (parts.length === 2) {
      return `${cycleReference(parts[0])}:${cycleReference(parts[1])}`;
    }
    return ref;
  }

  const p = parseA1(ref);
  if (!p) return ref;

  if (!p.colAbs && !p.rowAbs) {
    // A1 -> $A$1
    return cellToA1(p.r, p.c, true, true);
  } else if (p.colAbs && p.rowAbs) {
    // $A$1 -> A$1
    return cellToA1(p.r, p.c, false, true);
  } else if (!p.colAbs && p.rowAbs) {
    // A$1 -> $A1
    return cellToA1(p.r, p.c, true, false);
  } else {
    // $A1 -> A1
    return cellToA1(p.r, p.c, false, false);
  }
}
