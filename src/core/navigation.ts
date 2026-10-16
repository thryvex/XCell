import { MergeRange } from './model';

/**
 * การเลื่อนด้วยคีย์บอร์ดในตาราง (SPEC 5.1) — ฟังก์ชันล้วน ไม่ขึ้นกับ UI
 * อ้างอิงเอกสารคีย์ลัดของ Google Sheets และข้อตกลงกับเจ้าของโปรเจกต์ (ยังไม่ได้เทียบกับ Google Sheets จริง):
 * - เซลมีข้อมูล = มีค่าหรือสูตร (รวม ="") ; สี/รูปแบบตัวเลขอย่างเดียวนับเป็นว่าง
 * - แถวที่ตัวกรองซ่อนถูกข้ามทุกกรณี
 * - เซลผสานถือเป็นเซลเดียว (ข้อมูลอยู่ที่เซลบนซ้าย) วิ่งข้ามทั้งก้อน
 * - ขอบตาราง = แถว/คอลัมน์สุดท้ายที่มีอยู่ตอนนั้น (ไม่ขยายแถว)
 */
export interface NavGrid {
  rowCount: number;
  colCount: number;
  hiddenRows: ReadonlySet<number>;
  merges: readonly MergeRange[];
  hasValue: (r: number, c: number) => boolean;
}

export interface Pos {
  r: number;
  c: number;
}

export type Direction = 'up' | 'down' | 'left' | 'right';

const DELTA: Record<Direction, { dr: number; dc: number }> = {
  up: { dr: -1, dc: 0 },
  down: { dr: 1, dc: 0 },
  left: { dr: 0, dc: -1 },
  right: { dr: 0, dc: 1 },
};

function mergeAt(g: NavGrid, r: number, c: number): MergeRange | undefined {
  return g.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
}

/** เซลมีข้อมูลไหม (เซลในเซลผสานดูที่เซลบนซ้าย) */
export function cellHasData(g: NavGrid, r: number, c: number): boolean {
  const m = mergeAt(g, r, c);
  return m ? g.hasValue(m.r1, m.c1) : g.hasValue(r, c);
}

/** เซลบนซ้ายของเซลผสาน (หรือเซลเดิม) */
function normalize(g: NavGrid, p: Pos): Pos {
  const m = mergeAt(g, p.r, p.c);
  return m ? { r: m.r1, c: m.c1 } : p;
}

export function firstVisibleRow(g: NavGrid): number {
  for (let r = 0; r < g.rowCount; r++) if (!g.hiddenRows.has(r)) return r;
  return 0;
}

export function lastVisibleRow(g: NavGrid): number {
  for (let r = g.rowCount - 1; r >= 0; r--) if (!g.hiddenRows.has(r)) return r;
  return g.rowCount - 1;
}

/** เซลถัดไปหนึ่งช่องตามทิศ (ข้ามแถวซ่อน และข้ามเซลผสานทั้งก้อน) หรือ null ถ้าเลยขอบ */
export function stepCell(g: NavGrid, p: Pos, dir: Direction): Pos | null {
  const { dr, dc } = DELTA[dir];
  let r = p.r;
  let c = p.c;
  const m = mergeAt(g, r, c);
  if (m) {
    if (dr > 0) r = m.r2;
    if (dr < 0) r = m.r1;
    if (dc > 0) c = m.c2;
    if (dc < 0) c = m.c1;
  }
  r += dr;
  c += dc;
  if (dr !== 0) while (r >= 0 && r < g.rowCount && g.hiddenRows.has(r)) r += dr;
  if (r < 0 || r >= g.rowCount || c < 0 || c >= g.colCount) return null;
  return { r, c };
}

/**
 * Ctrl+ลูกศร: ไปขอบของกลุ่มข้อมูล
 * - เซลปัจจุบันและเซลถัดไปมีข้อมูล → เซลสุดท้ายที่มีข้อมูลต่อเนื่อง
 * - ไม่เช่นนั้น → เซลแรกที่มีข้อมูลถัดไป ถ้าไม่มีเลย → ขอบตาราง
 */
export function jumpToDataEdge(g: NavGrid, from: Pos, dir: Direction): Pos {
  let cur = from;
  let next = stepCell(g, cur, dir);
  if (!next) return normalize(g, cur);

  if (cellHasData(g, cur.r, cur.c) && cellHasData(g, next.r, next.c)) {
    while (next && cellHasData(g, next.r, next.c)) {
      cur = next;
      next = stepCell(g, cur, dir);
    }
    return normalize(g, cur);
  }

  cur = next;
  while (!cellHasData(g, cur.r, cur.c)) {
    const n = stepCell(g, cur, dir);
    if (!n) break;
    cur = n;
  }
  return normalize(g, cur);
}

/** เลื่อนแถวทีละ count แถวที่มองเห็น (PageUp/PageDown) หยุดที่แถวแรก/สุดท้ายที่มองเห็น */
export function moveVisibleRows(g: NavGrid, fromRow: number, count: number, dir: 1 | -1): number {
  let r = fromRow;
  let moved = 0;
  while (moved < count) {
    let n = r + dir;
    while (n >= 0 && n < g.rowCount && g.hiddenRows.has(n)) n += dir;
    if (n < 0 || n >= g.rowCount) break;
    r = n;
    moved++;
  }
  return r;
}

/** Home / End / Ctrl+Home / Ctrl+End */
export function homePos(_g: NavGrid, p: Pos): Pos {
  return { r: p.r, c: 0 };
}
export function endPos(g: NavGrid, p: Pos): Pos {
  return { r: p.r, c: g.colCount - 1 };
}
export function sheetStartPos(g: NavGrid): Pos {
  return { r: firstVisibleRow(g), c: 0 };
}
export function sheetEndPos(g: NavGrid): Pos {
  return { r: lastVisibleRow(g), c: g.colCount - 1 };
}

/** มุมของช่วงที่เลือกฝั่งตรงข้ามกับเซลที่ใช้งานอยู่ (ขอบที่กำลังขยาย ถ้าไม่ได้เก็บไว้) */
export function oppositeCorner(
  active: Pos,
  sel: { r1: number; c1: number; r2: number; c2: number }
): Pos {
  return {
    r: active.r === sel.r1 ? sel.r2 : sel.r1,
    c: active.c === sel.c1 ? sel.c2 : sel.c1,
  };
}
