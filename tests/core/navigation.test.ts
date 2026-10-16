import { describe, it, expect } from 'vitest';
import {
  NavGrid,
  jumpToDataEdge,
  stepCell,
  moveVisibleRows,
  homePos,
  endPos,
  sheetStartPos,
  sheetEndPos,
  cellHasData,
  oppositeCorner,
} from '../../src/core/navigation';
import { parseA1, cellToA1 } from '../../src/core/address';
import { MergeRange } from '../../src/core/model';

/** ตาราง 20 แถว × 10 คอลัมน์ (A–J) จากรายการเซลที่มีข้อมูล */
function grid(cells: string[], opts: { hidden?: number[]; merges?: MergeRange[]; raw?: Record<string, string> } = {}): NavGrid {
  const set = new Set(cells);
  return {
    rowCount: 20,
    colCount: 10,
    hiddenRows: new Set(opts.hidden ?? []),
    merges: opts.merges ?? [],
    hasValue: (r, c) => {
      const key = cellToA1(r, c);
      if (opts.raw && key in opts.raw) return opts.raw[key] !== '';
      return set.has(key);
    },
  };
}
const at = (a1: string) => parseA1(a1)!;
const a1 = (p: { r: number; c: number }) => cellToA1(p.r, p.c);

describe('Ctrl+ลูกศร: ไปขอบของกลุ่มข้อมูล', () => {
  // คอลัมน์ A: A1:A4 มีข้อมูล, A5:A7 ว่าง, A8:A9 มีข้อมูล
  const g = grid(['A1', 'A2', 'A3', 'A4', 'A8', 'A9']);

  it('อยู่ในกลุ่มข้อมูล → เซลสุดท้ายของกลุ่ม', () => {
    expect(a1(jumpToDataEdge(g, at('A1'), 'down'))).toBe('A4');
    expect(a1(jumpToDataEdge(g, at('A2'), 'down'))).toBe('A4');
    expect(a1(jumpToDataEdge(g, at('A9'), 'up'))).toBe('A8');
  });

  it('อยู่ท้ายกลุ่ม/ในช่องว่าง → เซลแรกที่มีข้อมูลถัดไป', () => {
    expect(a1(jumpToDataEdge(g, at('A4'), 'down'))).toBe('A8');
    expect(a1(jumpToDataEdge(g, at('A6'), 'down'))).toBe('A8');
    expect(a1(jumpToDataEdge(g, at('A8'), 'up'))).toBe('A4');
  });

  it('ไม่มีข้อมูลต่อไป → ขอบตาราง (แถว/คอลัมน์สุดท้ายที่มีอยู่)', () => {
    expect(a1(jumpToDataEdge(g, at('A9'), 'down'))).toBe('A20');
    expect(a1(jumpToDataEdge(g, at('A1'), 'right'))).toBe('J1');
    expect(a1(jumpToDataEdge(g, at('C5'), 'up'))).toBe('C1');
    expect(a1(jumpToDataEdge(g, at('A1'), 'up'))).toBe('A1'); // อยู่ขอบแล้ว
  });

  it('แนวนอนทำงานแบบเดียวกัน', () => {
    const h = grid(['B2', 'C2', 'D2', 'G2']);
    expect(a1(jumpToDataEdge(h, at('A2'), 'right'))).toBe('B2');
    expect(a1(jumpToDataEdge(h, at('B2'), 'right'))).toBe('D2');
    expect(a1(jumpToDataEdge(h, at('D2'), 'right'))).toBe('G2');
    expect(a1(jumpToDataEdge(h, at('G2'), 'left'))).toBe('D2');
  });

  it('สูตรที่ได้ ="" นับเป็นมีข้อมูล; สี/รูปแบบตัวเลขอย่างเดียว (raw ว่าง) นับเป็นว่าง', () => {
    const h = grid([], { raw: { A1: '1', A2: '=""', A3: '', A4: '5' } });
    expect(cellHasData(h, 1, 0)).toBe(true);
    expect(cellHasData(h, 2, 0)).toBe(false);
    expect(a1(jumpToDataEdge(h, at('A1'), 'down'))).toBe('A2');
    expect(a1(jumpToDataEdge(h, at('A2'), 'down'))).toBe('A4');
  });

  it('ข้ามแถวที่ตัวกรองซ่อน', () => {
    // A1:A3 มีข้อมูล, A4 (มีข้อมูล) และ A5 ซ่อน, A6 มีข้อมูล → แถวที่มองเห็น A1,A2,A3,A6 ต่อกันเป็นกลุ่มเดียว
    const h = grid(['A1', 'A2', 'A3', 'A4', 'A6'], { hidden: [3, 4] });
    expect(a1(jumpToDataEdge(h, at('A1'), 'down'))).toBe('A6');
    expect(a1(jumpToDataEdge(h, at('A6'), 'up'))).toBe('A1');
    // ช่องว่างที่มองเห็นยังแบ่งกลุ่ม: ซ่อน A5 อย่างเดียว, A4 ว่าง
    const k = grid(['A1', 'A2', 'A3', 'A6'], { hidden: [4] });
    expect(a1(jumpToDataEdge(k, at('A1'), 'down'))).toBe('A3');
    expect(a1(jumpToDataEdge(k, at('A3'), 'down'))).toBe('A6');
    expect(a1(stepCell(h, at('A3'), 'down')!)).toBe('A6');
  });

  it('เซลผสานถือเป็นเซลเดียว ข้อมูลอยู่ที่เซลบนซ้าย และวิ่งข้ามทั้งก้อน', () => {
    // B3:B5 ผสาน (ข้อมูลที่ B3), B1:B2 มีข้อมูล, B6 มีข้อมูล
    const merges = [{ r1: 2, c1: 1, r2: 4, c2: 1 }];
    const h = grid(['B1', 'B2', 'B3', 'B6'], { merges });
    expect(cellHasData(h, 3, 1)).toBe(true); // B4 อยู่ในเซลผสานที่มีข้อมูล
    expect(a1(jumpToDataEdge(h, at('B1'), 'down'))).toBe('B6');
    expect(a1(stepCell(h, at('B3'), 'down')!)).toBe('B6');
    expect(a1(jumpToDataEdge(h, at('B6'), 'up'))).toBe('B1');

    // เซลผสานว่าง ข้ามทั้งก้อน แล้วหยุดที่ข้อมูลถัดไป; ถ้าหยุดในเซลผสานให้ไปเซลบนซ้าย
    const e = grid(['C1', 'C9'], { merges: [{ r1: 3, c1: 1, r2: 5, c2: 3 }] });
    expect(a1(jumpToDataEdge(e, at('C1'), 'down'))).toBe('C9');
    const f = grid(['C1', 'B4'], { merges: [{ r1: 3, c1: 1, r2: 5, c2: 3 }] });
    expect(a1(jumpToDataEdge(f, at('C1'), 'down'))).toBe('B4');
  });
});

describe('Home / End / Ctrl+Home / Ctrl+End', () => {
  const g = grid(['C3']);
  it('Home ไปคอลัมน์ A ของแถวเดิม, End ไปคอลัมน์สุดท้ายของตาราง', () => {
    expect(a1(homePos(g, at('E7')))).toBe('A7');
    expect(a1(endPos(g, at('E7')))).toBe('J7');
  });
  it('Ctrl+Home ไป A1, Ctrl+End ไปมุมขวาล่างของตารางที่มีอยู่', () => {
    expect(a1(sheetStartPos(g))).toBe('A1');
    expect(a1(sheetEndPos(g))).toBe('J20');
  });
  it('ข้ามแถวซ่อนที่หัว/ท้ายตาราง', () => {
    const h = grid([], { hidden: [0, 19, 18] });
    expect(a1(sheetStartPos(h))).toBe('A2');
    expect(a1(sheetEndPos(h))).toBe('J18');
  });
});

describe('PageUp / PageDown (นับแถวที่มองเห็น)', () => {
  it('เลื่อนตามจำนวนแถว หยุดที่แถวแรก/สุดท้าย และข้ามแถวซ่อน', () => {
    const g = grid([], { hidden: [5, 6] });
    expect(moveVisibleRows(g, 0, 4, 1)).toBe(4);
    expect(moveVisibleRows(g, 3, 3, 1)).toBe(8); // ข้ามแถว 6–7 (index 5, 6)
    expect(moveVisibleRows(g, 15, 10, 1)).toBe(19);
    expect(moveVisibleRows(g, 2, 10, -1)).toBe(0);
  });
});

describe('ขอบที่กำลังขยายเมื่อไม่ได้เก็บไว้', () => {
  it('ใช้มุมตรงข้ามกับเซลที่ใช้งานอยู่', () => {
    expect(oppositeCorner({ r: 1, c: 1 }, { r1: 1, c1: 1, r2: 5, c2: 3 })).toEqual({ r: 5, c: 3 });
    expect(oppositeCorner({ r: 5, c: 3 }, { r1: 1, c1: 1, r2: 5, c2: 3 })).toEqual({ r: 1, c: 1 });
  });
});
