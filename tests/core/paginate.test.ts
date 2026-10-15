import { describe, it, expect } from 'vitest';
import {
  paginateSheet,
  printExtent,
  pageGeometry,
  buildPrintJob,
  splitIntoChunks,
  PX_TO_MM,
} from '../../src/core/paginate';
import { createEmptySheet, Sheet, DEFAULT_PRINT_SETTINGS, PrintSettings } from '../../src/core/model';
import { cellToA1 } from '../../src/core/address';

function filled(rows: number, cols: number): Sheet {
  const sheet = createEmptySheet('ชีต1');
  for (let r = 0; r < rows; r++) for (let c = 0; c < cols; c++) sheet.cells[cellToA1(r, c)] = { raw: 'x' };
  return sheet;
}

const s = (o: Partial<PrintSettings> = {}): PrintSettings => ({ ...DEFAULT_PRINT_SETTINGS, ...o });
const AT100 = s({ scaleMode: '100' });
const firstLast = (a: number[]) => [a[0], a[a.length - 1]];

/*
 * คำนวณมือ (A4 แนวตั้ง ขอบปกติ มีเลขหน้าท้ายกระดาษ 8 มม.):
 *   พื้นที่กว้าง = 210 − 17.78×2 = 174.44 มม., สูง = 297 − 19.05×2 − 8 = 250.9 มม.
 *   คอลัมน์ 110px = 29.104 มม. → 5 คอลัมน์/หน้า (6 × 29.104 = 174.63 > 174.44)
 *   แถว 28px = 7.408 มม. → 33 แถว/หน้า (34 × 7.408 = 251.9 > 250.9)
 * A4 แนวนอน: กว้าง 261.44 → 8 คอลัมน์, สูง 163.9 → 22 แถว
 */
describe('pageGeometry', () => {
  it('matches the hand calculation', () => {
    const g = pageGeometry(AT100);
    expect(g.contentW).toBeCloseTo(174.44, 2);
    expect(g.contentH).toBeCloseTo(250.9, 2);
    expect(110 * PX_TO_MM).toBeCloseTo(29.104, 3);
    expect(28 * PX_TO_MM).toBeCloseTo(7.408, 3);
    const land = pageGeometry(s({ orientation: 'landscape' }));
    expect([land.pageW, land.pageH]).toEqual([297, 210]);
    expect(pageGeometry(s({ paper: 'A5', header: true, pageNumbers: false })).contentH).toBeCloseTo(210 - 38.1 - 8, 2);
  });
});

describe('paginate: จำนวนหน้าเทียบคำนวณมือ', () => {
  it('1) small table → 1 page', () => {
    const sheet = filled(10, 3);
    expect(paginateSheet(sheet, printExtent(sheet), AT100).pages).toHaveLength(1);
  });

  it('2) 100 rows × 3 cols at 100% → 4 pages (33+33+33+1)', () => {
    const sheet = filled(100, 3);
    const { pages } = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(pages.map((p) => p.rows.length)).toEqual([33, 33, 33, 1]);
  });

  it('3) wide table at 100% splits by width too, down first then right (40×12 → 6 pages)', () => {
    const sheet = filled(40, 12);
    const { pages } = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(pages).toHaveLength(6);
    expect(pages.map((p) => [firstLast(p.rows), firstLast(p.cols)])).toEqual([
      [[0, 32], [0, 4]],
      [[33, 39], [0, 4]],
      [[0, 32], [5, 9]],
      [[33, 39], [5, 9]],
      [[0, 32], [10, 11]],
      [[33, 39], [10, 11]],
    ]);
  });

  it('4) fit to width shrinks 12 columns onto one page width (100 rows → 2 pages)', () => {
    const sheet = filled(100, 12);
    const { pages, scale } = paginateSheet(sheet, printExtent(sheet), s());
    expect(scale).toBeCloseTo(174.44 / (12 * 29.104), 3); // ≈ 0.4995
    // 7.408 × 0.4995 = 3.700 มม./แถว → 67 แถว/หน้า
    expect(pages.map((p) => p.rows.length)).toEqual([67, 33]);
    expect(pages.every((p) => p.cols.length === 12)).toBe(true);
  });

  it('5) fit to width never enlarges beyond 100%', () => {
    const sheet = filled(100, 2);
    const { pages, scale } = paginateSheet(sheet, printExtent(sheet), s());
    expect(scale).toBe(1);
    expect(pages).toHaveLength(4);
  });

  it('6) A4 landscape 50 rows × 10 cols → 6 pages (22,22,6 × 8,2)', () => {
    const sheet = filled(50, 10);
    const { pages } = paginateSheet(sheet, printExtent(sheet), s({ scaleMode: '100', orientation: 'landscape' }));
    expect(pages).toHaveLength(6);
    expect(pages.map((p) => p.rows.length)).toEqual([22, 22, 6, 22, 22, 6]);
    expect(pages.map((p) => p.cols.length)).toEqual([8, 8, 8, 2, 2, 2]);
  });

  it('custom scale 50% fits twice as many rows', () => {
    const sheet = filled(100, 3);
    const { pages } = paginateSheet(sheet, printExtent(sheet), s({ scaleMode: 'custom', scalePercent: 50 }));
    expect(pages.map((p) => p.rows.length)).toEqual([67, 33]);
  });
});

describe('paginate: ตรึง ผสาน ซ่อน ช่วง', () => {
  it('frozen rows repeat on every page and reduce the body height', () => {
    const sheet = filled(100, 3);
    sheet.frozenRows = 1;
    const { pages } = paginateSheet(sheet, printExtent(sheet), AT100);
    // (250.9 − 7.408) / 7.408 = 32.9 → 32 แถวเนื้อหา/หน้า, 99 แถว → 4 หน้า
    expect(pages.map((p) => p.rows.length)).toEqual([32, 32, 32, 3]);
    expect(pages.every((p) => p.titleRows.join() === '0')).toBe(true);
    expect(pages[0].rows[0]).toBe(1);
    const noRepeat = paginateSheet(sheet, printExtent(sheet), s({ scaleMode: '100', repeatFrozen: false }));
    expect(noRepeat.pages[1].titleRows).toEqual([]);
  });

  it('frozen columns repeat on every page across', () => {
    const sheet = filled(5, 12);
    sheet.frozenCols = 1;
    const { pages } = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(pages.every((p) => p.titleCols.join() === '0')).toBe(true);
    expect(pages[0].cols[0]).toBe(1);
  });

  it('a merged block is not cut: the break moves before it', () => {
    const sheet = filled(60, 2);
    sheet.merges = [{ r1: 30, c1: 0, r2: 35, c2: 0 }];
    const res = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(res.pages[0].rows.length).toBe(30);
    expect(res.pages[1].rows[0]).toBe(30);
    expect(res.mergeSplit).toBe(false);
  });

  it('a merged block taller than a page is cut and reported', () => {
    const sheet = filled(60, 2);
    sheet.merges = [{ r1: 0, c1: 0, r2: 40, c2: 0 }];
    const res = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(res.mergeSplit).toBe(true);
    expect(res.pages[0].rows.length).toBe(33);
  });

  it('merges are also kept together across columns at 100%', () => {
    const sheet = filled(3, 8);
    sheet.merges = [{ r1: 0, c1: 3, r2: 0, c2: 6 }]; // D1:G1 คร่อมจุดแบ่งหลังคอลัมน์ E
    const { pages } = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(pages[0].cols).toEqual([0, 1, 2]);
    expect(pages[1].cols).toEqual([3, 4, 5, 6, 7]);
  });

  it('rows hidden by a filter are skipped', () => {
    const sheet = filled(40, 2);
    sheet.hiddenRows = [1, 2, 3, 4, 5];
    const { pages } = paginateSheet(sheet, printExtent(sheet), AT100);
    expect(pages.map((p) => p.rows.length)).toEqual([33, 2]);
    expect(pages.flatMap((p) => p.rows)).not.toContain(3);
  });

  it('print range = A1 to the last cell with data or style', () => {
    const sheet = createEmptySheet();
    sheet.cells['B3'] = { raw: 'x' };
    sheet.cells['D2'] = { raw: '', style: { fill: '#ff0000' } };
    sheet.cells['Z99'] = { raw: '', fmt: 'baht' }; // รูปแบบตัวเลขอย่างเดียวไม่นับ
    expect(printExtent(sheet)).toEqual({ r1: 0, c1: 0, r2: 2, c2: 3 });
    expect(printExtent(createEmptySheet())).toBeNull();
  });

  it('empty sheet → no pages', () => {
    const sheet = createEmptySheet();
    expect(paginateSheet(sheet, printExtent(sheet), s()).pages).toEqual([]);
    expect(buildPrintJob({ sheets: [sheet], activeSheet: 0 }, 0, { r1: 0, c1: 0, r2: 0, c2: 0 }, s()).pages).toEqual([]);
  });

  it('selection prints only the selected range; all sheets skip hidden sheets', () => {
    const sheet = filled(50, 10);
    const sel = { r1: 1, c1: 1, r2: 4, c2: 2 };
    const job = buildPrintJob({ sheets: [sheet], activeSheet: 0 }, 0, sel, s({ what: 'selection' }));
    expect(job.pages).toHaveLength(1);
    expect(job.pages[0].rows).toEqual([1, 2, 3, 4]);
    expect(job.pages[0].cols).toEqual([1, 2]);

    const second = filled(5, 2);
    const hidden = filled(5, 2);
    hidden.hidden = 'hidden';
    const all = buildPrintJob({ sheets: [sheet, hidden, second], activeSheet: 0 }, 0, sel, s({ what: 'all' }));
    expect(all.pages.map((p) => p.sheetIndex)).toEqual([0, 2]); // พอดีความกว้าง: ชีตแรก 1 หน้า
  });

  it('splitIntoChunks keeps at least one item per chunk (row taller than a page)', () => {
    const { chunks } = splitIntoChunks([0, 1, 2], (i) => (i === 1 ? 500 : 10), 100, []);
    expect(chunks).toEqual([[0], [1], [2]]);
  });
});
