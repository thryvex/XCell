import { describe, it, expect, beforeEach } from 'vitest';
import { findMatches, nextMatchIndex, DEFAULT_SEARCH_OPTIONS, SearchOptions } from '../../src/core/search';
import { Evaluator } from '../../src/core/evaluator';
import { Workbook, createEmptySheet } from '../../src/core/model';
import { useSheetStore } from '../../src/ui/store';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';

function wb(): Workbook {
  const s1 = createEmptySheet('ชีต1');
  s1.cells['A1'] = { raw: 'ค่าอาหาร' };
  s1.cells['A2'] = { raw: 'อาหารเช้า' };
  s1.cells['A3'] = { raw: 'Apple' };
  s1.cells['B1'] = { raw: '1234', fmt: 'baht' };
  s1.cells['B2'] = { raw: '=B1*2', fmt: 'baht' };
  s1.cells['B3'] = { raw: '=SUM(B1:B2)' };
  s1.cells['C5'] = { raw: 'apple' };
  const s2 = createEmptySheet('ชีต2');
  s2.cells['A1'] = { raw: 'อาหารกลางวัน' };
  return { sheets: [s1, s2], activeSheet: 0 };
}

const opts = (o: Partial<SearchOptions> = {}): SearchOptions => ({ ...DEFAULT_SEARCH_OPTIONS, ...o });
const keys = (ms: { sheetIndex: number; r: number; c: number }[]) =>
  ms.map((m) => `${m.sheetIndex}:${String.fromCharCode(65 + m.c)}${m.r + 1}`);

describe('findMatches (ค้นหา)', () => {
  let book: Workbook;
  let ev: Evaluator;
  beforeEach(() => {
    book = wb();
    ev = new Evaluator(book);
  });

  it('finds Thai text in display values, ordered by row then column', () => {
    expect(keys(findMatches(book, ev, 'อาหาร', opts(), 0))).toEqual(['0:A1', '0:A2']);
  });

  it('matches formatted numbers by what is displayed', () => {
    expect(keys(findMatches(book, ev, '฿1,234.00', opts(), 0))).toEqual(['0:B1']);
    expect(keys(findMatches(book, ev, '2,468', opts(), 0))).toEqual(['0:B2']);
    expect(findMatches(book, ev, '1234', opts(), 0)).toEqual([]); // ค่าที่แสดงคือ ฿1,234.00
  });

  it('searches formula text only when includeFormulas is on', () => {
    expect(findMatches(book, ev, 'SUM', opts(), 0)).toEqual([]);
    expect(keys(findMatches(book, ev, 'SUM', opts({ includeFormulas: true }), 0))).toEqual(['0:B3']);
    // ยังเจอค่าที่แสดงด้วยเมื่อเปิดค้นในสูตร
    expect(keys(findMatches(book, ev, '3702', opts({ includeFormulas: true }), 0))).toEqual(['0:B3']);
  });

  it('is case-insensitive by default and case-sensitive on request', () => {
    expect(keys(findMatches(book, ev, 'apple', opts(), 0))).toEqual(['0:A3', '0:C5']);
    expect(keys(findMatches(book, ev, 'apple', opts({ matchCase: true }), 0))).toEqual(['0:C5']);
  });

  it('whole-cell match', () => {
    expect(keys(findMatches(book, ev, 'อาหาร', opts({ wholeCell: true }), 0))).toEqual([]);
    expect(keys(findMatches(book, ev, 'ค่าอาหาร', opts({ wholeCell: true }), 0))).toEqual(['0:A1']);
  });

  it('searches all sheets when asked (and only the active one otherwise)', () => {
    expect(keys(findMatches(book, ev, 'กลางวัน', opts(), 0))).toEqual([]);
    expect(keys(findMatches(book, ev, 'อาหาร', opts({ allSheets: true }), 0))).toEqual([
      '0:A1',
      '0:A2',
      '1:A1',
    ]);
  });

  it('skips rows hidden by a filter and hidden sheets', () => {
    book.sheets[0].hiddenRows = [1];
    expect(keys(findMatches(book, ev, 'อาหาร', opts(), 0))).toEqual(['0:A1']);
    book.sheets[1].hidden = 'hidden';
    expect(keys(findMatches(book, ev, 'อาหาร', opts({ allSheets: true }), 0))).toEqual(['0:A1']);
  });

  it('empty query finds nothing', () => {
    expect(findMatches(book, ev, '', opts(), 0)).toEqual([]);
  });

  it('nextMatchIndex wraps around in both directions', () => {
    const ms = findMatches(book, ev, 'อาหาร', opts({ allSheets: true }), 0);
    expect(nextMatchIndex(ms, { sheetIndex: 0, r: 0, c: 0 }, 1)).toBe(1);
    expect(nextMatchIndex(ms, { sheetIndex: 0, r: 0, c: 0 }, 1, true)).toBe(0);
    expect(nextMatchIndex(ms, { sheetIndex: 1, r: 0, c: 0 }, 1)).toBe(0);
    expect(nextMatchIndex(ms, { sheetIndex: 0, r: 0, c: 0 }, -1)).toBe(2);
  });
});

describe('ค้นหาใน store', () => {
  const S = () => useSheetStore.getState();
  beforeEach(() => {
    S().newWorkbook();
    S().loadTemplate(wb(), 'ทดสอบ');
    S().setActiveCell(0, 0);
  });

  it('typing jumps to the first match from the current cell; next/prev move and wrap', () => {
    S().openSearch();
    S().setSearchQuery('อาหาร');
    expect(S().search.matches).toHaveLength(2);
    expect(S().search.current).toBe(0);
    expect(S().activeCell).toEqual({ r: 0, c: 0 });
    S().searchNext();
    expect(S().activeCell).toEqual({ r: 1, c: 0 });
    S().searchNext();
    expect(S().activeCell).toEqual({ r: 0, c: 0 });
    S().searchPrev();
    expect(S().search.current).toBe(1);
  });

  it('jumps to another sheet when searching all sheets', () => {
    S().openSearch();
    S().setSearchOption('allSheets', true);
    S().setSearchQuery('กลางวัน');
    expect(S().activeSheetIndex).toBe(1);
    expect(S().activeCell).toEqual({ r: 0, c: 0 });
  });

  it('not found leaves current at -1 without changing the selection', () => {
    S().setActiveCell(4, 4);
    S().openSearch();
    S().setSearchQuery('ไม่มีคำนี้');
    expect(S().search.matches).toEqual([]);
    expect(S().search.current).toBe(-1);
    expect(S().activeCell).toEqual({ r: 4, c: 4 });
  });

  it('closing keeps the found cell selected; search is not in undo and not saved', async () => {
    S().openSearch();
    S().setSearchQuery('Apple');
    S().closeSearch();
    expect(S().search.isOpen).toBe(false);
    expect(S().activeCell).toEqual({ r: 2, c: 0 });
    expect(S().history.canUndo()).toBe(false);
    expect(S().isDirty).toBe(false);

    const reloaded = await importFromXlsxBuffer(await exportToXlsxBuffer(S().workbook));
    expect(JSON.stringify(reloaded)).not.toContain('Apple"]');
    expect(Object.keys(reloaded.sheets[0])).not.toContain('search');
  });
});
