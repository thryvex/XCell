import { describe, it, expect, beforeEach } from 'vitest';
import {
  compareSortValues,
  detectDataRegion,
  detectHeaderRow,
  formulasReferenceOtherRows,
  sortRange,
  rangeHasMerge,
} from '../../src/core/sort';
import { shiftFormula } from '../../src/core/refshift';
import { Evaluator } from '../../src/core/evaluator';
import { Sheet, Workbook, createEmptySheet, CellValue } from '../../src/core/model';
import { useSheetStore } from '../../src/ui/store';
import { th } from '../../src/i18n/th';

function column(sheet: Sheet, col: string, from: number, to: number): string[] {
  const out: string[] = [];
  for (let r = from; r <= to; r++) out.push(sheet.cells[`${col}${r}`]?.raw ?? '');
  return out;
}

function sortValues(values: CellValue[], dir: 'asc' | 'desc'): CellValue[] {
  return [...values].sort((a, b) => compareSortValues(a, b, dir));
}

describe('compareSortValues', () => {
  const mixed: CellValue[] = ['', 'ขนม', 42, '#DIV/0!', 'กล้วย', 3.5, '', 46000, 'มะม่วง'];

  it('A→Z: numbers/dates by value → Thai text → errors → blanks', () => {
    expect(sortValues(mixed, 'asc')).toEqual([3.5, 42, 46000, 'กล้วย', 'ขนม', 'มะม่วง', '#DIV/0!', '', '']);
  });

  it('Z→A reverses everything except blanks, which stay last', () => {
    expect(sortValues(mixed, 'desc')).toEqual(['#DIV/0!', 'มะม่วง', 'ขนม', 'กล้วย', 46000, 42, 3.5, '', '']);
  });

  it('text ignores case', () => {
    expect(compareSortValues('apple', 'APPLE', 'asc')).toBe(0);
    expect(sortValues(['banana', 'Apple', 'cherry'], 'asc')).toEqual(['Apple', 'banana', 'cherry']);
  });
});

describe('sortRange', () => {
  let sheet: Sheet;
  let wb: Workbook;

  beforeEach(() => {
    sheet = createEmptySheet('ชีต1');
    wb = { sheets: [sheet], activeSheet: 0 };
    const rows = [
      ['สินค้า', 'ราคา', 'รวม'],
      ['มะม่วง', '30', '=B2*2'],
      ['กล้วย', '10', '=B3*2'],
      ['ขนม', '', '=B4*2'],
      ['กล้วย', '5', '=B5*2'],
      ['แตงโม', '20', '=B6*2'],
    ];
    rows.forEach((row, r) =>
      row.forEach((raw, c) => {
        if (raw !== '') sheet.cells[`${'ABC'[c]}${r + 1}`] = { raw };
      })
    );
    sheet.cells['A3'].style = { bold: true, fill: '#ff0000' };
    sheet.cells['B3'].fmt = 'baht';
    sheet.cells['E1'] = { raw: '=B2' }; // นอกช่วง ชี้เข้ามาในช่วง
  });

  const range = { r1: 0, c1: 0, r2: 5, c2: 2 };

  it('detects the data region and the header row', () => {
    expect(detectDataRegion(sheet, 3, 1)).toEqual(range);
    expect(detectHeaderRow(sheet, new Evaluator(wb), 0, range)).toBe(true);
    expect(detectHeaderRow(sheet, new Evaluator(wb), 0, { ...range, r1: 1 })).toBe(false);
  });

  it('sorts numbers ascending with blanks last, keeping the header in place', () => {
    sortRange(sheet, new Evaluator(wb), 0, range, [{ col: 1, dir: 'asc' }], true);
    expect(column(sheet, 'A', 1, 6)).toEqual(['สินค้า', 'กล้วย', 'กล้วย', 'แตงโม', 'มะม่วง', 'ขนม']);
    expect(column(sheet, 'B', 1, 6)).toEqual(['ราคา', '5', '10', '20', '30', '']);
  });

  it('sorts descending with blanks still last', () => {
    sortRange(sheet, new Evaluator(wb), 0, range, [{ col: 1, dir: 'desc' }], true);
    expect(column(sheet, 'B', 2, 6)).toEqual(['30', '20', '10', '5', '']);
  });

  it('moves styles and number formats with their rows', () => {
    sortRange(sheet, new Evaluator(wb), 0, range, [{ col: 1, dir: 'asc' }], true);
    // เรียงตามราคา: 5 (แถว 5 เดิม) → แถว 2, 10 (แถว 3 เดิม ตัวหนา) → แถว 3
    expect(sheet.cells['B3'].raw).toBe('10');
    expect(sheet.cells['A3'].style).toEqual({ bold: true, fill: '#ff0000' });
    expect(sheet.cells['B3'].fmt).toBe('baht');
    expect(sheet.cells['A2'].style).toBeUndefined();
  });

  it('formulas move like copy/paste; formulas outside the range keep pointing at the old position', () => {
    sortRange(sheet, new Evaluator(wb), 0, range, [{ col: 1, dir: 'asc' }], true);
    // แถว "กล้วย 5" เดิมอยู่แถว 5 (=B5*2) ย้ายมาแถว 2
    expect(sheet.cells['C2'].raw).toBe(shiftFormula('=B5*2', 2 - 5, 0));
    expect(sheet.cells['C2'].raw).toBe('=B2*2');
    expect(new Evaluator(wb).getCellValue('C2')).toBe(10);
    expect(sheet.cells['E1'].raw).toBe('=B2');
  });

  it('is stable for equal keys', () => {
    sortRange(sheet, new Evaluator(wb), 0, range, [{ col: 0, dir: 'asc' }], true);
    // กล้วย 10 (แถว 3 เดิม) มาก่อน กล้วย 5 (แถว 5 เดิม)
    expect(column(sheet, 'A', 2, 3)).toEqual(['กล้วย', 'กล้วย']);
    expect(column(sheet, 'B', 2, 3)).toEqual(['10', '5']);
  });

  it('sorts by multiple columns', () => {
    sortRange(
      sheet,
      new Evaluator(wb),
      0,
      range,
      [
        { col: 0, dir: 'asc' },
        { col: 1, dir: 'asc' },
      ],
      true
    );
    expect(column(sheet, 'A', 2, 6)).toEqual(['กล้วย', 'กล้วย', 'ขนม', 'แตงโม', 'มะม่วง']);
    expect(column(sheet, 'B', 2, 3)).toEqual(['5', '10']);
  });

  it('without a header the first row is sorted too', () => {
    const noHeader = { r1: 1, c1: 0, r2: 5, c2: 1 };
    sortRange(sheet, new Evaluator(wb), 0, noHeader, [{ col: 1, dir: 'desc' }], false);
    expect(column(sheet, 'B', 2, 6)).toEqual(['30', '20', '10', '5', '']);
    expect(sheet.cells['A1'].raw).toBe('สินค้า');
  });

  it('detects formulas that reference other rows in the range (running balance)', () => {
    expect(formulasReferenceOtherRows(sheet, range, true)).toBe(false);
    sheet.cells['C4'] = { raw: '=C3+B4' };
    expect(formulasReferenceOtherRows(sheet, range, true)).toBe(true);
  });

  it('detects merged cells in the range', () => {
    expect(rangeHasMerge(sheet, range)).toBe(false);
    sheet.merges.push({ r1: 2, c1: 0, r2: 2, c2: 1 });
    expect(rangeHasMerge(sheet, range)).toBe(true);
  });
});

describe('เรียงลำดับใน store', () => {
  const S = () => useSheetStore.getState();

  function load(rows: string[][]) {
    const sheet = createEmptySheet('ชีต1');
    rows.forEach((row, r) =>
      row.forEach((raw, c) => {
        if (raw !== '') sheet.cells[`${'ABCDEF'[c]}${r + 1}`] = { raw };
      })
    );
    S().newWorkbook();
    S().loadTemplate({ sheets: [sheet], activeSheet: 0 }, 'ทดสอบ');
  }

  it('quick sort from a single cell guesses the region and header, then undo restores in one step', () => {
    load([
      ['ชื่อ', 'คะแนน'],
      ['ก', '3'],
      ['ข', '1'],
      ['ค', '2'],
    ]);
    S().setActiveCell(2, 1);
    S().requestSort('asc');
    expect(S().pendingSort).toMatchObject({ hasHeader: true, range: { r1: 0, c1: 0, r2: 3, c2: 1 } });
    S().confirmSort();
    const raws = () => [2, 3, 4].map((r) => S().workbook.sheets[0].cells[`A${r}`].raw);
    expect(raws()).toEqual(['ข', 'ค', 'ก']);
    S().undo();
    expect(raws()).toEqual(['ก', 'ข', 'ค']);
    S().redo();
    expect(raws()).toEqual(['ข', 'ค', 'ก']);
  });

  it('multi-cell selection sorts only that range', () => {
    load([['3'], ['1'], ['2'], ['0']]);
    S().setActiveCell(0, 0);
    S().setSelection(0, 0, 2, 0);
    S().requestSort('asc');
    S().updatePendingSort({ hasHeader: false });
    S().confirmSort();
    expect(['A1', 'A2', 'A3', 'A4'].map((k) => S().workbook.sheets[0].cells[k].raw)).toEqual(['1', '2', '3', '0']);
  });

  it('warns about formulas referencing other rows before sorting', () => {
    load([
      ['รายการ', 'จำนวน', 'คงเหลือ'],
      ['ก', '5', '=B2'],
      ['ข', '3', '=C2+B3'],
    ]);
    S().setActiveCell(1, 1);
    S().requestSort('desc');
    S().confirmSort();
    expect(S().pendingSort?.step).toBe('formulaWarning');
    S().cancelSort();
    expect(S().workbook.sheets[0].cells['A2'].raw).toBe('ก');
  });

  it('refuses ranges with merged cells', () => {
    load([
      ['ก', '1'],
      ['ข', '2'],
    ]);
    S().setActiveCell(0, 0);
    S().setSelection(0, 0, 0, 1);
    S().mergeCells();
    S().setActiveCell(1, 0);
    S().requestSort('asc');
    expect(S().pendingSort).toBeNull();
    expect(S().notice).toBe(th.sort.mergedError);
  });
});
