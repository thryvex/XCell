import { describe, it, expect, beforeEach } from 'vitest';
import {
  conditionMatches,
  computeHiddenRows,
  uniqueColumnValues,
  createFilter,
  adjustFilterOnInsertRows,
  adjustFilterOnDeleteRows,
  adjustFilterOnInsertCols,
  adjustFilterOnDeleteCols,
} from '../../src/core/filter';
import { insertRow, deleteRow } from '../../src/core/sheetOps';
import { Evaluator } from '../../src/core/evaluator';
import { Sheet, createEmptySheet, Workbook } from '../../src/core/model';
import { exportToXlsxBuffer, importFromXlsxBuffer, workbookToExcelJS } from '../../src/core/xlsx';
import { serializeRecoveryData, deserializeRecoveryData } from '../../src/core/recovery';
import { useSheetStore } from '../../src/ui/store';

describe('conditionMatches (กรองตามเงื่อนไข)', () => {
  it('empty / notEmpty', () => {
    expect(conditionMatches({ type: 'empty' }, '', '')).toBe(true);
    expect(conditionMatches({ type: 'empty' }, 0, '0')).toBe(false);
    expect(conditionMatches({ type: 'notEmpty' }, 'ก', 'ก')).toBe(true);
    expect(conditionMatches({ type: 'notEmpty' }, '', '')).toBe(false);
  });

  it('text contains / not contains / starts with (case-insensitive, Thai)', () => {
    expect(conditionMatches({ type: 'textContains', value: 'อาหาร' }, 'ค่าอาหารเช้า', 'ค่าอาหารเช้า')).toBe(true);
    expect(conditionMatches({ type: 'textContains', value: 'APP' }, 'apple', 'apple')).toBe(true);
    expect(conditionMatches({ type: 'textNotContains', value: 'อาหาร' }, 'เดินทาง', 'เดินทาง')).toBe(true);
    expect(conditionMatches({ type: 'textNotContains', value: 'อาหาร' }, 'อาหาร', 'อาหาร')).toBe(false);
    expect(conditionMatches({ type: 'startsWith', value: 'ค่า' }, 'ค่าน้ำ', 'ค่าน้ำ')).toBe(true);
    expect(conditionMatches({ type: 'startsWith', value: 'น้ำ' }, 'ค่าน้ำ', 'ค่าน้ำ')).toBe(false);
  });

  it('equals compares numbers by value and text by display', () => {
    expect(conditionMatches({ type: 'equals', value: '1,000' }, 1000, '฿1,000.00')).toBe(true);
    expect(conditionMatches({ type: 'equals', value: 'อาหาร' }, 'อาหาร', 'อาหาร')).toBe(true);
    expect(conditionMatches({ type: 'equals', value: 'อาหาร' }, 'อาหารเช้า', 'อาหารเช้า')).toBe(false);
  });

  it('greater / less than work on numbers and dates only', () => {
    expect(conditionMatches({ type: 'greaterThan', value: '100' }, 150, '150')).toBe(true);
    expect(conditionMatches({ type: 'greaterThan', value: '100' }, 100, '100')).toBe(false);
    expect(conditionMatches({ type: 'greaterThan', value: '100' }, 'ข้อความ', 'ข้อความ')).toBe(false);
    expect(conditionMatches({ type: 'lessThan', value: '100' }, '', '')).toBe(false);
    // วันที่: 1/10/2026 = serial 46296
    expect(conditionMatches({ type: 'lessThan', value: '1/10/2026' }, 46295, '30/09/2026')).toBe(true);
  });

  it('between is inclusive and accepts min/max in either order', () => {
    const between = { type: 'between' as const, min: '10', max: '20' };
    expect(conditionMatches(between, 10, '10')).toBe(true);
    expect(conditionMatches(between, 20, '20')).toBe(true);
    expect(conditionMatches(between, 21, '21')).toBe(false);
    expect(conditionMatches({ type: 'between', min: '20', max: '10' }, 15, '15')).toBe(true);
    expect(conditionMatches(between, 'สิบห้า', 'สิบห้า')).toBe(false);
  });
});

function dataSheet(): Sheet {
  const sheet = createEmptySheet('ชีต1');
  const rows = [
    ['หมวด', 'จำนวน'],
    ['อาหาร', '100'],
    ['เดินทาง', '50'],
    ['อาหาร', '30'],
    ['', '70'],
    ['ของใช้', '200'],
  ];
  rows.forEach((row, r) =>
    row.forEach((raw, c) => {
      if (raw !== '') sheet.cells[`${'AB'[c]}${r + 1}`] = { raw };
    })
  );
  sheet.filter = createFilter({ r1: 0, c1: 0, r2: 5, c2: 1 });
  return sheet;
}

describe('computeHiddenRows / uniqueColumnValues', () => {
  let sheet: Sheet;
  let ev: Evaluator;
  beforeEach(() => {
    sheet = dataSheet();
    ev = new Evaluator(sheet);
  });

  it('filters by values (blank included as a value)', () => {
    sheet.filter!.columns[0] = { kind: 'values', hiddenValues: ['อาหาร', ''] };
    expect(computeHiddenRows(sheet, ev, 0)).toEqual([1, 3, 4]);
  });

  it('combines several columns with AND', () => {
    sheet.filter!.columns[0] = { kind: 'values', hiddenValues: ['เดินทาง'] };
    sheet.filter!.columns[1] = { kind: 'condition', condition: { type: 'greaterThan', value: '60' } };
    // แสดงเฉพาะแถวที่หมวดไม่ใช่เดินทาง และจำนวน > 60
    expect(computeHiddenRows(sheet, ev, 0)).toEqual([2, 3]);
  });

  it('lists unique values with counts, numbers sorted by value, blank last', () => {
    expect(uniqueColumnValues(sheet, ev, 0, 0)).toEqual([
      { display: 'ของใช้', count: 1 },
      { display: 'เดินทาง', count: 1 },
      { display: 'อาหาร', count: 2 },
      { display: '', count: 1 },
    ]);
    expect(uniqueColumnValues(sheet, ev, 0, 1).map((v) => v.display)).toEqual(['30', '50', '70', '100', '200']);
  });
});

describe('ตัวกรองกับการแทรก/ลบแถว-คอลัมน์', () => {
  it('inserting a row inside the data shifts hidden rows and grows the range', () => {
    const sheet = dataSheet();
    sheet.hiddenRows = [2, 4];
    insertRow(sheet, 3, 1);
    expect(sheet.hiddenRows).toEqual([2, 5]);
    expect(sheet.filter).toMatchObject({ r1: 0, r2: 6 });
    // แถวที่ซ่อนยังเป็นข้อมูลเดิม
    expect(sheet.cells['A3'].raw).toBe('เดินทาง');
    expect(sheet.cells['B6'].raw).toBe('70');
  });

  it('inserting above the header moves the whole filter', () => {
    const sheet = dataSheet();
    sheet.hiddenRows = [2];
    insertRow(sheet, 0, 2);
    expect(sheet.filter).toMatchObject({ r1: 2, r2: 7 });
    expect(sheet.hiddenRows).toEqual([4]);
  });

  it('deleting rows removes/shifts hidden rows and shrinks the range', () => {
    const sheet = dataSheet();
    sheet.hiddenRows = [2, 4];
    deleteRow(sheet, 2, 1); // ลบแถวที่ซ่อนอยู่เอง
    expect(sheet.hiddenRows).toEqual([3]);
    expect(sheet.filter).toMatchObject({ r1: 0, r2: 4 });
    expect(sheet.cells['B4'].raw).toBe('70');
  });

  it('deleting the header row removes the filter and shows all rows', () => {
    const sheet = dataSheet();
    sheet.hiddenRows = [2];
    adjustFilterOnDeleteRows(sheet, 0, 1);
    expect(sheet.filter).toBeUndefined();
    expect(sheet.hiddenRows).toBeUndefined();
  });

  it('inserting/deleting columns shifts criteria keys', () => {
    const sheet = dataSheet();
    sheet.filter!.columns[1] = { kind: 'condition', condition: { type: 'notEmpty' } };
    adjustFilterOnInsertCols(sheet, 1, 1);
    expect(sheet.filter).toMatchObject({ c1: 0, c2: 2 });
    expect(Object.keys(sheet.filter!.columns)).toEqual(['2']);

    expect(adjustFilterOnDeleteCols(sheet, 0, 1)).toBe(false);
    expect(sheet.filter).toMatchObject({ c1: 0, c2: 1 });
    expect(Object.keys(sheet.filter!.columns)).toEqual(['1']);

    expect(adjustFilterOnDeleteCols(sheet, 1, 1)).toBe(true); // คอลัมน์ที่มีเงื่อนไขถูกลบ
    expect(sheet.filter!.columns).toEqual({});
  });

  it('insert row exactly below the range does not extend it', () => {
    const sheet = dataSheet();
    adjustFilterOnInsertRows(sheet, 6, 1);
    expect(sheet.filter).toMatchObject({ r2: 5 });
  });
});

describe('ตัวกรองใน store', () => {
  const S = () => useSheetStore.getState();
  const hidden = () => S().workbook.sheets[0].hiddenRows ?? [];

  beforeEach(() => {
    S().newWorkbook();
    S().loadTemplate({ sheets: [dataSheet()], activeSheet: 0 }, 'ทดสอบ');
    // ลบตัวกรองที่ติดมากับข้อมูลทดสอบ เพื่อเริ่มจากไม่มีตัวกรอง
    delete S().workbook.sheets[0].filter;
    S().setActiveCell(2, 0);
  });

  it('toggle creates a filter on the data region and removes it again; both undoable', () => {
    S().toggleFilter();
    expect(S().workbook.sheets[0].filter).toMatchObject({ r1: 0, c1: 0, r2: 5, c2: 1 });
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['อาหาร'] });
    expect(hidden()).toEqual([1, 3]);

    S().toggleFilter(); // ปิด: แสดงทุกแถว
    expect(S().workbook.sheets[0].filter).toBeUndefined();
    expect(hidden()).toEqual([]);

    S().undo();
    expect(hidden()).toEqual([1, 3]);
    S().undo();
    expect(hidden()).toEqual([]);
    expect(S().workbook.sheets[0].filter).toBeDefined();
    S().redo();
    expect(hidden()).toEqual([1, 3]);
  });

  it('editing a filtered column does not re-filter until "reapply"', () => {
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['อาหาร'] });
    S().setActiveCell(2, 0);
    S().startEditing('อาหาร');
    S().commitEditing();
    expect(hidden()).toEqual([1, 3]);
    S().reapplyFilter();
    expect(hidden()).toEqual([1, 2, 3]);
  });

  it('insert/delete rows keep the hidden rows on the same data, and undo restores them', () => {
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['เดินทาง'] });
    expect(hidden()).toEqual([2]);
    S().insertRowAt(1);
    expect(hidden()).toEqual([3]);
    expect(S().workbook.sheets[0].cells['A4'].raw).toBe('เดินทาง');
    S().undo();
    expect(hidden()).toEqual([2]);
    S().deleteRowAt(1, true);
    expect(hidden()).toEqual([1]);
    expect(S().workbook.sheets[0].cells['A2'].raw).toBe('เดินทาง');
  });

  it('copy skips hidden rows; paste and clear do not touch hidden rows', async () => {
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['เดินทาง'] }); // ซ่อนแถว 3 (index 2)
    S().setActiveCell(1, 1);
    S().setSelection(1, 1, 3, 1); // B2:B4 = 100, (50 ซ่อน), 30
    S().copySelection();
    expect(S().clipboard).toMatchObject({ r1: 1, r2: 2 });

    S().setActiveCell(1, 3);
    await S().paste(false);
    const cells = S().workbook.sheets[0].cells;
    expect(cells['D2'].raw).toBe('100');
    expect(cells['D3']).toBeUndefined(); // แถวที่ซ่อนไม่ถูกวางทับ
    expect(cells['D4'].raw).toBe('30');

    S().setActiveCell(1, 1);
    S().setSelection(1, 1, 3, 1);
    S().clearSelectionContent();
    expect(cells['B2']).toBeUndefined();
    expect(cells['B3'].raw).toBe('50'); // ซ่อนอยู่ ไม่ถูกล้าง
    expect(cells['B4']).toBeUndefined();
  });

  it('fill down does not write into hidden rows', () => {
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['เดินทาง'] });
    S().setActiveCell(1, 2);
    S().startEditing('x');
    S().commitEditing();
    S().setActiveCell(1, 2);
    S().setSelection(1, 2, 3, 2);
    S().fillDownAction();
    const cells = S().workbook.sheets[0].cells;
    expect(cells['C3']).toBeUndefined();
    expect(cells['C4'].raw).toBe('x');
  });

  it('sorting from the filter menu sorts data under the header and re-applies criteria', () => {
    S().toggleFilter();
    S().setColumnFilter(1, { kind: 'condition', condition: { type: 'greaterThan', value: '60' } });
    expect(hidden()).toEqual([2, 3]);
    S().sortFilterColumn(1, 'desc');
    const col = [1, 2, 3, 4, 5, 6].map((r) => S().workbook.sheets[0].cells[`B${r}`]?.raw);
    expect(col).toEqual(['จำนวน', '200', '100', '70', '50', '30']);
    expect(hidden()).toEqual([4, 5]);
    S().undo();
    expect(hidden()).toEqual([2, 3]);
  });

  it('file round-trip keeps the autoFilter range and hidden rows (criteria are not stored)', async () => {
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['อาหาร'] });
    const excel = await workbookToExcelJS(S().workbook);
    const ws = excel.worksheets[0];
    expect(ws.autoFilter).toEqual({ from: { row: 1, column: 1 }, to: { row: 6, column: 2 } });
    expect(ws.getRow(2).hidden).toBe(true);
    expect(ws.getRow(3).hidden).toBe(false);

    const buffer = await exportToXlsxBuffer(S().workbook);
    const reloaded = await importFromXlsxBuffer(buffer);
    expect(reloaded.sheets[0].filter).toEqual({ r1: 0, c1: 0, r2: 5, c2: 1, columns: {} });
    expect(reloaded.sheets[0].hiddenRows).toEqual([1, 3]);

    await S().loadWorkbook(buffer, 'C:\\ไฟล์\\กรอง.xlsx');
    expect(S().filteredRowsNotice).toBe(true);
    expect(hidden()).toEqual([1, 3]);
    // ไม่มีเงื่อนไขที่รู้จัก: ใช้ตัวกรองอีกครั้งจะไม่ไปแตะแถวที่ซ่อน
    S().reapplyFilter();
    expect(hidden()).toEqual([1, 3]);
    S().showAllFilteredRows();
    expect(hidden()).toEqual([]);
    expect(S().filteredRowsNotice).toBe(false);
    S().undo();
    expect(hidden()).toEqual([1, 3]);
  });

  it('sorting keeps hidden rows with their data when criteria are unknown (opened file)', async () => {
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['เดินทาง'] }); // ซ่อน "เดินทาง 50"
    await S().loadWorkbook(await exportToXlsxBuffer(S().workbook), 'x.xlsx');
    S().sortFilterColumn(1, 'asc');
    const sheet = S().workbook.sheets[0];
    const hiddenValues = (sheet.hiddenRows ?? []).map((r) => sheet.cells[`B${r + 1}`]?.raw);
    expect(hiddenValues).toEqual(['50']);
  });

  it('recovery data keeps criteria and hidden rows', () => {
    S().toggleFilter();
    S().setColumnFilter(1, { kind: 'condition', condition: { type: 'between', min: '40', max: '120' } });
    const wb: Workbook = S().workbook;
    const restored = deserializeRecoveryData(
      serializeRecoveryData({ id: 'r', originalPath: null, fileName: 'x', savedAt: 0, workbook: wb })
    )!;
    expect(restored.workbook.sheets[0].filter).toEqual(wb.sheets[0].filter);
    expect(restored.workbook.sheets[0].hiddenRows).toEqual([3, 5]);

    S().restoreRecoveryWorkbook(restored.workbook, null);
    expect(S().workbook.sheets[0].filter?.columns[1]).toEqual({
      kind: 'condition',
      condition: { type: 'between', min: '40', max: '120' },
    });
    expect(hidden()).toEqual([3, 5]);
  });
});
