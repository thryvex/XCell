import { describe, it, expect } from 'vitest';
import { HistoryManager } from '../../src/core/history';
import { createEmptyWorkbook } from '../../src/core/model';

describe('src/core/history', () => {
  it('records actions and undos cell changes', () => {
    const wb = createEmptyWorkbook();
    const history = new HistoryManager();

    // User types in A1: "100"
    wb.sheets[0].cells['A1'] = { raw: '100' };
    history.push({
      description: 'Edit A1',
      sheetIndex: 0,
      patch: {
        cellPatches: [{ key: 'A1', before: undefined, after: { raw: '100' } }],
      },
    });

    expect(wb.sheets[0].cells['A1'].raw).toBe('100');
    expect(history.canUndo()).toBe(true);
    expect(history.canRedo()).toBe(false);

    // Undo
    history.undo(wb);
    expect(wb.sheets[0].cells['A1']).toBeUndefined();
    expect(history.canUndo()).toBe(false);
    expect(history.canRedo()).toBe(true);

    // Redo
    history.redo(wb);
    expect(wb.sheets[0].cells['A1'].raw).toBe('100');
    expect(history.canUndo()).toBe(true);
  });

  it('records merge undo and redo', () => {
    const wb = createEmptyWorkbook();
    const history = new HistoryManager();

    const newMerge = { r1: 0, c1: 0, r2: 1, c2: 1 };
    wb.sheets[0].merges = [newMerge];
    history.push({
      description: 'Merge A1:B2',
      sheetIndex: 0,
      patch: {
        mergesBefore: [],
        mergesAfter: [newMerge],
      },
    });

    expect(wb.sheets[0].merges.length).toBe(1);

    history.undo(wb);
    expect(wb.sheets[0].merges.length).toBe(0);

    history.redo(wb);
    expect(wb.sheets[0].merges.length).toBe(1);
  });

  it('records full sheet snapshot for row/col insert/delete undo and redo', () => {
    const wb = createEmptyWorkbook();
    const history = new HistoryManager();
    const sheet = wb.sheets[0];

    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['A2'] = { raw: '=A1*2' };

    const snapshotBefore = {
      cells: JSON.parse(JSON.stringify(sheet.cells)),
      merges: [],
      rowCount: 100,
      colCount: 26,
    };

    // Simulate inserting a row at index 0
    sheet.cells = {
      A2: { raw: '10' },
      A3: { raw: '=A2*2' },
    };

    const snapshotAfter = {
      cells: JSON.parse(JSON.stringify(sheet.cells)),
      merges: [],
      rowCount: 101,
      colCount: 26,
    };

    history.push({
      description: 'Insert row 1',
      sheetIndex: 0,
      patch: { fullSheetBefore: snapshotBefore, fullSheetAfter: snapshotAfter },
    });

    expect(wb.sheets[0].cells['A3'].raw).toBe('=A2*2');

    // Undo row insertion
    const undoAction = history.undo(wb);
    expect(undoAction?.patch.fullSheetBefore?.rowCount).toBe(100);
    expect(wb.sheets[0].cells['A1'].raw).toBe('10');
    expect(wb.sheets[0].cells['A2'].raw).toBe('=A1*2');
    expect(wb.sheets[0].cells['A3']).toBeUndefined();

    // Redo row insertion
    history.redo(wb);
    expect(wb.sheets[0].cells['A2'].raw).toBe('10');
    expect(wb.sheets[0].cells['A3'].raw).toBe('=A2*2');
  });

  it('records column insert and delete with undo and redo', () => {
    const wb = createEmptyWorkbook();
    const history = new HistoryManager();
    const sheet = wb.sheets[0];

    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['B1'] = { raw: '=A1*2' };

    const snapshotBefore = {
      cells: JSON.parse(JSON.stringify(sheet.cells)),
      merges: [],
      rowCount: 100,
      colCount: 26,
    };

    // Insert column at col 1 (B) -> old B becomes C
    sheet.cells = {
      A1: { raw: '10' },
      C1: { raw: '=A1*2' },
    };

    const snapshotAfter = {
      cells: JSON.parse(JSON.stringify(sheet.cells)),
      merges: [],
      rowCount: 100,
      colCount: 27,
    };

    history.push({
      description: 'Insert column 1',
      sheetIndex: 0,
      patch: { fullSheetBefore: snapshotBefore, fullSheetAfter: snapshotAfter },
    });

    expect(wb.sheets[0].cells['C1'].raw).toBe('=A1*2');

    // Undo
    history.undo(wb);
    expect(wb.sheets[0].cells['B1'].raw).toBe('=A1*2');
    expect(wb.sheets[0].cells['C1']).toBeUndefined();

    // Redo
    history.redo(wb);
    expect(wb.sheets[0].cells['B1']).toBeUndefined();
    expect(wb.sheets[0].cells['C1'].raw).toBe('=A1*2');
  });

  it('records clear format and clear content undo and redo', () => {
    const wb = createEmptyWorkbook();
    const history = new HistoryManager();
    const sheet = wb.sheets[0];

    sheet.cells['A1'] = { raw: '500', fmt: 'baht', style: { bold: true } };

    // Clear format
    const formatBefore = { ...sheet.cells['A1'] };
    const formatAfter = { raw: '500' };
    sheet.cells['A1'] = formatAfter;

    history.push({
      description: 'Clear formatting',
      sheetIndex: 0,
      patch: {
        cellPatches: [{ key: 'A1', before: formatBefore, after: formatAfter }],
      },
    });

    expect(wb.sheets[0].cells['A1'].fmt).toBeUndefined();
    expect(wb.sheets[0].cells['A1'].style).toBeUndefined();
    expect(wb.sheets[0].cells['A1'].raw).toBe('500');

    // Undo clear format
    history.undo(wb);
    expect(wb.sheets[0].cells['A1'].fmt).toBe('baht');
    expect(wb.sheets[0].cells['A1'].style?.bold).toBe(true);

    // Redo clear format
    history.redo(wb);
    expect(wb.sheets[0].cells['A1'].fmt).toBeUndefined();
  });
});

