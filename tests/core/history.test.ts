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
});
