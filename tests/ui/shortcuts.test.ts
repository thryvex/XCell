import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { todaySerial } from '../../src/core/numfmt';

describe('Keyboard Shortcuts & Date insertion (SPEC 5.12)', () => {
  beforeEach(() => {
    useSheetStore.getState().newWorkbook();
  });

  it('inserts current date into single cell with date format and undo support', () => {
    const store = useSheetStore.getState();
    store.setActiveCell(1, 2, false); // C2

    store.insertCurrentDate();

    const state = useSheetStore.getState();
    const val = state.getRawValue(1, 2);
    // 5A: Ctrl+; เก็บเป็นวันที่จริง (serial ค.ศ.) และแสดงเป็น dd/mm/yyyy
    expect(val).toBe(String(todaySerial()));
    expect(state.getDisplayValue(1, 2)).toMatch(/^\d{2}\/\d{2}\/\d{4}$/);

    const sheet = state.workbook.sheets[0];
    expect(sheet.cells['C2']?.fmt).toBe('date');
    expect(state.isDirty).toBe(true);

    // Can be undone
    store.undo();
    expect(useSheetStore.getState().getRawValue(1, 2)).toBe('');
  });

  it('inserts current date across a multi-cell selection', () => {
    const store = useSheetStore.getState();
    store.setSelection(1, 0, 3, 1); // A2:B4

    store.insertCurrentDate();

    const sheet = useSheetStore.getState().workbook.sheets[0];
    expect(sheet.cells['A2']?.fmt).toBe('date');
    expect(sheet.cells['A3']?.fmt).toBe('date');
    expect(sheet.cells['B4']?.fmt).toBe('date');
    expect(sheet.cells['A2']?.raw).toBe(String(todaySerial())); // 5A: วันที่จริง (serial)
  });

  it('selectAll selects entire grid from 0,0 to rowCount-1, colCount-1', () => {
    const store = useSheetStore.getState();
    store.selectAll();

    const sel = useSheetStore.getState().selection;
    expect(sel.r1).toBe(0);
    expect(sel.c1).toBe(0);
    expect(sel.r2).toBe(store.rowCount - 1);
    expect(sel.c2).toBe(store.colCount - 1);
  });
});
