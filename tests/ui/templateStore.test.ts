import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { createLedgerWorkbook, createInventoryWorkbook } from '../../src/templates';

describe('useSheetStore Template Loading (Phase 3D)', () => {
  beforeEach(() => {
    useSheetStore.getState().newWorkbook();
  });

  it('loads Ledger template into store with clean state and row 2 active cell', () => {
    const store = useSheetStore.getState();
    const ledgerWb = createLedgerWorkbook();

    store.loadTemplate(ledgerWb, 'ตารางรายรับรายจ่าย');

    const state = useSheetStore.getState();
    expect(state.fileName).toBe('ตารางรายรับรายจ่าย');
    expect(state.filePath).toBeNull();
    expect(state.isDirty).toBe(false);
    expect(state.activeSheetIndex).toBe(0);
    expect(state.workbook.sheets[0].name).toBe('บันทึก');
    expect(state.activeCell).toEqual({ r: 1, c: 0 }); // Focus row 2 (first data row)
    expect(state.getRawValue(0, 0)).toBe('วันที่');
    expect(state.getRawValue(0, 5)).toBe('คงเหลือ');
  });

  it('loads Inventory template into store with clean state and row 2 active cell', () => {
    const store = useSheetStore.getState();
    const invWb = createInventoryWorkbook();

    store.loadTemplate(invWb, 'ตารางรายการสินค้า');

    const state = useSheetStore.getState();
    expect(state.fileName).toBe('ตารางรายการสินค้า');
    expect(state.filePath).toBeNull();
    expect(state.isDirty).toBe(false);
    expect(state.workbook.sheets[0].name).toBe('รายการสินค้า');
    expect(state.activeCell).toEqual({ r: 1, c: 0 });
    expect(state.getRawValue(0, 0)).toBe('ลำดับ');
    expect(state.getRawValue(0, 4)).toBe('รวม');
  });
});
