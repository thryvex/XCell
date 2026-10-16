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
    expect(state.getRawValue(0, 4)).toBe('น้ำหนักสุทธิ');
    expect(state.getRawValue(0, 6)).toBe('รวม');
  });

  it('Inventory template: typing a new row 5 adds to the J1 total', () => {
    const store = useSheetStore.getState();
    store.loadTemplate(createInventoryWorkbook(), 'ตารางรายการสินค้า');
    const type = (r: number, c: number, text: string) => {
      useSheetStore.getState().setActiveCell(r, c);
      useSheetStore.getState().startEditing(text);
      useSheetStore.getState().commitEditing();
    };
    const total = () => useSheetStore.getState().getComputedValue(0, 9) as number; // J1
    expect(total()).toBeCloseTo(5784.8);

    // แถว 5: น้ำหนัก 50, เปอร์เซ็น 10, ราคา 20 พร้อมสูตร E/G แบบเดียวกับแถวตัวอย่าง
    type(4, 2, '50');
    type(4, 3, '10');
    type(4, 4, '=C5*D5%');
    type(4, 5, '20');
    type(4, 6, '=E5*F5');
    expect(useSheetStore.getState().getComputedValue(4, 6)).toBeCloseTo(100);
    expect(total()).toBeCloseTo(5884.8);
  });
});
