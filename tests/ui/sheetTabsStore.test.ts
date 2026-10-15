import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { th } from '../../src/i18n/th';

const S = () => useSheetStore.getState();
const names = () => S().workbook.sheets.map((s) => s.name);

function typeInto(r: number, c: number, value: string) {
  S().setActiveCell(r, c);
  S().startEditing(value);
  S().commitEditing(0, 0);
}

describe('Phase 4B ขั้น 1: แท็บชีตใน store', () => {
  beforeEach(() => {
    S().newWorkbook();
  });

  it('starts with a single sheet named ชีต1', () => {
    expect(names()).toEqual(['ชีต1']);
    expect(S().activeSheetIndex).toBe(0);
  });

  it('adds sheets with unique default names and switches to the new sheet', () => {
    S().addSheet();
    S().addSheet();
    expect(names()).toEqual(['ชีต1', 'ชีต2', 'ชีต3']);
    expect(S().activeSheetIndex).toBe(2);
    expect(S().isDirty).toBe(true);
  });

  it('keeps cell data separate per sheet', () => {
    typeInto(0, 0, '10');
    S().addSheet();
    expect(S().getRawValue(0, 0)).toBe('');
    typeInto(0, 0, '99');
    S().switchSheet(0);
    expect(S().getRawValue(0, 0)).toBe('10');
    expect(S().getComputedValue(0, 0)).toBe(10);
  });

  it('restores selection, active cell, row count and scroll per sheet', () => {
    S().setActiveCell(5, 3);
    S().setSelection(5, 3, 7, 4);
    S().rememberScroll(400, 120);
    S().addRowsAtBottom(50);

    S().addSheet();
    expect(S().activeCell).toEqual({ r: 0, c: 0 });
    expect(S().rowCount).toBe(100);
    expect(S().getActiveSheetView().scrollTop).toBe(0);
    S().setActiveCell(2, 2);

    S().switchSheet(0);
    expect(S().activeCell).toEqual({ r: 5, c: 3 });
    expect(S().selection).toEqual({ r1: 5, c1: 3, r2: 7, c2: 4 });
    expect(S().rowCount).toBe(150);
    expect(S().getActiveSheetView().scrollTop).toBe(400);
    expect(S().getActiveSheetView().scrollLeft).toBe(120);

    S().switchSheet(1);
    expect(S().activeCell).toEqual({ r: 2, c: 2 });
  });

  it('keeps column width and freeze settings per sheet', () => {
    S().setColWidth(0, 200);
    S().freezeTo(1, 0);
    S().addSheet();
    const sheet2 = S().workbook.sheets[1];
    expect(sheet2.colWidths[0]).toBeUndefined();
    expect(sheet2.frozenRows).toBe(0);
    S().switchSheet(0);
    expect(S().workbook.sheets[0].colWidths[0]).toBe(200);
    expect(S().workbook.sheets[0].frozenRows).toBe(1);
  });

  it('nextSheet / prevSheet (Ctrl+PageDown / Ctrl+PageUp) move between sheets without wrapping', () => {
    S().addSheet();
    S().addSheet();
    S().switchSheet(0);
    S().nextSheet();
    expect(S().activeSheetIndex).toBe(1);
    S().nextSheet();
    S().nextSheet();
    expect(S().activeSheetIndex).toBe(2);
    S().prevSheet();
    expect(S().activeSheetIndex).toBe(1);
  });

  it('status bar values follow the active sheet', () => {
    typeInto(0, 0, '5');
    S().addSheet();
    typeInto(0, 0, '7');
    expect(S().getComputedValue(0, 0)).toBe(7);
    S().switchSheet(0);
    expect(S().getComputedValue(0, 0)).toBe(5);
  });

  it('renames with validation errors in Thai', () => {
    S().addSheet();
    expect(S().renameSheet(1, '')).toEqual({ success: false, error: th.sheetTabs.nameEmpty });
    expect(S().renameSheet(1, 'ชีต1').success).toBe(false);
    expect(S().renameSheet(1, 'a/b')).toEqual({ success: false, error: th.sheetTabs.nameInvalidChars });
    expect(S().renameSheet(1, '  สรุป  ')).toEqual({ success: true, error: undefined });
    expect(names()).toEqual(['ชีต1', 'สรุป']);
  });

  it('duplicates next to the source with copied cells', () => {
    typeInto(0, 0, '42');
    S().addSheet();
    S().duplicateSheet(0);
    expect(names()).toEqual(['ชีต1', 'สำเนาของ ชีต1', 'ชีต2']);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().getRawValue(0, 0)).toBe('42');
    expect(S().workbook.sheets[1].id).not.toBe(S().workbook.sheets[0].id);
  });

  it('moves sheets and keeps showing the same sheet', () => {
    S().addSheet(); // ชีต2 active
    S().addSheet(); // ชีต3 active
    S().moveSheet(2, 0);
    expect(names()).toEqual(['ชีต3', 'ชีต1', 'ชีต2']);
    expect(S().workbook.sheets[S().activeSheetIndex].name).toBe('ชีต3');
    S().moveSheet(1, 2);
    expect(names()).toEqual(['ชีต3', 'ชีต2', 'ชีต1']);
    expect(S().workbook.sheets[S().activeSheetIndex].name).toBe('ชีต3');
  });

  it('refuses to delete the last sheet', () => {
    expect(S().deleteSheet(0)).toEqual({ success: false, error: th.sheetTabs.cannotDeleteLastSheet });
    expect(S().pendingDeleteSheet).toBeNull();
  });

  it('always asks before deleting, and cancel keeps the sheet', () => {
    S().addSheet();
    expect(S().deleteSheet(1).success).toBe(true);
    expect(S().pendingDeleteSheet).toEqual({ sheetIndex: 1, sheetName: 'ชีต2', referencingFormulaCount: 0 });
    S().cancelDeleteSheet();
    expect(names()).toEqual(['ชีต1', 'ชีต2']);
  });

  it('deletes after confirmation and shows a neighbouring sheet', () => {
    S().addSheet();
    S().addSheet();
    S().switchSheet(1);
    S().deleteSheet(1);
    S().confirmDeleteSheet();
    expect(names()).toEqual(['ชีต1', 'ชีต3']);
    expect(S().workbook.sheets[S().activeSheetIndex].name).toBe('ชีต3');
  });

  it('deleting a non-active sheet keeps the active sheet', () => {
    S().addSheet();
    S().addSheet(); // active = ชีต3
    S().deleteSheet(0);
    S().confirmDeleteSheet();
    expect(names()).toEqual(['ชีต2', 'ชีต3']);
    expect(S().workbook.sheets[S().activeSheetIndex].name).toBe('ชีต3');
  });
});

describe('Phase 4B ขั้น 1: Undo/Redo ของการจัดการชีต', () => {
  beforeEach(() => {
    S().newWorkbook();
  });

  it('undo/redo add sheet and switches sheets', () => {
    S().addSheet();
    expect(S().activeSheetIndex).toBe(1);
    S().undo();
    expect(names()).toEqual(['ชีต1']);
    expect(S().activeSheetIndex).toBe(0);
    S().redo();
    expect(names()).toEqual(['ชีต1', 'ชีต2']);
    expect(S().activeSheetIndex).toBe(1);
  });

  it('undo delete brings back the sheet with its data and shows it', () => {
    S().addSheet();
    typeInto(0, 0, 'ข้อมูล');
    S().switchSheet(0);
    S().deleteSheet(1);
    S().confirmDeleteSheet();
    expect(names()).toEqual(['ชีต1']);

    S().undo();
    expect(names()).toEqual(['ชีต1', 'ชีต2']);
    expect(S().activeSheetIndex).toBe(1);
    expect(S().getRawValue(0, 0)).toBe('ข้อมูล');

    S().redo();
    expect(names()).toEqual(['ชีต1']);
  });

  it('undo rename, duplicate and move', () => {
    S().addSheet();
    S().renameSheet(1, 'สรุป');
    S().duplicateSheet(1);
    S().moveSheet(2, 0);
    expect(names()).toEqual(['สำเนาของ สรุป', 'ชีต1', 'สรุป']);

    S().undo();
    expect(names()).toEqual(['ชีต1', 'สรุป', 'สำเนาของ สรุป']);
    expect(S().activeSheetIndex).toBe(2);
    S().undo();
    expect(names()).toEqual(['ชีต1', 'สรุป']);
    expect(S().activeSheetIndex).toBe(1);
    S().undo();
    expect(names()).toEqual(['ชีต1', 'ชีต2']);
    expect(S().activeSheetIndex).toBe(1);
  });

  it('undo of a cell edit switches back to the sheet that was edited', () => {
    typeInto(0, 0, '1');
    S().addSheet();
    expect(S().activeSheetIndex).toBe(1);
    S().switchSheet(0);
    S().switchSheet(1);
    // ประวัติล่าสุดคือเพิ่มชีต → undo ลบชีต2 แล้วกลับชีต1
    S().undo();
    S().undo();
    expect(S().activeSheetIndex).toBe(0);
    expect(S().getRawValue(0, 0)).toBe('');
  });

  it('cell edits on different sheets undo in order across sheets', () => {
    typeInto(0, 0, 'a');
    S().addSheet();
    typeInto(0, 0, 'b');
    S().switchSheet(0);
    S().undo(); // undo 'b' on ชีต2
    expect(S().activeSheetIndex).toBe(1);
    expect(S().getRawValue(0, 0)).toBe('');
    S().redo();
    expect(S().getRawValue(0, 0)).toBe('b');
  });
});

describe('Phase 4B: คัดลอก/ตัด ข้ามชีต', () => {
  beforeEach(() => {
    S().newWorkbook();
  });

  it('copies a range from one sheet and pastes into another', async () => {
    typeInto(0, 0, '10');
    typeInto(1, 0, '=A1*2');
    S().setActiveCell(0, 0);
    S().setSelection(0, 0, 1, 0);
    S().copySelection();

    S().addSheet();
    S().setActiveCell(4, 1);
    await S().paste(false);
    expect(S().getRawValue(4, 1)).toBe('10');
    expect(S().getRawValue(5, 1)).toBe('=B5*2');
    expect(S().getComputedValue(5, 1)).toBe(20);
  });

  it('paste values only uses values computed on the source sheet', async () => {
    typeInto(0, 0, '3');
    typeInto(0, 1, '=A1*5');
    S().setActiveCell(0, 1);
    S().copySelection();

    S().addSheet();
    S().setActiveCell(0, 0);
    await S().paste(true);
    expect(S().getRawValue(0, 0)).toBe('15');
  });

  it('cut from one sheet clears the source sheet, and undo restores both', async () => {
    typeInto(0, 0, 'ย้าย');
    S().setActiveCell(0, 0);
    S().cutSelection();

    S().addSheet();
    S().setActiveCell(2, 2);
    await S().paste(false);
    expect(S().getRawValue(2, 2)).toBe('ย้าย');
    expect(S().workbook.sheets[0].cells['A1']).toBeUndefined();

    S().undo();
    expect(S().workbook.sheets[1].cells['C3']).toBeUndefined();
    expect(S().workbook.sheets[0].cells['A1'].raw).toBe('ย้าย');
  });
});
