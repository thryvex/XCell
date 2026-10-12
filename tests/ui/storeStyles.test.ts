import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';

describe('useSheetStore styling actions (Phase 3A)', () => {
  beforeEach(() => {
    useSheetStore.getState().newWorkbook();
  });

  it('toggles bold, italic, underline, and wrap on active cell', () => {
    const store = useSheetStore.getState();
    store.setActiveCell(0, 0); // A1

    expect(store.getActiveCellStyle().bold).toBeUndefined();

    // Toggle bold on
    store.toggleBold();
    expect(useSheetStore.getState().getActiveCellStyle().bold).toBe(true);

    // Toggle italic on
    store.toggleItalic();
    expect(useSheetStore.getState().getActiveCellStyle().italic).toBe(true);

    // Toggle underline on
    store.toggleUnderline();
    expect(useSheetStore.getState().getActiveCellStyle().underline).toBe(true);

    // Toggle wrap on
    store.toggleWrap();
    expect(useSheetStore.getState().getActiveCellStyle().wrap).toBe(true);

    // Toggle bold off
    store.toggleBold();
    expect(useSheetStore.getState().getActiveCellStyle().bold).toBeUndefined();
    // other styles remain
    expect(useSheetStore.getState().getActiveCellStyle().italic).toBe(true);
  });

  it('sets text color, fill color, alignment, and border', () => {
    const store = useSheetStore.getState();
    store.setActiveCell(1, 1); // B2

    store.setTextColor('#ef4444');
    expect(useSheetStore.getState().getActiveCellStyle().color).toBe('#ef4444');

    store.setFillColor('#bbf7d0');
    expect(useSheetStore.getState().getActiveCellStyle().fill).toBe('#bbf7d0');

    store.setAlignment('center');
    expect(useSheetStore.getState().getActiveCellStyle().align).toBe('center');

    store.setBorder('all');
    expect(useSheetStore.getState().getActiveCellStyle().border).toBe('all');
  });

  it('applies styles across a multi-cell selection and supports undo/redo', () => {
    const store = useSheetStore.getState();
    store.setSelection(0, 0, 1, 1); // A1:B2

    store.setCellStyle({ bold: true, color: '#3b82f6' });

    const sheet = useSheetStore.getState().workbook.sheets[0];
    expect(sheet.cells['A1']?.style?.bold).toBe(true);
    expect(sheet.cells['A1']?.style?.color).toBe('#3b82f6');
    expect(sheet.cells['B2']?.style?.bold).toBe(true);
    expect(sheet.cells['B2']?.style?.color).toBe('#3b82f6');

    // Undo
    useSheetStore.getState().undo();
    const sheetAfterUndo = useSheetStore.getState().workbook.sheets[0];
    expect(sheetAfterUndo.cells['A1']?.style).toBeUndefined();
    expect(sheetAfterUndo.cells['B2']?.style).toBeUndefined();

    // Redo
    useSheetStore.getState().redo();
    const sheetAfterRedo = useSheetStore.getState().workbook.sheets[0];
    expect(sheetAfterRedo.cells['A1']?.style?.bold).toBe(true);
    expect(sheetAfterRedo.cells['B2']?.style?.color).toBe('#3b82f6');
  });

  it('preserves target cell styling when paste(true) (valuesOnly) is used', async () => {
    const store = useSheetStore.getState();

    // Setup source cell A1: raw: 'Hello', style: bold
    store.setActiveCell(0, 0);
    store.startEditing('Hello');
    store.commitEditing();
    store.toggleBold();
    store.copySelection();

    // Setup target cell B1: raw: 'Old', style: italic, text color: red
    store.setActiveCell(0, 1);
    store.startEditing('Old');
    store.commitEditing();
    store.toggleItalic();
    store.setTextColor('#ef4444');

    // Paste values only into B1
    await store.paste(true);

    const sheet = useSheetStore.getState().workbook.sheets[0];
    const b1 = sheet.cells['B1'];
    expect(b1.raw).toBe('Hello');
    expect(b1.style?.italic).toBe(true);
    expect(b1.style?.color).toBe('#ef4444');
    expect(b1.style?.bold).toBeUndefined(); // Did not overwrite with source bold!
  });
});
