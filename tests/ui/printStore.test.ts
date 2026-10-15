import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { serializeRecoveryData, deserializeRecoveryData } from '../../src/core/recovery';
import { DEFAULT_PRINT_SETTINGS } from '../../src/core/model';

const S = () => useSheetStore.getState();

describe('ตัวเลือกพิมพ์ใน store', () => {
  beforeEach(() => S().newWorkbook());

  it('remembers the last settings per file without marking it unsaved or adding undo steps', () => {
    expect(S().workbook.printSettings).toBeUndefined();
    S().setPrintSettings({ orientation: 'landscape', paper: 'A5' });
    S().setPrintSettings({ gridlines: true });
    expect(S().workbook.printSettings).toEqual({
      ...DEFAULT_PRINT_SETTINGS,
      orientation: 'landscape',
      paper: 'A5',
      gridlines: true,
    });
    expect(S().isDirty).toBe(false);
    expect(S().history.canUndo()).toBe(false);
  });

  it('is stored in recovery data and restored', () => {
    S().setPrintSettings({ scaleMode: 'custom', scalePercent: 80, header: true });
    const data = deserializeRecoveryData(
      serializeRecoveryData({ id: 'r', originalPath: null, fileName: 'x', savedAt: 0, workbook: S().workbook })
    )!;
    S().newWorkbook();
    expect(S().workbook.printSettings).toBeUndefined();
    S().restoreRecoveryWorkbook(data.workbook, null);
    expect(S().workbook.printSettings).toMatchObject({ scaleMode: 'custom', scalePercent: 80, header: true });
  });
});
