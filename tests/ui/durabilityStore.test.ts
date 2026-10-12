import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { createEmptyWorkbook } from '../../src/core/model';

describe('useSheetStore durability & file state (Phase 3C)', () => {
  beforeEach(() => {
    useSheetStore.getState().newWorkbook();
  });

  it('updates title with ● when isDirty is true and removes ● on markSaved', () => {
    const store = useSheetStore.getState();
    expect(useSheetStore.getState().isDirty).toBe(false);

    // Edit a cell to make sheet dirty
    store.startEditing('12345');
    store.commitEditing();

    expect(useSheetStore.getState().isDirty).toBe(true);

    // Mark saved
    useSheetStore.getState().markSaved('C:\\docs\\budget.xlsx');
    expect(useSheetStore.getState().isDirty).toBe(false);
    expect(useSheetStore.getState().fileName).toBe('budget');
    expect(useSheetStore.getState().filePath).toBe('C:\\docs\\budget.xlsx');
  });

  it('restores a recovery workbook and marks isDirty true', () => {
    const recoveredWb = createEmptyWorkbook();
    recoveredWb.sheets[0].name = 'กู้คืนงาน';
    recoveredWb.sheets[0].cells['A1'] = { raw: 'ข้อมูลสำคัญ' };
    recoveredWb.unsupportedFeatures = ['ตาราง Pivot'];

    const store = useSheetStore.getState();
    store.restoreRecoveryWorkbook(recoveredWb, 'C:\\work\\old.xlsx');

    const state = useSheetStore.getState();
    expect(state.isDirty).toBe(true);
    expect(state.filePath).toBe('C:\\work\\old.xlsx');
    expect(state.fileName).toBe('old');
    expect(state.workbook.sheets[0].name).toBe('กู้คืนงาน');
    expect(state.getRawValue(0, 0)).toBe('ข้อมูลสำคัญ');
    expect(state.unsupportedFeatures).toEqual(['ตาราง Pivot']);
  });

  it('handles unsupported features banner dismissal and overwrite warning flag', () => {
    const store = useSheetStore.getState();
    useSheetStore.setState({ unsupportedFeatures: ['รูปภาพ', 'แผนภูมิ/กราฟ'] });

    expect(useSheetStore.getState().unsupportedFeatures).toHaveLength(2);
    expect(useSheetStore.getState().hasWarnedUnsupportedOverwrite).toBe(false);

    store.setHasWarnedUnsupportedOverwrite(true);
    expect(useSheetStore.getState().hasWarnedUnsupportedOverwrite).toBe(true);

    store.dismissUnsupportedBanner();
    expect(useSheetStore.getState().unsupportedFeatures).toHaveLength(0);
  });
});
