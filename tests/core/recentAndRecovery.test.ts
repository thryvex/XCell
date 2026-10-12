import { describe, it, expect } from 'vitest';
import { updateRecentFilesList, removeRecentFileItem } from '../../src/core/recentFiles';
import {
  RecoverySessionData,
  serializeRecoveryData,
  deserializeRecoveryData,
} from '../../src/core/recovery';
import { createEmptyWorkbook } from '../../src/core/model';

describe('src/core/recentFiles (SPEC 6.2 - Recent Files List)', () => {
  it('adds files to the front and moves duplicates to the front', () => {
    let list: string[] = [];
    list = updateRecentFilesList(list, 'C:\\docs\\file1.xlsx');
    list = updateRecentFilesList(list, 'C:\\docs\\file2.xlsx');
    expect(list).toEqual(['C:\\docs\\file2.xlsx', 'C:\\docs\\file1.xlsx']);

    // Duplicate should move to top
    list = updateRecentFilesList(list, 'C:\\docs\\file1.xlsx');
    expect(list).toEqual(['C:\\docs\\file1.xlsx', 'C:\\docs\\file2.xlsx']);
  });

  it('limits the list to 10 items maximum', () => {
    let list: string[] = [];
    for (let i = 1; i <= 15; i++) {
      list = updateRecentFilesList(list, `C:\\data\\test_${i}.xlsx`);
    }

    expect(list.length).toBe(10);
    expect(list[0]).toBe('C:\\data\\test_15.xlsx');
    expect(list[9]).toBe('C:\\data\\test_6.xlsx');
  });

  it('removes an item from the recent files list', () => {
    const initial = ['C:\\a.xlsx', 'C:\\b.xlsx', 'C:\\c.xlsx'];
    const updated = removeRecentFileItem(initial, 'c:\\b.xlsx');
    expect(updated).toEqual(['C:\\a.xlsx', 'C:\\c.xlsx']);
  });
});

describe('src/core/recovery (SPEC 6.3 - Auto-recovery Serialization)', () => {
  it('serializes and deserializes full workbook recovery snapshot', () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0];
    sheet.name = 'บันทึกค่าใช้จ่าย';
    sheet.cells['A1'] = { raw: 'ค่าอาหาร', style: { bold: true, color: '#ef4444' } };
    sheet.cells['B1'] = { raw: '120', fmt: 'baht' };
    sheet.merges.push({ r1: 2, c1: 0, r2: 2, c2: 1 });
    sheet.colWidths[0] = 150;
    sheet.rowHeights[0] = 32;
    sheet.frozenRows = 1;
    sheet.frozenCols = 0;

    const session: RecoverySessionData = {
      id: 'recovery-123456',
      originalPath: 'C:\\Work\\Expense.xlsx',
      fileName: 'Expense',
      savedAt: 1776000000000,
      workbook: wb,
    };

    const serialized = serializeRecoveryData(session);
    expect(typeof serialized).toBe('string');

    const restored = deserializeRecoveryData(serialized);
    expect(restored).not.toBeNull();
    expect(restored?.id).toBe('recovery-123456');
    expect(restored?.originalPath).toBe('C:\\Work\\Expense.xlsx');
    expect(restored?.fileName).toBe('Expense');

    const restoredSheet = restored!.workbook.sheets[0];
    expect(restoredSheet.name).toBe('บันทึกค่าใช้จ่าย');
    expect(restoredSheet.cells['A1'].raw).toBe('ค่าอาหาร');
    expect(restoredSheet.cells['A1'].style?.bold).toBe(true);
    expect(restoredSheet.cells['A1'].style?.color).toBe('#ef4444');
    expect(restoredSheet.cells['B1'].fmt).toBe('baht');
    expect(restoredSheet.merges).toEqual([{ r1: 2, c1: 0, r2: 2, c2: 1 }]);
    expect(restoredSheet.colWidths[0]).toBe(150);
    expect(restoredSheet.rowHeights[0]).toBe(32);
    expect(restoredSheet.frozenRows).toBe(1);
  });
});
