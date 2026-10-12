import { describe, it, expect } from 'vitest';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import { createEmptyWorkbook } from '../../src/core/model';

describe('src/core/xlsx (Round-trip test)', () => {
  it('exports and imports formulas, values, formats, merges, and styles', async () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0];
    sheet.name = 'รายรับรายจ่าย';

    // Cell values
    sheet.cells['A1'] = { raw: 'วันที่', style: { bold: true, align: 'center' } };
    sheet.cells['B1'] = { raw: 'รายการ', style: { bold: true } };
    sheet.cells['C1'] = { raw: 'ยอดเงิน', style: { bold: true, align: 'right' } };

    sheet.cells['A2'] = { raw: '1/10/2026' };
    sheet.cells['B2'] = { raw: 'เงินเดือน' };
    sheet.cells['C2'] = { raw: '25000', fmt: 'baht' };

    sheet.cells['A3'] = { raw: '2/10/2026' };
    sheet.cells['B3'] = { raw: 'ค่าอาหาร' };
    sheet.cells['C3'] = { raw: '150', fmt: 'baht' };

    sheet.cells['B4'] = { raw: 'รวมทั้งสิ้น', style: { bold: true } };
    sheet.cells['C4'] = { raw: '=SUM(C2:C3)', fmt: 'baht', style: { bold: true } };

    // Merged cell A5:B5
    sheet.merges.push({ r1: 4, c1: 0, r2: 4, c2: 1 });
    sheet.cells['A5'] = { raw: 'หมายเหตุ' };

    // Frozen row 1
    sheet.frozenRows = 1;

    // Export to binary buffer
    const buffer = await exportToXlsxBuffer(wb);
    expect(buffer).toBeInstanceOf(Uint8Array);
    expect(buffer.length).toBeGreaterThan(100);

    // Import back
    const importedWb = await importFromXlsxBuffer(buffer);
    const importedSheet = importedWb.sheets[0];

    expect(importedSheet.name).toBe('รายรับรายจ่าย');
    expect(importedSheet.cells['A1'].raw).toBe('วันที่');
    expect(importedSheet.cells['A1'].style?.bold).toBe(true);

    expect(importedSheet.cells['C2'].raw).toBe('25000');
    expect(importedSheet.cells['C2'].fmt).toBe('baht');

    expect(importedSheet.cells['C4'].raw).toBe('=SUM(C2:C3)');
    expect(importedSheet.cells['C4'].fmt).toBe('baht');
    expect(importedSheet.cells['C4'].style?.bold).toBe(true);

    // Merge roundtrip
    expect(importedSheet.merges.length).toBe(1);
    expect(importedSheet.merges[0]).toEqual({ r1: 4, c1: 0, r2: 4, c2: 1 });

    // Frozen row roundtrip
    expect(importedSheet.frozenRows).toBe(1);
  });
});
