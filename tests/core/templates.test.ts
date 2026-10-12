import { describe, it, expect } from 'vitest';
import { createLedgerWorkbook, createInventoryWorkbook } from '../../src/templates';
import { Evaluator } from '../../src/core/evaluator';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';

describe('Templates (Phase 3D)', () => {
  describe('Ledger Template (รายรับรายจ่าย)', () => {
    it('creates ledger workbook with correct structure and styling', () => {
      const wb = createLedgerWorkbook();
      expect(wb.sheets.length).toBe(1);
      const sheet = wb.sheets[0];
      expect(sheet.name).toBe('บันทึก');
      expect(sheet.frozenRows).toBe(1);

      // Check header row (A1:F1)
      expect(sheet.cells['A1']?.raw).toBe('วันที่');
      expect(sheet.cells['B1']?.raw).toBe('รายการ');
      expect(sheet.cells['C1']?.raw).toBe('หมวด');
      expect(sheet.cells['D1']?.raw).toBe('รายรับ');
      expect(sheet.cells['E1']?.raw).toBe('รายจ่าย');
      expect(sheet.cells['F1']?.raw).toBe('คงเหลือ');

      // Header styling
      expect(sheet.cells['A1']?.style?.bold).toBe(true);
      expect(sheet.cells['A1']?.style?.align).toBe('center');
      expect(sheet.cells['A1']?.style?.fill).toBe('#e2e8f0');

      // Column widths (B is wider than A for Thai text)
      expect(sheet.colWidths?.[1]).toBeGreaterThan(sheet.colWidths?.[0] || 0);

      // Summary on right (H1:I3)
      expect(sheet.cells['H1']?.raw).toBe('รวมรายรับ');
      expect(sheet.cells['I1']?.raw).toBe('=SUM(D2:D1000)');
      expect(sheet.cells['I1']?.fmt).toBe('baht');

      expect(sheet.cells['H2']?.raw).toBe('รวมรายจ่าย');
      expect(sheet.cells['I2']?.raw).toBe('=SUM(E2:E1000)');
      expect(sheet.cells['I2']?.fmt).toBe('baht');

      expect(sheet.cells['H3']?.raw).toBe('คงเหลือ');
      expect(sheet.cells['I3']?.raw).toBe('=I1-I2');
      expect(sheet.cells['I3']?.fmt).toBe('baht');
    });

    it('computes initial sample data and right summary correctly', () => {
      const wb = createLedgerWorkbook();
      const sheet = wb.sheets[0];
      const evaluator = new Evaluator(sheet);

      // Sample rows:
      // Row 2: Income 25000, Expense 0 -> Balance 25000
      expect(evaluator.getCellValue('F2')).toBe(25000);
      // Row 3: Income 0, Expense 150 -> Balance 25000 - 150 = 24850
      expect(evaluator.getCellValue('F3')).toBe(24850);
      // Row 4: Income 0, Expense 80 -> Balance 24850 - 80 = 24770
      expect(evaluator.getCellValue('F4')).toBe(24770);

      // Summary:
      // Total Income = 25000
      expect(evaluator.getCellValue('I1')).toBe(25000);
      // Total Expense = 150 + 80 = 230
      expect(evaluator.getCellValue('I2')).toBe(230);
      // Balance = 25000 - 230 = 24770
      expect(evaluator.getCellValue('I3')).toBe(24770);
    });

    it('allows deleting sample data without breaking summary formulas', () => {
      const wb = createLedgerWorkbook();
      const sheet = wb.sheets[0];

      // Clear sample rows 2-4 (like a user selecting and hitting Delete)
      for (const row of ['2', '3', '4']) {
        sheet.cells[`A${row}`] = { raw: '', fmt: 'date' };
        sheet.cells[`B${row}`] = { raw: '' };
        sheet.cells[`C${row}`] = { raw: '' };
        sheet.cells[`D${row}`] = { raw: '', fmt: 'baht' };
        sheet.cells[`E${row}`] = { raw: '', fmt: 'baht' };
        sheet.cells[`F${row}`] = { raw: '', fmt: 'baht' };
      }

      const evaluator = new Evaluator(sheet);

      // Summary formulas should still evaluate gracefully to 0
      expect(evaluator.getCellValue('I1')).toBe(0);
      expect(evaluator.getCellValue('I2')).toBe(0);
      expect(evaluator.getCellValue('I3')).toBe(0);
    });

    it('round-trips through .xlsx buffer preserving formulas, formats, and frozen rows', async () => {
      const wb = createLedgerWorkbook();
      const buf = await exportToXlsxBuffer(wb);
      expect(buf.length).toBeGreaterThan(0);

      const imported = await importFromXlsxBuffer(buf);
      const sheet = imported.sheets[0];

      expect(sheet.name).toBe('บันทึก');
      expect(sheet.frozenRows).toBe(1);
      expect(sheet.cells['A1']?.raw).toBe('วันที่');
      expect(sheet.cells['F2']?.raw).toBe('=D2-E2');
      expect(sheet.cells['I1']?.raw).toBe('=SUM(D2:D1000)');
      expect(sheet.cells['I1']?.fmt).toBe('baht');
      expect(sheet.cells['I3']?.raw).toBe('=I1-I2');

      const evaluator = new Evaluator(sheet);
      expect(evaluator.getCellValue('I3')).toBe(24770);
    });
  });

  describe('Inventory Template (รายการสินค้า)', () => {
    it('creates inventory workbook with correct structure and styling', () => {
      const wb = createInventoryWorkbook();
      expect(wb.sheets.length).toBe(1);
      const sheet = wb.sheets[0];
      expect(sheet.name).toBe('รายการสินค้า');
      expect(sheet.frozenRows).toBe(1);

      // Headers (A1:E1)
      expect(sheet.cells['A1']?.raw).toBe('ลำดับ');
      expect(sheet.cells['B1']?.raw).toBe('สินค้า');
      expect(sheet.cells['C1']?.raw).toBe('จำนวน');
      expect(sheet.cells['D1']?.raw).toBe('ราคาต่อหน่วย');
      expect(sheet.cells['E1']?.raw).toBe('รวม');

      // Summary on right (G1:H1)
      expect(sheet.cells['G1']?.raw).toBe('รวมทั้งหมด');
      expect(sheet.cells['H1']?.raw).toBe('=SUM(E2:E1000)');
      expect(sheet.cells['H1']?.fmt).toBe('baht');
    });

    it('computes initial inventory sample data and total sum correctly', () => {
      const wb = createInventoryWorkbook();
      const sheet = wb.sheets[0];
      const evaluator = new Evaluator(sheet);

      // Row 2: 5 * 35 = 175
      expect(evaluator.getCellValue('E2')).toBe(175);
      // Row 3: 10 * 15 = 150
      expect(evaluator.getCellValue('E3')).toBe(150);
      // Row 4: 3 * 45 = 135
      expect(evaluator.getCellValue('E4')).toBe(135);

      // Total sum: 175 + 150 + 135 = 460
      expect(evaluator.getCellValue('H1')).toBe(460);
    });

    it('round-trips inventory workbook through .xlsx buffer', async () => {
      const wb = createInventoryWorkbook();
      const buf = await exportToXlsxBuffer(wb);
      expect(buf.length).toBeGreaterThan(0);

      const imported = await importFromXlsxBuffer(buf);
      const sheet = imported.sheets[0];

      expect(sheet.name).toBe('รายการสินค้า');
      expect(sheet.frozenRows).toBe(1);
      expect(sheet.cells['E2']?.raw).toBe('=C2*D2');
      expect(sheet.cells['H1']?.raw).toBe('=SUM(E2:E1000)');
      expect(sheet.cells['H1']?.fmt).toBe('baht');

      const evaluator = new Evaluator(sheet);
      expect(evaluator.getCellValue('H1')).toBe(460);
    });
  });
});
