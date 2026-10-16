import { describe, it, expect } from 'vitest';
import { createLedgerWorkbook, createInventoryWorkbook } from '../../src/templates';
import { Evaluator } from '../../src/core/evaluator';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import JSZip from 'jszip';

describe('Templates (Phase 3D)', () => {
  describe('Ledger Template (รายรับรายจ่าย)', () => {
    it('creates ledger workbook with correct structure and styling', () => {
      const wb = createLedgerWorkbook();
      // 4D: เทมเพลตมี 2 ชีต "บันทึก" + "สรุป"
      expect(wb.sheets.map((s) => s.name)).toEqual(['บันทึก', 'สรุป']);
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
    const HEADERS = ['ลำดับ', 'ชื่อ', 'น้ำหนัก', 'เปอร์เซ็น', 'น้ำหนักสุทธิ', 'ราคา', 'รวม'];

    it('creates inventory workbook with correct structure and styling', () => {
      const wb = createInventoryWorkbook();
      expect(wb.sheets.length).toBe(1);
      const sheet = wb.sheets[0];
      expect(sheet.name).toBe('รายการสินค้า');

      // หัวตาราง A1:G1 สีพื้น #93C47D ไม่หนา รูปแบบทั่วไป
      HEADERS.forEach((text, i) => {
        const cell = sheet.cells[`${'ABCDEFG'[i]}1`];
        expect(cell?.raw).toBe(text);
        expect(cell?.style).toEqual({ fill: '#93c47d' });
        expect(cell?.fmt).toBeUndefined();
      });
      expect(sheet.cells['H1']).toBeUndefined();
      expect(sheet.cells['I1']).toEqual({ raw: 'รวมทั้งหมด', style: { fill: '#6d9eeb' } });
      expect(sheet.cells['J1']).toEqual({ raw: '=SUM(G:G)', style: { fill: '#ffffff' } });

      // ตัวอย่าง 3 แถว พร้อมสูตร
      for (const [r, name] of [[2, 'ตัวอย่าง 1'], [3, 'ตัวอย่าง 2'], [4, 'ตัวอย่าง 3']] as const) {
        expect(sheet.cells[`A${r}`]?.raw).toBe(String(r - 1));
        expect(sheet.cells[`B${r}`]?.raw).toBe(name);
        expect(sheet.cells[`E${r}`]?.raw).toBe(`=C${r}*D${r}%`);
        expect(sheet.cells[`G${r}`]?.raw).toBe(`=E${r}*F${r}`);
      }
      expect(['C2', 'D2', 'F2', 'C3', 'D3', 'F3', 'C4', 'D4', 'F4'].map((k) => sheet.cells[k]?.raw)).toEqual([
        '100', '38', '40', '250', '21', '40', '80', '66', '41',
      ]);
      expect(Object.keys(sheet.cells).some((k) => /^[A-Z]+([5-9]|\d{2,})$/.test(k))).toBe(false);

      // ความกว้าง: A ≈ 6.6, B ≈ 17.9, H ≈ 4.4 หน่วย (7 px ต่อหน่วย); อื่นๆ ค่าเริ่มต้น
      expect(sheet.colWidths).toEqual({ 0: 46, 1: 125, 7: 31 });
      expect(sheet.frozenRows).toBe(0);
      expect(sheet.frozenCols).toBe(0);
      expect(sheet.merges).toEqual([]);
      expect(sheet.filter).toBeUndefined();
    });

    it('computes initial inventory sample data and total sum correctly', () => {
      const sheet = createInventoryWorkbook().sheets[0];
      const evaluator = new Evaluator(sheet);
      // น้ำหนักสุทธิ = น้ำหนัก × เปอร์เซ็น%
      expect(evaluator.getCellValue('E2')).toBeCloseTo(38);
      expect(evaluator.getCellValue('E3')).toBeCloseTo(52.5);
      expect(evaluator.getCellValue('E4')).toBeCloseTo(52.8);
      // รวม = น้ำหนักสุทธิ × ราคา
      expect(evaluator.getCellValue('G2')).toBeCloseTo(1520);
      expect(evaluator.getCellValue('G3')).toBeCloseTo(2100);
      expect(evaluator.getCellValue('G4')).toBeCloseTo(2164.8);
      expect(evaluator.getCellValue('J1')).toBeCloseTo(5784.8);
    });

    it('includes a new row 5 (and rows far below) in the J1 total without a cycle', () => {
      const sheet = createInventoryWorkbook().sheets[0];
      const evaluator = new Evaluator(sheet);
      expect(evaluator.getCellValue('J1')).toBeCloseTo(5784.8);

      Object.assign(sheet.cells, {
        C5: { raw: '50' },
        D5: { raw: '10' },
        E5: { raw: '=C5*D5%' },
        F5: { raw: '20' },
        G5: { raw: '=E5*F5' },
      });
      for (const k of ['C5', 'D5', 'E5', 'F5', 'G5']) evaluator.invalidate(k, 0);
      expect(evaluator.getCellValue('G5')).toBeCloseTo(100);
      expect(evaluator.getCellValue('J1')).toBeCloseTo(5884.8);

      sheet.cells['G10000'] = { raw: '1' };
      evaluator.invalidate('G10000', 0);
      expect(evaluator.getCellValue('J1')).toBeCloseTo(5885.8);
    });

    it('round-trips inventory workbook through .xlsx buffer with values and colors', async () => {
      const wb = createInventoryWorkbook();
      const buf = await exportToXlsxBuffer(wb);
      const sheet = (await importFromXlsxBuffer(buf)).sheets[0];
      const original = wb.sheets[0];

      expect(sheet.name).toBe('รายการสินค้า');
      for (const key of Object.keys(original.cells)) {
        expect(sheet.cells[key]?.raw, key).toBe(original.cells[key].raw);
        expect(sheet.cells[key]?.style?.fill, key).toBe(original.cells[key].style?.fill);
        expect(sheet.cells[key]?.style?.bold, key).toBeFalsy();
      }
      expect(sheet.colWidths).toEqual({ 0: 46, 1: 125, 7: 31 });
      expect(sheet.frozenRows).toBe(0);
      expect(sheet.merges).toEqual([]);

      const evaluator = new Evaluator(sheet);
      expect(evaluator.getCellValue('E4')).toBeCloseTo(52.8);
      expect(evaluator.getCellValue('J1')).toBeCloseTo(5784.8);
    });

    it('writes the agreed column widths and colors into the .xlsx', async () => {
      const zip = await JSZip.loadAsync(await exportToXlsxBuffer(createInventoryWorkbook()));
      const xml = await zip.file('xl/worksheets/sheet1.xml')!.async('string');
      expect(xml).toMatch(/<col min="1" max="1" width="6.57"/);
      expect(xml).toMatch(/<col min="2" max="2" width="17.86"/);
      expect(xml).toMatch(/<col min="8" max="8" width="4.43"/);
      expect(xml).not.toContain('<pane');
      expect(xml).not.toContain('<mergeCell');
      expect(xml).not.toContain('<autoFilter');
      const styles = await zip.file('xl/styles.xml')!.async('string');
      expect(styles).toContain('FF93C47D');
      expect(styles).toContain('FF6D9EEB');
    });
  });
});
