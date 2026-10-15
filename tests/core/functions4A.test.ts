import { describe, it, expect } from 'vitest';
import { Evaluator } from '../../src/core/evaluator';
import { convertShorthand } from '../../src/core/shorthand';
import { dateToSerial, serialToDate } from '../../src/core/numfmt';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import { createEmptyWorkbook, createEmptySheet } from '../../src/core/model';
import { Sheet, CellValue } from '../../src/core/model';

function createSheetWithGrid(data: Record<string, string | number>): Sheet {
  const sheet = createEmptySheet();
  for (const [key, val] of Object.entries(data)) {
    sheet.cells[key] = { raw: String(val) };
  }
  return sheet;
}

function evalAt(sheet: Sheet, formula: string): CellValue {
  const evaluator = new Evaluator(sheet);
  return evaluator.evaluateFormula(formula);
}

describe('Phase 4A: Conditional, Date, and Text Functions', () => {
  describe('COUNTIF', () => {
    it('handles numeric equality criteria (number and string "=100")', () => {
      const sheet = createSheetWithGrid({
        A1: 100,
        A2: 200,
        A3: 100,
        A4: 50,
      });
      expect(evalAt(sheet, '=COUNTIF(A1:A4, 100)')).toBe(2);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "=100")')).toBe(2);
    });

    it('handles numeric comparison criteria (>, >=, <, <=, <>)', () => {
      const sheet = createSheetWithGrid({
        A1: 10,
        A2: 50,
        A3: 100,
        A4: 200,
      });
      expect(evalAt(sheet, '=COUNTIF(A1:A4, ">50")')).toBe(2);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, ">=50")')).toBe(3);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "<50")')).toBe(1);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "<=50")')).toBe(2);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "<>50")')).toBe(3);
    });

    it('handles text criteria, case-insensitivity, and exact matches', () => {
      const sheet = createSheetWithGrid({
        A1: 'Apple',
        A2: 'banana',
        A3: 'APPLE',
        A4: 'orange',
      });
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "Apple")')).toBe(2);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "=banana")')).toBe(1);
      expect(evalAt(sheet, '=COUNTIF(A1:A4, "<>apple")')).toBe(2);
    });

    it('handles Thai criteria and wildcards (* and ?)', () => {
      const sheet = createSheetWithGrid({
        A1: 'อาหารคาว',
        A2: 'อาหารหวาน',
        A3: 'เครื่องดื่ม',
        A4: 'ของหวาน',
        A5: 'อาหารจาน',
      });
      expect(evalAt(sheet, '=COUNTIF(A1:A5, "อาหาร*")')).toBe(3);
      expect(evalAt(sheet, '=COUNTIF(A1:A5, "*หวาน")')).toBe(2);
      expect(evalAt(sheet, '=COUNTIF(A1:A5, "อาหาร???")')).toBe(2); // อาหารคาว, อาหารจาน (3 Thai chars each after อาหาร)
      expect(evalAt(sheet, '=COUNTIF(A1:A5, "เครื่องดื่ม")')).toBe(1);
    });

    it('supports criteria concatenated with cell references and DATE formulas', () => {
      const sheet = createSheetWithGrid({
        A1: 100,
        B1: 50,
        B2: 100,
        B3: 150,
      });
      expect(evalAt(sheet, '=COUNTIF(B1:B3, ">="&A1)')).toBe(2);
    });
  });

  describe('SUMIF', () => {
    it('sums with 2 arguments (criteria range is sum range)', () => {
      const sheet = createSheetWithGrid({
        A1: 10,
        A2: 20,
        A3: 30,
        A4: 40,
      });
      expect(evalAt(sheet, '=SUMIF(A1:A4, ">=25")')).toBe(70);
    });

    it('sums with 3 arguments (separate sum range)', () => {
      const sheet = createSheetWithGrid({
        A1: 'อาหาร',
        A2: 'เครื่องดื่ม',
        A3: 'อาหาร',
        A4: 'บริการ',
        B1: 150,
        B2: 50,
        B3: 250,
        B4: 100,
      });
      expect(evalAt(sheet, '=SUMIF(A1:A4, "อาหาร", B1:B4)')).toBe(400);
      expect(evalAt(sheet, '=SUMIF(A1:A4, "<>อาหาร", B1:B4)')).toBe(150);
    });

    it('ignores text in sum range without erroring', () => {
      const sheet = createSheetWithGrid({
        A1: 'A',
        A2: 'A',
        B1: 'not-a-number',
        B2: 200,
      });
      expect(evalAt(sheet, '=SUMIF(A1:A2, "A", B1:B2)')).toBe(200);
    });
  });

  describe('AVERAGEIF', () => {
    it('calculates average of matching cells', () => {
      const sheet = createSheetWithGrid({
        A1: 'A',
        A2: 'B',
        A3: 'A',
        B1: 10,
        B2: 50,
        B3: 30,
      });
      expect(evalAt(sheet, '=AVERAGEIF(A1:A3, "A", B1:B3)')).toBe(20);
    });

    it('returns #DIV/0! when no cells match criteria', () => {
      const sheet = createSheetWithGrid({
        A1: 'A',
        A2: 'B',
        B1: 10,
        B2: 20,
      });
      expect(evalAt(sheet, '=AVERAGEIF(A1:A2, "C", B1:B2)')).toBe('#DIV/0!');
    });
  });

  describe('COUNTIFS & SUMIFS', () => {
    it('COUNTIFS evaluates multiple conditions', () => {
      const sheet = createSheetWithGrid({
        A1: 'อาหาร',
        A2: 'อาหาร',
        A3: 'เครื่องดื่ม',
        A4: 'อาหาร',
        B1: 100,
        B2: 200,
        B3: 150,
        B4: 50,
      });
      expect(evalAt(sheet, '=COUNTIFS(A1:A4, "อาหาร", B1:B4, ">=100")')).toBe(2);
      expect(evalAt(sheet, '=COUNTIFS(A1:A4, "อาหาร", B1:B4, "<100")')).toBe(1);
    });

    it('COUNTIFS returns #VALUE! if criteria ranges have mismatched dimensions', () => {
      const sheet = createSheetWithGrid({
        A1: 1,
        A2: 2,
        B1: 1,
        B2: 2,
        B3: 3,
      });
      expect(evalAt(sheet, '=COUNTIFS(A1:A2, 1, B1:B3, 1)')).toBe('#VALUE!');
    });

    it('SUMIFS sums based on multiple criteria ranges', () => {
      const sheet = createSheetWithGrid({
        A1: 100,
        A2: 200,
        A3: 300,
        A4: 400,
        B1: 'ก',
        B2: 'ก',
        B3: 'ข',
        B4: 'ก',
        C1: 'Active',
        C2: 'Inactive',
        C3: 'Active',
        C4: 'Active',
      });
      expect(evalAt(sheet, '=SUMIFS(A1:A4, B1:B4, "ก", C1:C4, "Active")')).toBe(500); // 100 + 400
    });

    it('SUMIFS returns #VALUE! if sum range and criteria ranges have mismatched dimensions', () => {
      const sheet = createSheetWithGrid({
        A1: 10,
        A2: 20,
        B1: 'X',
        B2: 'X',
        B3: 'X',
      });
      expect(evalAt(sheet, '=SUMIFS(A1:A2, B1:B3, "X")')).toBe('#VALUE!');
    });
  });

  describe('DATE, YEAR, MONTH, DAY, TODAY, NOW', () => {
    it('DATE generates correct serial numbers matching numfmt', () => {
      const sheet = createEmptySheet();
      // 2026-10-12
      const serial = evalAt(sheet, '=DATE(2026, 10, 12)') as number;
      expect(typeof serial).toBe('number');
      const d = serialToDate(serial);
      expect(d.getFullYear()).toBe(2026);
      expect(d.getMonth() + 1).toBe(10);
      expect(d.getDate()).toBe(12);
    });

    it('DATE handles overflow of months and days like Google Sheets', () => {
      const sheet = createEmptySheet();
      // DATE(2026, 13, 1) -> 2027-01-01
      const serial1 = evalAt(sheet, '=DATE(2026, 13, 1)') as number;
      const d1 = serialToDate(serial1);
      expect(d1.getFullYear()).toBe(2027);
      expect(d1.getMonth() + 1).toBe(1);
      expect(d1.getDate()).toBe(1);

      // DATE(2026, 1, 32) -> 2026-02-01
      const serial2 = evalAt(sheet, '=DATE(2026, 1, 32)') as number;
      const d2 = serialToDate(serial2);
      expect(d2.getFullYear()).toBe(2026);
      expect(d2.getMonth() + 1).toBe(2);
      expect(d2.getDate()).toBe(1);
    });

    it('YEAR, MONTH, DAY extract parts from serial and date strings', () => {
      const sheet = createSheetWithGrid({
        A1: dateToSerial(new Date(2026, 9, 12)), // month is 0-indexed in JS (9 = Oct)
        A2: '2026-10-12',
      });
      expect(evalAt(sheet, '=YEAR(A1)')).toBe(2026);
      expect(evalAt(sheet, '=MONTH(A1)')).toBe(10);
      expect(evalAt(sheet, '=DAY(A1)')).toBe(12);

      expect(evalAt(sheet, '=YEAR(A2)')).toBe(2026);
      expect(evalAt(sheet, '=MONTH(A2)')).toBe(10);
      expect(evalAt(sheet, '=DAY(A2)')).toBe(12);
    });

    it('TODAY and NOW return numeric serial numbers', () => {
      const sheet = createEmptySheet();
      const todaySerial = evalAt(sheet, '=TODAY()') as number;
      const nowSerial = evalAt(sheet, '=NOW()') as number;
      expect(typeof todaySerial).toBe('number');
      expect(typeof nowSerial).toBe('number');
      expect(todaySerial).toBeGreaterThan(40000);
      expect(nowSerial).toBeGreaterThanOrEqual(todaySerial);
    });

    it('allows criteria with DATE comparison', () => {
      const sheet = createSheetWithGrid({
        A1: dateToSerial(new Date(2026, 8, 1)), // 2026-09-01
        A2: dateToSerial(new Date(2026, 9, 15)), // 2026-10-15
        A3: dateToSerial(new Date(2026, 10, 1)), // 2026-11-01
      });
      expect(evalAt(sheet, '=COUNTIF(A1:A3, ">="&DATE(2026, 10, 1))')).toBe(2);
    });
  });

  describe('TEXT function', () => {
    it('formats numbers with allowed formats: 0, 0.00, #,##0, #,##0.00', () => {
      const sheet = createSheetWithGrid({
        A1: 1234.56,
      });
      expect(evalAt(sheet, '=TEXT(A1, "0")')).toBe('1235');
      expect(evalAt(sheet, '=TEXT(A1, "0.00")')).toBe('1234.56');
      expect(evalAt(sheet, '=TEXT(A1, "#,##0")')).toBe('1,235');
      expect(evalAt(sheet, '=TEXT(A1, "#,##0.00")')).toBe('1,234.56');
    });

    it('formats percentages with allowed formats: 0%, 0.00%', () => {
      const sheet = createSheetWithGrid({
        A1: 0.125,
      });
      expect(evalAt(sheet, '=TEXT(A1, "0%")')).toBe('13%');
      expect(evalAt(sheet, '=TEXT(A1, "0.00%")')).toBe('12.50%');
    });

    it('formats dates with allowed formats: dd/mm/yyyy, mm/yyyy, yyyy-mm-dd, d/m/yyyy', () => {
      const serial = dateToSerial(new Date(2026, 4, 5)); // 2026-05-05
      const sheet = createSheetWithGrid({
        A1: serial,
      });
      expect(evalAt(sheet, '=TEXT(A1, "dd/mm/yyyy")')).toBe('05/05/2026');
      expect(evalAt(sheet, '=TEXT(A1, "mm/yyyy")')).toBe('05/2026');
      expect(evalAt(sheet, '=TEXT(A1, "yyyy-mm-dd")')).toBe('2026-05-05');
      expect(evalAt(sheet, '=TEXT(A1, "d/m/yyyy")')).toBe('5/5/2026');
    });

    it('returns #VALUE! for unsupported format patterns', () => {
      const sheet = createSheetWithGrid({
        A1: 123,
      });
      expect(evalAt(sheet, '=TEXT(A1, "mmm dd, yyyy")')).toBe('#VALUE!');
      expect(evalAt(sheet, '=TEXT(A1, "$#,##0")')).toBe('#VALUE!');
    });
  });

  describe('CONCAT & LEN', () => {
    it('CONCAT concatenates multiple arguments and ranges', () => {
      const sheet = createSheetWithGrid({
        A1: 'Hello',
        A2: ' ',
        A3: 'World',
      });
      expect(evalAt(sheet, '=CONCAT("Sheet", "Easy")')).toBe('SheetEasy');
      expect(evalAt(sheet, '=CONCAT(A1, A2, A3)')).toBe('Hello World');
      expect(evalAt(sheet, '=CONCAT(A1:A3)')).toBe('Hello World');
    });

    it('& operator behaves identically to CONCAT', () => {
      const sheet = createSheetWithGrid({
        A1: 'ยอดรวม: ',
        A2: 500,
      });
      expect(evalAt(sheet, '=A1 & A2')).toBe('ยอดรวม: 500');
    });

    it('LEN correctly counts characters including Thai Unicode code units matching Google Sheets', () => {
      const sheet = createSheetWithGrid({
        A1: 'SheetEasy',
        A2: 'สวัสดี', // ส (1) + ว (1) + ั (1) + ส (1) + ด (1) + ี (1) = 6 code units
        A3: '',
      });
      expect(evalAt(sheet, '=LEN(A1)')).toBe(9);
      expect(evalAt(sheet, '=LEN(A2)')).toBe(6);
      expect(evalAt(sheet, '=LEN(A3)')).toBe(0);
    });
  });

  describe('Shorthand Thai Aliases & Expansion', () => {
    it('expands unambiguous Thai aliases to uppercase formula names', () => {
      expect(convertShorthand('รวมถ้า(a1:a10, ">0")')).toBe('=SUMIF(a1:a10, ">0")');
      expect(convertShorthand('นับถ้า(a1:a10, "ผ่าน")')).toBe('=COUNTIF(a1:a10, "ผ่าน")');
      expect(convertShorthand('เฉลี่ยถ้า(a1:a10, ">0")')).toBe('=AVERAGEIF(a1:a10, ">0")');
    });

    it('expands lowercase function calls without = prefix', () => {
      expect(convertShorthand('sumif(a1:a10, ">0")')).toBe('=SUMIF(a1:a10, ">0")');
      expect(convertShorthand('countif(a1:a10, "x")')).toBe('=COUNTIF(a1:a10, "x")');
      expect(convertShorthand('today()')).toBe('=TODAY()');
      expect(convertShorthand('date(2026,10,1)')).toBe('=DATE(2026,10,1)');
    });
  });

  describe('.xlsx Round-trip with new Phase 4A formulas', () => {
    it('saves and reloads formulas for SUMIF, COUNTIF, TODAY, DATE, TEXT, CONCAT', async () => {
      const wb = createEmptyWorkbook();
      const sheet = wb.sheets[0];
      sheet.cells['A1'] = { raw: '100' };
      sheet.cells['A2'] = { raw: '200' };
      sheet.cells['A3'] = { raw: '300' };
      sheet.cells['B1'] = { raw: '=SUMIF(A1:A3, ">100")' };
      sheet.cells['B2'] = { raw: '=COUNTIF(A1:A3, ">=200")' };
      sheet.cells['B3'] = { raw: '=DATE(2026, 10, 12)' };
      sheet.cells['B4'] = { raw: '=TEXT(1234.5, "#,##0.00")' };
      sheet.cells['B5'] = { raw: '=CONCAT("A", "B")' };
      sheet.cells['B6'] = { raw: '=LEN("สวัสดี")' };

      const buffer = await exportToXlsxBuffer(wb);
      const reloaded = await importFromXlsxBuffer(buffer);

      const reloadedSheet = reloaded.sheets[0];
      expect(reloadedSheet.cells['B1']?.raw).toBe('=SUMIF(A1:A3, ">100")');
      expect(reloadedSheet.cells['B2']?.raw).toBe('=COUNTIF(A1:A3, ">=200")');
      expect(reloadedSheet.cells['B3']?.raw).toBe('=DATE(2026, 10, 12)');
      expect(reloadedSheet.cells['B4']?.raw).toBe('=TEXT(1234.5, "#,##0.00")');
      expect(reloadedSheet.cells['B5']?.raw).toBe('=CONCAT("A", "B")');
      expect(reloadedSheet.cells['B6']?.raw).toBe('=LEN("สวัสดี")');

      // Re-evaluate in the reloaded sheet
      const evaluator = new Evaluator(reloadedSheet);
      expect(evaluator.getCellValue('B1')).toBe(500);
      expect(evaluator.getCellValue('B2')).toBe(2);
      expect(evaluator.getCellValue('B4')).toBe('1,234.50');
      expect(evaluator.getCellValue('B5')).toBe('AB');
      expect(evaluator.getCellValue('B6')).toBe(6);
    });
  });
});
