import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import {
  detectUnsupportedFeatures,
  hasUnsupportedFunction,
  excelJSToWorkbook,
  workbookToExcelJS,
  exportToXlsxBuffer,
  importFromXlsxBuffer,
} from '../../src/core/xlsx';
import { createEmptyWorkbook } from '../../src/core/model';
import { Evaluator } from '../../src/core/evaluator';

describe('src/core/xlsx Durability & Unsupported Features (SPEC 6.4)', () => {
  it('detects unsupported features in an ExcelJS workbook', () => {
    const wb = new ExcelJS.Workbook();
    const ws1 = wb.addWorksheet('Sheet1');
    const wsHidden = wb.addWorksheet('SecretSheet');
    wsHidden.state = 'hidden';

    // Mock conditional formatting
    (ws1 as any).conditionalFormattings = [{ ref: 'A1:A10', rules: [] }];

    // Mock data validation
    (ws1 as any).dataValidations = { model: { A1: { type: 'list', formulae: ['"A,B,C"'] } } };

    // Mock pivot table
    (ws1 as any).pivotTables = [{ name: 'Pivot1' }];

    // Mock drawings/charts
    (ws1 as any).drawings = [{ name: 'Chart1' }];

    // Mock media/images
    (ws1 as any).getImages = () => [{ imageId: 'img1', range: 'B2:D5' }];

    const detected = detectUnsupportedFeatures(wb);

    // 4B: ชีตซ่อนรองรับแล้ว (เก็บไว้ตามเดิม) จึงไม่อยู่ในรายการส่วนที่ไม่รองรับอีกต่อไป
    expect(detected).not.toContain('แผ่นงานที่ซ่อนอยู่ (Hidden Sheets)');
    expect(detected).toContain('การจัดรูปแบบตามเงื่อนไข (Conditional Formatting)');
    expect(detected).toContain('การตรวจสอบความถูกต้องของข้อมูล (Data Validation)');
    expect(detected).toContain('ตาราง Pivot');
    expect(detected).toContain('แผนภูมิ/กราฟ');
    expect(detected).toContain('รูปภาพ');
  });

  it('returns an empty array when workbook has no unsupported features', () => {
    const wb = new ExcelJS.Workbook();
    wb.addWorksheet('CleanSheet');
    const detected = detectUnsupportedFeatures(wb);
    expect(detected).toEqual([]);
  });

  it('identifies supported vs unsupported formulas accurately', () => {
    // Supported core functions
    expect(hasUnsupportedFunction('=SUM(A1:A10)')).toBe(false);
    expect(hasUnsupportedFunction('=AVERAGE(B1:B5)')).toBe(false);
    expect(hasUnsupportedFunction('=MEAN(B1:B5)')).toBe(false);
    expect(hasUnsupportedFunction('=IF(A1>0, 1, 0)')).toBe(false);
    expect(hasUnsupportedFunction('=ROUND(C1, 2)')).toBe(false);
    expect(hasUnsupportedFunction('=A1 + B1 * 10')).toBe(false);

    // Unsupported functions
    expect(hasUnsupportedFunction('=VLOOKUP(A1, B1:C10, 2, FALSE)')).toBe(true);
    expect(hasUnsupportedFunction('=XLOOKUP(A1, B1:B10, C1:C10)')).toBe(true);
    expect(hasUnsupportedFunction('=QUERY(A1:B10, "select A")')).toBe(true);
    expect(hasUnsupportedFunction('=INDEX(A1:B10, 2, 2)')).toBe(true);
  });

  it('preserves unsupported formula text and cached value in round-trip (SPEC 6.4)', async () => {
    // 1. Create an Excel workbook with an unsupported formula and cached result
    const excelWb = new ExcelJS.Workbook();
    const ws = excelWb.addWorksheet('DataSheet');

    const cell = ws.getCell('B2');
    cell.value = {
      formula: 'VLOOKUP(A1, X1:Z100, 3, FALSE)',
      result: 4500,
    };

    // 2. Import into XCell
    const sheetEasyWb = await excelJSToWorkbook(excelWb);
    const sheetModel = sheetEasyWb.sheets[0];
    const cellRecord = sheetModel.cells['B2'];

    expect(cellRecord).toBeDefined();
    expect(cellRecord.raw).toBe('=VLOOKUP(A1, X1:Z100, 3, FALSE)');
    expect(cellRecord.unsupportedFormula).toBe(true);
    expect(cellRecord.cachedValue).toBe(4500);

    // 3. Evaluator must return the cached value directly
    const evaluator = new Evaluator(sheetModel);
    expect(evaluator.getCellValue('B2')).toBe(4500);

    // 4. Export back to ExcelJS
    const exportedExcelWb = await workbookToExcelJS(sheetEasyWb);
    const exportedWs = exportedExcelWb.getWorksheet('DataSheet')!;
    const exportedCell = exportedWs.getCell('B2');

    expect(exportedCell.type).toBe(ExcelJS.ValueType.Formula);
    const formulaVal = exportedCell.value as { formula: string; result: any };
    expect(formulaVal.formula).toBe('VLOOKUP(A1, X1:Z100, 3, FALSE)');
    expect(formulaVal.result).toBe(4500);
  });
});
