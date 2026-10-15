import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import {
  exportToXlsxBuffer,
  importFromXlsxBuffer,
  workbookToExcelJS,
} from '../../src/core/xlsx';
import { Workbook, createEmptySheet } from '../../src/core/model';
import { Evaluator } from '../../src/core/evaluator';

function buildWorkbook(): Workbook {
  const summary = createEmptySheet('สรุป');
  const data = createEmptySheet('ข้อมูล ปี 69');
  const hidden = createEmptySheet('ค่าคงที่');
  const plain = createEmptySheet('Data_2');

  data.cells['A1'] = { raw: 'อาหาร', style: { bold: true, fill: '#fde68a' } };
  data.cells['B1'] = { raw: '120', fmt: 'baht' };
  data.cells['A2'] = { raw: 'เดินทาง' };
  data.cells['B2'] = { raw: '80', fmt: 'baht' };
  data.cells['A3'] = { raw: 'อาหาร' };
  data.cells['B3'] = { raw: '30', fmt: 'baht' };
  data.merges = [{ r1: 5, c1: 0, r2: 5, c2: 2 }];
  data.colWidths = { 0: 160, 1: 120 };
  data.rowHeights = { 0: 40 };
  data.frozenRows = 1;

  hidden.hidden = 'hidden';
  hidden.cells['A1'] = { raw: '0.07' };

  plain.cells['A1'] = { raw: '5' };

  summary.cells['A1'] = { raw: "=SUM('ข้อมูล ปี 69'!B1:B3)", fmt: 'baht', style: { italic: true } };
  summary.cells['A2'] = { raw: `=SUMIF('ข้อมูล ปี 69'!A1:A3,"อาหาร",'ข้อมูล ปี 69'!B1:B3)` };
  summary.cells['A3'] = { raw: '=A1*ค่าคงที่!$A$1' }; // ชื่อไทยไม่มีเครื่องหมายคำพูด → ต้องใส่ตอนบันทึก
  summary.cells['A4'] = { raw: '=Data_2!A1*2' };
  summary.frozenCols = 1;

  return { sheets: [summary, data, hidden, plain], activeSheet: 1 };
}

describe('.xlsx หลายชีต (Phase 4B ขั้น 3)', () => {
  it('writes sheet order, hidden state, active tab and quoted cross-sheet formulas', async () => {
    const excel = await workbookToExcelJS(buildWorkbook());
    expect(excel.worksheets.map((ws) => ws.name)).toEqual(['สรุป', 'ข้อมูล ปี 69', 'ค่าคงที่', 'Data_2']);
    expect(excel.worksheets[2].state).toBe('hidden');
    expect((excel.views[0] as any).activeTab).toBe(1);

    const summary = excel.getWorksheet('สรุป')!;
    expect(summary.getCell('A1').formula).toBe("SUM('ข้อมูล ปี 69'!B1:B3)");
    expect(summary.getCell('A3').formula).toBe("A1*'ค่าคงที่'!$A$1");
    expect(summary.getCell('A4').formula).toBe('Data_2!A1*2');
    // ค่าที่คำนวณข้ามชีตถูกเขียนเป็น cache
    expect(summary.getCell('A1').result).toBe(230);
    expect(summary.getCell('A2').result).toBe(150);
    expect(summary.getCell('A3').result).toBeCloseTo(16.1);
  });

  it('round-trips all sheets with formulas, styles, merges, sizes, freeze, hidden and active sheet', async () => {
    const original = buildWorkbook();
    const buffer = await exportToXlsxBuffer(original);
    const loaded = await importFromXlsxBuffer(buffer);

    expect(loaded.sheets.map((s) => s.name)).toEqual(['สรุป', 'ข้อมูล ปี 69', 'ค่าคงที่', 'Data_2']);
    expect(loaded.activeSheet).toBe(1);
    expect(loaded.sheets.map((s) => s.hidden)).toEqual([undefined, undefined, 'hidden', undefined]);
    expect(loaded.unsupportedFeatures).toBeUndefined();

    const [summary, data, hidden] = loaded.sheets;
    expect(summary.cells['A1'].raw).toBe("=SUM('ข้อมูล ปี 69'!B1:B3)");
    expect(summary.cells['A1'].fmt).toBe('baht');
    expect(summary.cells['A1'].style?.italic).toBe(true);
    expect(summary.cells['A3'].raw).toBe("=A1*'ค่าคงที่'!$A$1");
    expect(summary.frozenCols).toBe(1);

    expect(data.cells['A1'].style).toMatchObject({ bold: true, fill: '#fde68a' });
    expect(data.merges).toEqual([{ r1: 5, c1: 0, r2: 5, c2: 2 }]);
    expect(data.colWidths[0]).toBe(160);
    expect(data.colWidths[1]).toBe(120);
    expect(data.rowHeights[0]).toBe(40);
    expect(data.frozenRows).toBe(1);

    expect(hidden.cells['A1'].raw).toBe('0.07');

    const ev = new Evaluator(loaded);
    expect(ev.getCellValue('A1', 0)).toBe(230);
    expect(ev.getCellValue('A2', 0)).toBe(150);
    expect(ev.getCellValue('A3', 0)).toBeCloseTo(16.1);
    expect(ev.getCellValue('A4', 0)).toBe(10);
  });

  it('reads veryHidden sheets and keeps them hidden on save', async () => {
    const excel = new ExcelJS.Workbook();
    excel.addWorksheet('หลัก').getCell('A1').value = { formula: "'ลับ'!A1+1", result: 2 } as any;
    const secret = excel.addWorksheet('ลับ', { state: 'veryHidden' });
    secret.getCell('A1').value = 41;
    const buf = new Uint8Array(await excel.xlsx.writeBuffer());

    const loaded = await importFromXlsxBuffer(buf);
    expect(loaded.sheets[1].hidden).toBe('veryHidden');
    expect(new Evaluator(loaded).getCellValue('A1', 0)).toBe(42);

    const again = await workbookToExcelJS(loaded);
    expect(again.worksheets[1].state).toBe('veryHidden');
  });

  it('falls back to the first sheet when the file has no active tab', async () => {
    const excel = new ExcelJS.Workbook();
    excel.addWorksheet('A');
    excel.addWorksheet('B');
    const loaded = await importFromXlsxBuffer(new Uint8Array(await excel.xlsx.writeBuffer()));
    expect(loaded.activeSheet).toBe(0);
  });
});
