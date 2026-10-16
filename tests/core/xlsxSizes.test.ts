import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  exportToXlsxBuffer,
  importFromXlsxBuffer,
  pxToColWidth,
  colWidthToPx,
  pxToRowHeight,
  rowHeightToPx,
} from '../../src/core/xlsx';
import { createEmptyWorkbook, createEmptySheet } from '../../src/core/model';

async function sheetXml(buffer: Uint8Array, n = 1): Promise<string> {
  const zip = await JSZip.loadAsync(buffer);
  return zip.file(`xl/worksheets/sheet${n}.xml`)!.async('string');
}

describe('ความกว้างคอลัมน์/ความสูงแถวใน .xlsx (7 px ต่อหน่วย)', () => {
  it('แปลง px เป็นหน่วยความกว้าง 7 px ต่อหน่วย ทศนิยม 2 ตำแหน่ง', () => {
    expect(pxToColWidth(70)).toBe(10);
    expect(pxToColWidth(110)).toBe(15.71);
    expect(pxToColWidth(200)).toBe(28.57);
    expect(colWidthToPx(15.71)).toBe(110);
    expect(colWidthToPx(28.57)).toBe(200);
  });

  it('ความกว้างต่ำสุดตอนเขียนคือ 30 px', () => {
    expect(pxToColWidth(10)).toBe(4.29);
    expect(colWidthToPx(pxToColWidth(10))).toBe(30);
  });

  it('แปลงความสูง px เป็น pt (× 0.75) ทศนิยม 2 ตำแหน่ง', () => {
    expect(pxToRowHeight(28)).toBe(21);
    expect(pxToRowHeight(30)).toBe(22.5);
    expect(pxToRowHeight(37)).toBe(27.75);
    expect(rowHeightToPx(22.5)).toBe(30);
  });

  it('round-trip ความกว้าง 30–600 px และความสูง 18–200 px ได้ค่าเดิมทุกค่า', async () => {
    const wb = createEmptyWorkbook();
    const sheet = wb.sheets[0];
    const widths: number[] = [];
    for (let px = 30, c = 0; px <= 600; px += 7, c++) {
      sheet.colWidths[c] = px + (c % 7); // ให้มีทุกเศษของการหาร 7
      widths.push(sheet.colWidths[c]);
    }
    for (let px = 18, r = 0; px <= 200; px++, r++) sheet.rowHeights[r] = px;
    sheet.cells['A1'] = { raw: 'x' };

    const back = (await importFromXlsxBuffer(await exportToXlsxBuffer(wb))).sheets[0];
    expect(back.colWidths).toEqual(sheet.colWidths);
    expect(back.rowHeights).toEqual(sheet.rowHeights);
    expect(widths.length).toBeGreaterThan(80);
  });

  it('เขียนความกว้าง/ความสูงที่ตั้งไว้และค่าเริ่มต้นลง XML ของชีต', async () => {
    const wb = createEmptyWorkbook();
    wb.sheets[0].colWidths[1] = 200; // B
    wb.sheets[0].rowHeights[2] = 30; // แถว 3
    wb.sheets[0].cells['A1'] = { raw: '1' };
    const xml = await sheetXml(await exportToXlsxBuffer(wb));

    const fmt = /<sheetFormatPr\b[^>]*\/>/.exec(xml)?.[0] ?? '';
    expect(fmt).toContain('defaultColWidth="15.71"'); // 110 px
    expect(fmt).toContain('defaultRowHeight="21"'); // 28 px
    expect(xml).toMatch(/<col min="2" max="2" width="28.57"/);
    expect(xml).toMatch(/<row r="3"[^>]*ht="22.5"/);
  });

  it('เขียนค่าเริ่มต้นลงทุกชีต', async () => {
    const wb = createEmptyWorkbook();
    wb.sheets.push(createEmptySheet('ชีต2'));
    const buffer = await exportToXlsxBuffer(wb);
    for (const n of [1, 2]) {
      expect(await sheetXml(buffer, n)).toMatch(/<sheetFormatPr\b[^>]*defaultColWidth="15.71"[^>]*\/>/);
    }
  });

  it('คอลัมน์/แถวที่ไม่ได้ตั้ง ไม่ถูกเขียนเป็นค่าเฉพาะ และเปิดกลับมาก็ไม่มีค่า', async () => {
    const wb = createEmptyWorkbook();
    wb.sheets[0].cells['C3'] = { raw: 'x' };
    const buffer = await exportToXlsxBuffer(wb);
    expect(await sheetXml(buffer)).not.toContain('<cols>');
    const back = (await importFromXlsxBuffer(buffer)).sheets[0];
    expect(back.colWidths).toEqual({});
    expect(back.rowHeights).toEqual({});
  });

  it('ไฟล์เก่าที่บันทึกแบบ 8 px ต่อหน่วย (25 หน่วย = 200 px เดิม) เปิดได้ 175 px และบันทึกซ้ำได้ 25 หน่วยเท่าเดิม', async () => {
    const old = new ExcelJS.Workbook();
    const ws = old.addWorksheet('เก่า');
    ws.getColumn(2).width = 25;
    ws.getCell('A1').value = 1;
    const oldBuffer = new Uint8Array(await old.xlsx.writeBuffer());

    const wb = await importFromXlsxBuffer(oldBuffer);
    expect(wb.sheets[0].colWidths[1]).toBe(175);

    const again = await sheetXml(await exportToXlsxBuffer(wb));
    expect(again).toMatch(/<col min="2" max="2" width="25"/);
  });
});
