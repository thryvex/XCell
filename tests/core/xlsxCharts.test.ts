import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { exportToXlsxBuffer, importFromXlsxBuffer, xlsxHasCharts } from '../../src/core/xlsx';
import { createEmptyWorkbook } from '../../src/core/model';

// ไฟล์จาก scripts/chart-spike/out/ บน phase-5d (มี xl/charts/)
const fixture = (name: string) => new Uint8Array(fs.readFileSync(path.join(__dirname, '../fixtures', name)));

describe('ตรวจไฟล์ .xlsx ที่มีกราฟจากชื่อไฟล์ในซิป', () => {
  it('เจอกราฟในไฟล์ที่มี xl/charts/', async () => {
    expect(await xlsxHasCharts(fixture('chart-bar.xlsx'))).toBe(true);
    expect(await xlsxHasCharts(fixture('sample-b-cross-sheet.xlsx'))).toBe(true);
  });

  it('ไม่เจอกราฟในไฟล์ที่แอปบันทึก และไม่พังกับข้อมูลที่ไม่ใช่ซิป', async () => {
    const wb = createEmptyWorkbook();
    wb.sheets[0].cells['A1'] = { raw: '1' };
    expect(await xlsxHasCharts(await exportToXlsxBuffer(wb))).toBe(false);
    expect(await xlsxHasCharts(new Uint8Array([1, 2, 3, 4]))).toBe(false);
  });

  it('ไม่นับโฟลเดอร์ xl/charts/ ที่ว่างเปล่า', async () => {
    const zip = new JSZip();
    zip.folder('xl/charts');
    zip.file('xl/workbook.xml', '<workbook/>');
    expect(await xlsxHasCharts(await zip.generateAsync({ type: 'uint8array' }))).toBe(false);
  });

  it('เปิดไฟล์ที่มีกราฟ: ตั้ง sourceHasCharts และข้อมูลในเซลยังอ่านได้', async () => {
    const wb = await importFromXlsxBuffer(fixture('chart-bar.xlsx'));
    expect(wb.sourceHasCharts).toBe(true);
    expect(wb.unsupportedFeatures ?? []).not.toContain('แผนภูมิ/กราฟ');
    expect(wb.sheets[0].name).toBe('ข้อมูล');
    expect(wb.sheets[0].cells['B2']?.raw).toBe('4500');
  });

  it('เปิดไฟล์ที่กราฟอยู่ชีตอื่น (ชื่อชีตไทยมีเว้นวรรค) ก็ตรวจเจอ', async () => {
    const wb = await importFromXlsxBuffer(fixture('sample-b-cross-sheet.xlsx'));
    expect(wb.sourceHasCharts).toBe(true);
    expect(wb.sheets.map((s) => s.name)).toContain('สรุป กราฟ');
  });

  it('ไฟล์ที่ไม่มีกราฟไม่ตั้ง sourceHasCharts', async () => {
    const wb = createEmptyWorkbook();
    wb.sheets[0].cells['A1'] = { raw: 'ไม่มีกราฟ' };
    const back = await importFromXlsxBuffer(await exportToXlsxBuffer(wb));
    expect(back.sourceHasCharts).toBeUndefined();
  });

  it('บันทึกไฟล์ที่เปิดมาแล้วกราฟหายจริง (เหตุผลที่ต้องเตือน) และไม่เขียน sourceHasCharts ลงไฟล์', async () => {
    const wb = await importFromXlsxBuffer(fixture('chart-bar.xlsx'));
    const saved = await exportToXlsxBuffer(wb);
    expect(await xlsxHasCharts(saved)).toBe(false);
    expect((await importFromXlsxBuffer(saved)).sourceHasCharts).toBeUndefined();
  });
});
