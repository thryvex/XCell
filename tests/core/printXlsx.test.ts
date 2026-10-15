import { describe, it, expect } from 'vitest';
import { exportToXlsxBuffer, importFromXlsxBuffer, workbookToExcelJS } from '../../src/core/xlsx';
import { createEmptySheet, Workbook, DEFAULT_PRINT_SETTINGS, PrintSettings } from '../../src/core/model';

function wb(ps: Partial<PrintSettings>, frozenRows = 0): Workbook {
  const sheet = createEmptySheet('ชีต1');
  sheet.cells['A1'] = { raw: 'x' };
  sheet.frozenRows = frozenRows;
  return { sheets: [sheet], activeSheet: 0, printSettings: { ...DEFAULT_PRINT_SETTINGS, ...ps } };
}

describe('.xlsx pageSetup (ตัวเลือกพิมพ์)', () => {
  it('writes orientation, paper, fit-to-width, margins, gridlines and print titles', async () => {
    const excel = await workbookToExcelJS(
      wb({ orientation: 'landscape', paper: 'Letter', margins: 'narrow', gridlines: true }, 2)
    );
    const ps = excel.worksheets[0].pageSetup;
    expect(ps.orientation).toBe('landscape');
    expect(ps.paperSize).toBe(1);
    expect(ps.fitToPage).toBe(true);
    expect(ps.fitToWidth).toBe(1);
    expect(ps.fitToHeight).toBe(0);
    expect(ps.margins?.left).toBeCloseTo(0.25, 3);
    expect(ps.showGridLines).toBe(true);
    expect(ps.printTitlesRow).toBe('1:2');
  });

  it('round-trips the settings that .xlsx can hold', async () => {
    const cases: Partial<PrintSettings>[] = [
      { orientation: 'landscape', paper: 'A5', scaleMode: 'fitWidth', margins: 'wide', gridlines: true },
      { orientation: 'portrait', paper: 'Letter', scaleMode: 'custom', scalePercent: 75, margins: 'narrow' },
      { paper: 'A4', scaleMode: '100', margins: 'normal', gridlines: false },
    ];
    for (const c of cases) {
      const back = await importFromXlsxBuffer(await exportToXlsxBuffer(wb(c, 1)));
      const ps = back.printSettings!;
      expect(ps.orientation).toBe(c.orientation ?? 'portrait');
      expect(ps.paper).toBe(c.paper);
      expect(ps.scaleMode).toBe(c.scaleMode);
      if (c.scaleMode === 'custom') expect(ps.scalePercent).toBe(75);
      expect(ps.margins).toBe(c.margins);
      expect(ps.gridlines).toBe(c.gridlines ?? false);
      expect(ps.repeatFrozen).toBe(true); // printTitlesRow อ่านกลับได้
    }
  });

  it('frozen rows without print titles in the file → no repeat', async () => {
    const back = await importFromXlsxBuffer(await exportToXlsxBuffer(wb({ repeatFrozen: false }, 1)));
    expect(back.printSettings!.repeatFrozen).toBe(false);
  });

  it('a workbook without print settings writes no custom page setup', async () => {
    const plain = wb({});
    delete plain.printSettings;
    const excel = await workbookToExcelJS(plain);
    expect(excel.worksheets[0].pageSetup.printTitlesRow).toBeUndefined();
    expect(excel.worksheets[0].pageSetup.fitToPage).toBeFalsy();
  });
});
