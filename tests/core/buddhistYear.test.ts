import { describe, it, expect, afterEach, beforeEach } from 'vitest';
import {
  toGregorianYear,
  parseDateString,
  formatDate,
  formatDisplayValue,
  setDisplayYearMode,
  dateToSerial,
  convertTypedDate,
  editTextForCell,
} from '../../src/core/numfmt';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet, Workbook } from '../../src/core/model';
import { findMatches, DEFAULT_SEARCH_OPTIONS } from '../../src/core/search';
import { sortRange } from '../../src/core/sort';
import { uniqueColumnValues, createFilter } from '../../src/core/filter';
import { convertCsvValue, sheetToCsvRows } from '../../src/core/csv';
import { exportToXlsxBuffer, importFromXlsxBuffer, workbookToExcelJS } from '../../src/core/xlsx';
import { parsePreferences, mergePreferences, DEFAULT_PREFERENCES } from '../../src/core/preferences';
import { useSheetStore } from '../../src/ui/store';

const serial = (y: number, m: number, d: number) => dateToSerial(new Date(y, m - 1, d));
const OCT15 = serial(2026, 10, 15);

afterEach(() => setDisplayYearMode('ce'));

describe('การแปลง พ.ศ. ↔ ค.ศ.', () => {
  it('years >= 2400 are Buddhist era, below are Gregorian', () => {
    expect(toGregorianYear(2569)).toBe(2026);
    expect(toGregorianYear(2400)).toBe(1857);
    expect(toGregorianYear(2399)).toBe(2399);
    expect(toGregorianYear(2026)).toBe(2026);
  });

  it('formatDate shows CE or BE (+543) for the same serial', () => {
    expect(formatDate(OCT15, 'ce')).toBe('15/10/2026');
    expect(formatDate(OCT15, 'be')).toBe('15/10/2569');
    setDisplayYearMode('be');
    expect(formatDate(OCT15)).toBe('15/10/2569');
    expect(formatDisplayValue(OCT15, 'date')).toBe('15/10/2569');
    expect(formatDisplayValue(OCT15, 'general')).toBe(String(OCT15)); // เฉพาะรูปแบบวันที่
  });
});

describe('การพิมพ์วันที่ (parseDateString)', () => {
  it('d/m/yyyy, dd-mm-yyyy and yyyy-mm-dd accept both eras', () => {
    expect(parseDateString('15/10/2569')).toBe(OCT15);
    expect(parseDateString('15/10/2026')).toBe(OCT15);
    expect(parseDateString('15-10-2569')).toBe(OCT15);
    expect(parseDateString('2569-10-15')).toBe(OCT15);
    expect(parseDateString('2026-10-15')).toBe(OCT15);
  });

  it('d/m without a year uses the current year; 2-digit years are 20xx', () => {
    expect(parseDateString('15/10')).toBe(serial(new Date().getFullYear(), 10, 15));
    expect(parseDateString('15/10/26')).toBe(OCT15);
  });

  it('leap years: 29 Feb 2567 (= 2024) is valid, 29 Feb 2568 (= 2025) is not', () => {
    expect(parseDateString('29/2/2567')).toBe(serial(2024, 2, 29));
    expect(parseDateString('29/02/2568')).toBeNull();
    expect(parseDateString('29/2/2024')).toBe(serial(2024, 2, 29));
    expect(parseDateString('29/2/2025')).toBeNull();
  });

  it('typed dates become real dates; invalid dates and other text stay text', () => {
    expect(convertTypedDate('15/10/2569')).toEqual({ raw: String(OCT15), fmt: 'date' });
    expect(convertTypedDate('29/2/2568')).toBeNull();
    expect(convertTypedDate('=DATE(2026,10,15)')).toBeNull();
    expect(convertTypedDate('ข้าว')).toBeNull();
    expect(convertTypedDate('1234')).toBeNull();
  });

  it('edit text of a date cell follows the year mode', () => {
    expect(editTextForCell(String(OCT15), 'date')).toBe('15/10/2026');
    setDisplayYearMode('be');
    expect(editTextForCell(String(OCT15), 'date')).toBe('15/10/2569');
    expect(editTextForCell('=TODAY()', 'date')).toBe('=TODAY()');
    expect(editTextForCell('123', 'baht')).toBe('123');
  });
});

describe('สูตรไม่ถูกกระทบจากโหมดปี', () => {
  it('YEAR / DATE / TEXT always use CE', () => {
    setDisplayYearMode('be');
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: String(OCT15), fmt: 'date' };
    const ev = new Evaluator(sheet);
    expect(ev.evaluateFormula('=YEAR(A1)')).toBe(2026);
    expect(ev.evaluateFormula('=DATE(2026,10,15)')).toBe(OCT15);
    expect(ev.evaluateFormula('=TEXT(A1,"dd/mm/yyyy")')).toBe('15/10/2026');
    expect(ev.evaluateFormula('=YEAR(DATE(2569,1,1))')).toBe(2569); // DATE ไม่แปลง พ.ศ. (เหมือน Google Sheets)
  });
});

function datedWorkbook(): Workbook {
  const sheet = createEmptySheet('ชีต1');
  sheet.cells['A1'] = { raw: 'วันที่' };
  sheet.cells['A2'] = { raw: String(OCT15), fmt: 'date' };
  sheet.cells['A3'] = { raw: String(serial(2025, 12, 1)), fmt: 'date' };
  sheet.cells['A4'] = { raw: '1/1/2569', fmt: 'date' }; // วันที่ที่เป็นข้อความ (ไฟล์เก่า) ปี พ.ศ.
  sheet.cells['B2'] = { raw: 'รายการ 2569' };
  return { sheets: [sheet], activeSheet: 0 };
}

describe('ค้นหา / เรียง / ตัวกรอง / CSV / xlsx กับโหมด พ.ศ.', () => {
  it('search matches the displayed year', () => {
    const wb = datedWorkbook();
    const ev = new Evaluator(wb);
    const keys = () =>
      findMatches(wb, ev, '2569', DEFAULT_SEARCH_OPTIONS, 0).map((m) => `${m.r},${m.c}`);
    expect(keys()).toEqual(['1,1', '3,0']); // CE: ข้อความ "รายการ 2569" และวันที่ข้อความ 1/1/2569
    setDisplayYearMode('be');
    expect(keys()).toEqual(['1,0', '1,1', '3,0']);
  });

  it('sorting uses the real date value, not the displayed text', () => {
    setDisplayYearMode('be');
    const wb = datedWorkbook();
    const sheet = wb.sheets[0];
    sortRange(sheet, new Evaluator(wb), 0, { r1: 0, c1: 0, r2: 3, c2: 0 }, [{ col: 0, dir: 'asc' }], true);
    expect([2, 3, 4].map((r) => sheet.cells[`A${r}`].raw)).toEqual([
      String(serial(2025, 12, 1)),
      '1/1/2569',
      String(OCT15),
    ]);
  });

  it('filter-by-values lists dates as displayed', () => {
    setDisplayYearMode('be');
    const wb = datedWorkbook();
    wb.sheets[0].filter = createFilter({ r1: 0, c1: 0, r2: 3, c2: 0 });
    const values = uniqueColumnValues(wb.sheets[0], new Evaluator(wb), 0, 0).map((v) => v.display);
    expect(values).toContain('15/10/2569');
    expect(values).toContain('01/12/2568');
  });

  it('CSV import reads BE years; CSV export uses the displayed year', () => {
    expect(convertCsvValue('15/10/2569', true)).toEqual({ raw: String(OCT15), fmt: 'date' });
    expect(convertCsvValue('2569-10-15', true)).toEqual({ raw: String(OCT15), fmt: 'date' });
    expect(convertCsvValue('29/2/2568', true)).toEqual({ raw: '29/2/2568' });
    const wb = datedWorkbook();
    expect(sheetToCsvRows(wb.sheets[0], new Evaluator(wb), 0)[1][0]).toBe('15/10/2026');
    setDisplayYearMode('be');
    expect(sheetToCsvRows(wb.sheets[0], new Evaluator(wb), 0)[1][0]).toBe('15/10/2569');
  });

  it('.xlsx keeps the CE serial and a plain date format in BE mode', async () => {
    setDisplayYearMode('be');
    const wb = datedWorkbook();
    const excel = await workbookToExcelJS(wb);
    const cell = excel.worksheets[0].getCell('A2');
    expect(cell.numFmt).toBe('dd/mm/yyyy');
    const reloaded = await importFromXlsxBuffer(await exportToXlsxBuffer(wb));
    expect(reloaded.sheets[0].cells['A2']).toMatchObject({ raw: String(OCT15), fmt: 'date' });
  });
});

describe('ค่าของแอป (preferences)', () => {
  it('year mode defaults to CE and survives partial saves', () => {
    expect(DEFAULT_PREFERENCES.yearMode).toBe('ce');
    expect(parsePreferences('{"yearMode":"be"}').yearMode).toBe('be');
    expect(parsePreferences('{"yearMode":"xx"}').yearMode).toBe('ce');
    const merged = mergePreferences(parsePreferences('{"yearMode":"be"}'), { showWelcomeGuide: false });
    expect(merged).toMatchObject({ yearMode: 'be', showWelcomeGuide: false, hideBuddhistSaveNotice: false });
  });
});

describe('การพิมพ์และวางวันที่ใน store', () => {
  const S = () => useSheetStore.getState();
  const type = (r: number, c: number, text: string) => {
    S().setActiveCell(r, c);
    S().startEditing(text);
    S().commitEditing();
  };

  beforeEach(() => {
    S().newWorkbook();
    S().setYearMode('ce', false);
  });
  afterEach(() => S().setYearMode('ce', false));

  it('typing a BE or CE date stores the same real date regardless of mode', () => {
    S().setYearMode('be', false);
    type(0, 0, '15/10/2569');
    type(1, 0, '15-10-2026');
    type(2, 0, '1/10');
    type(3, 0, '29/2/2568');
    const cells = S().workbook.sheets[0].cells;
    expect(cells['A1']).toMatchObject({ raw: String(OCT15), fmt: 'date' });
    expect(cells['A2']).toMatchObject({ raw: String(OCT15), fmt: 'date' });
    expect(cells['A3']).toMatchObject({ raw: String(serial(new Date().getFullYear(), 10, 1)), fmt: 'date' });
    expect(cells['A4'].raw).toBe('29/2/2568');
    expect(cells['A4'].fmt).toBeUndefined();
    expect(S().getDisplayValue(0, 0)).toBe('15/10/2569');
  });

  it('editor shows the date in the current mode and re-committing keeps the serial', () => {
    type(0, 0, '15/10/2026');
    S().setYearMode('be', false);
    S().setActiveCell(0, 0);
    S().startEditing();
    expect(S().editingValue).toBe('15/10/2569');
    expect(S().getEditText(0, 0)).toBe('15/10/2569');
    S().commitEditing();
    expect(S().workbook.sheets[0].cells['A1'].raw).toBe(String(OCT15));
    expect(S().history.canUndo()).toBe(true);
  });

  it('switching the year mode changes only the display', () => {
    type(0, 0, '15/10/2026');
    const before = JSON.stringify(S().workbook.sheets);
    S().setYearMode('be', false);
    expect(S().getDisplayValue(0, 0)).toBe('15/10/2569');
    expect(JSON.stringify(S().workbook.sheets)).toBe(before);
    expect(S().isDirty).toBe(true); // จากการพิมพ์เท่านั้น
  });

  it('pasting from outside converts BE dates the same way', async () => {
    const original = Object.getOwnPropertyDescriptor(globalThis, 'navigator');
    Object.defineProperty(globalThis, 'navigator', {
      value: { clipboard: { readText: async () => '15/10/2569\tข้อความ\n29/2/2568\t1/1/2026' } },
      configurable: true,
    });
    try {
      S().setActiveCell(0, 0);
      await S().paste(false);
      const cells = S().workbook.sheets[0].cells;
      expect(cells['A1']).toMatchObject({ raw: String(OCT15), fmt: 'date' });
      expect(cells['B1'].raw).toBe('ข้อความ');
      expect(cells['A2'].raw).toBe('29/2/2568');
      expect(cells['B2']).toMatchObject({ raw: String(serial(2026, 1, 1)), fmt: 'date' });
    } finally {
      if (original) Object.defineProperty(globalThis, 'navigator', original);
      else delete (globalThis as any).navigator;
    }
  });

  it('CSV export reports that the sheet has dates (for the year notice)', () => {
    type(0, 0, '15/10/2569');
    S().setYearMode('be', false);
    const { bytes, hasDates } = S().getCsvExport();
    expect(hasDates).toBe(true);
    expect(new TextDecoder().decode(bytes.slice(3))).toBe('15/10/2569');
  });
});
