import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore, isCsvPath } from '../../src/ui/store';
import { CSV_MAX_ROWS } from '../../src/core/csv';
import { th } from '../../src/i18n/th';

const S = () => useSheetStore.getState();
const utf8 = (text: string) => new TextEncoder().encode(text);

describe('นำเข้า CSV ผ่าน store', () => {
  beforeEach(() => {
    S().newWorkbook();
  });

  it('opening a .csv creates a new unnamed workbook with one sheet named after the file', () => {
    S().beginCsvImport(utf8('สินค้า;ราคา\nข้าว;"1,234.50"\n'), 'C:\\งาน\\ยอดขาย.csv', 'open');
    const p = S().pendingCsvImport!;
    expect(p).toMatchObject({ mode: 'open', encoding: 'utf-8', delimiter: ';', autoConvert: true });
    S().confirmCsvImport();
    expect(S().pendingCsvImport).toBeNull();
    expect(S().filePath).toBeNull();
    expect(S().fileName).toBe('ยอดขาย');
    expect(S().isDirty).toBe(false);
    expect(S().workbook.sheets.map((s) => s.name)).toEqual(['ยอดขาย']);
    expect(S().getComputedValue(1, 1)).toBe(1234.5);
  });

  it('options can be changed before importing (auto-convert off keeps text)', () => {
    S().beginCsvImport(utf8('a,007,5%'), 'x.csv', 'open');
    S().updateCsvImport({ autoConvert: false });
    S().confirmCsvImport();
    expect(S().getComputedValue(0, 1)).toBe('007');
    expect(S().getComputedValue(0, 2)).toBe('5%');
  });

  it('importing as a new sheet adds a uniquely named sheet and can be undone', () => {
    S().beginCsvImport(utf8('1,2'), 'C:\\ชีต1.csv', 'newSheet');
    S().confirmCsvImport();
    expect(S().workbook.sheets.map((s) => s.name)).toEqual(['ชีต1', 'ชีต1 (2)']);
    expect(S().activeSheetIndex).toBe(1);
    S().undo();
    expect(S().workbook.sheets.map((s) => s.name)).toEqual(['ชีต1']);
  });

  it('asks before importing more than 10,000 rows, then imports what fits', () => {
    const big = Array.from({ length: CSV_MAX_ROWS + 3 }, (_, i) => `r${i}`).join('\n');
    S().beginCsvImport(utf8(big), 'big.csv', 'open');
    S().confirmCsvImport();
    expect(S().pendingCsvImport?.step).toBe('limitWarning');
    S().confirmCsvImport();
    const cells = S().workbook.sheets[0].cells;
    expect(cells[`A${CSV_MAX_ROWS}`].raw).toBe(`r${CSV_MAX_ROWS - 1}`);
    expect(cells[`A${CSV_MAX_ROWS + 1}`]).toBeUndefined();
  });

  it('cancelling the limit warning imports nothing', () => {
    const big = Array.from({ length: CSV_MAX_ROWS + 1 }, () => 'x').join('\n');
    S().beginCsvImport(utf8(big), 'big.csv', 'newSheet');
    S().confirmCsvImport();
    expect(S().pendingCsvImport?.step).toBe('limitWarning');
    S().cancelCsvImport();
    expect(S().workbook.sheets).toHaveLength(1);
  });

  it('empty and binary files show a Thai message instead of crashing', () => {
    S().beginCsvImport(new Uint8Array(), 'empty.csv', 'open');
    expect(S().notice).toBe(th.csv.emptyFile);
    S().dismissNotice();
    S().beginCsvImport(utf8('\r\n\r\n'), 'blank.csv', 'open');
    expect(S().notice).toBe(th.csv.emptyFile);
    S().dismissNotice();
    S().beginCsvImport(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00]), 'fake.csv', 'open');
    expect(S().notice).toBe(th.csv.notText);
    expect(S().pendingCsvImport).toBeNull();
    S().dismissNotice();
  });

  it('a recovered workbook that came from a CSV is never bound to the .csv path', () => {
    S().beginCsvImport(utf8('a'), 'C:\\x\\data.csv', 'open');
    S().confirmCsvImport();
    S().restoreRecoveryWorkbook(S().workbook, 'C:\\x\\data.csv');
    expect(S().filePath).toBeNull();
    expect(isCsvPath('C:\\x\\data.CSV')).toBe(true);
    expect(isCsvPath('C:\\x\\data.xlsx')).toBe(false);
  });
});

describe('ส่งออก CSV ผ่าน store', () => {
  beforeEach(() => {
    S().newWorkbook();
  });

  it('exports the active sheet with BOM and a file + sheet default name', () => {
    S().startEditing('ราคา');
    S().commitEditing();
    S().setActiveCell(0, 1);
    S().startEditing('=2*3');
    S().commitEditing();
    const { bytes, defaultName, hasFilter } = S().getCsvExport();
    expect(Array.from(bytes.slice(0, 3))).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder().decode(bytes.slice(3))).toBe('ราคา,6');
    expect(defaultName).toBe('ตารางบันทึก - ชีต1.csv');
    expect(hasFilter).toBe(false);
  });

  it('reports an active filter so the UI can warn that hidden rows are exported too', () => {
    S().startEditing('หัว');
    S().commitEditing();
    S().setActiveCell(1, 0);
    S().startEditing('1');
    S().commitEditing();
    S().setActiveCell(0, 0);
    S().toggleFilter();
    S().setColumnFilter(0, { kind: 'values', hiddenValues: ['1'] });
    const { bytes, hasFilter } = S().getCsvExport();
    expect(hasFilter).toBe(true);
    expect(new TextDecoder().decode(bytes.slice(3))).toBe('หัว\r\n1');
  });
});
