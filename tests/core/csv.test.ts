import { describe, it, expect } from 'vitest';
import {
  writeCsv,
  quoteCsvField,
  encodeUtf8WithBom,
  sheetToCsvRows,
  parseCsv,
  detectEncoding,
  decodeBytes,
  detectDelimiter,
  looksBinary,
  convertCsvValue,
  buildSheetFromCsv,
  csvExceedsLimits,
  sheetNameFromFileName,
  CSV_MAX_ROWS,
} from '../../src/core/csv';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet } from '../../src/core/model';
import { dateToSerial } from '../../src/core/numfmt';

describe('CSV writer', () => {
  it('quotes commas, quotes and newlines; doubles quotes; Thai untouched', () => {
    expect(quoteCsvField('ข้าวผัด')).toBe('ข้าวผัด');
    expect(quoteCsvField('1,234')).toBe('"1,234"');
    expect(quoteCsvField('เขา "ว่า"')).toBe('"เขา ""ว่า"""');
    expect(quoteCsvField('บรรทัด1\nบรรทัด2')).toBe('"บรรทัด1\nบรรทัด2"');
  });

  it('joins rows with CRLF', () => {
    expect(writeCsv([['a', 'b'], ['c', 'd']])).toBe('a,b\r\nc,d');
  });

  it('adds a UTF-8 BOM', () => {
    const bytes = encodeUtf8WithBom('ก');
    expect(Array.from(bytes.slice(0, 3))).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder().decode(bytes.slice(3))).toBe('ก');
  });

  it('exports displayed values: formulas as results, baht/%/dates formatted, errors as standard codes', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'รายการ' };
    sheet.cells['B1'] = { raw: '1234.5', fmt: 'baht' };
    sheet.cells['C1'] = { raw: '=B1*2' };
    sheet.cells['A2'] = { raw: '0.075', fmt: 'percent' };
    sheet.cells['B2'] = { raw: String(dateToSerial(new Date(2026, 9, 1))), fmt: 'date' };
    sheet.cells['C2'] = { raw: '=1/0' };
    const rows = sheetToCsvRows(sheet, new Evaluator(sheet), 0);
    expect(rows).toEqual([
      ['รายการ', '฿1,234.50', '2469'],
      ['7.50%', '01/10/2026', '#DIV/0!'],
    ]);
    expect(writeCsv(rows)).toBe('รายการ,"฿1,234.50",2469\r\n7.50%,01/10/2026,#DIV/0!');
  });

  it('stops at the last cell with data and ignores format-only cells', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['B3'] = { raw: 'x' };
    sheet.cells['F20'] = { raw: '', fmt: 'baht' };
    expect(sheetToCsvRows(sheet, new Evaluator(sheet), 0)).toEqual([
      ['', ''],
      ['', ''],
      ['', 'x'],
    ]);
    expect(sheetToCsvRows(createEmptySheet(), new Evaluator(createEmptySheet()), 0)).toEqual([]);
  });

  it('merged cells: value at top-left only', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'หัวข้อ' };
    sheet.cells['B1'] = { raw: 'ค้าง' }; // ไม่ควรเกิด แต่ถ้ามีต้องไม่ถูกส่งออก
    sheet.cells['C1'] = { raw: 'z' };
    sheet.merges = [{ r1: 0, c1: 0, r2: 0, c2: 1 }];
    expect(sheetToCsvRows(sheet, new Evaluator(sheet), 0)).toEqual([['หัวข้อ', '', 'z']]);
  });
});

describe('CSV reader', () => {
  it('parses every delimiter', () => {
    expect(parseCsv('a,b\r\nc,d', ',')).toEqual([['a', 'b'], ['c', 'd']]);
    expect(parseCsv('a;b\nc;d\n', ';')).toEqual([['a', 'b'], ['c', 'd']]);
    expect(parseCsv('a\tb\nc\td', '\t')).toEqual([['a', 'b'], ['c', 'd']]);
    expect(parseCsv('a||b', '||')).toEqual([['a', 'b']]);
  });

  it('handles quotes, doubled quotes, newlines and delimiters inside quotes', () => {
    const text = '"ชื่อ, สกุล","เขา ""ว่า""","บรรทัด1\r\nบรรทัด2"\r\nx,,y';
    expect(parseCsv(text, ',')).toEqual([
      ['ชื่อ, สกุล', 'เขา "ว่า"', 'บรรทัด1\r\nบรรทัด2'],
      ['x', '', 'y'],
    ]);
  });

  it('keeps empty trailing fields and ignores the final newline', () => {
    expect(parseCsv('a,b,\n', ',')).toEqual([['a', 'b', '']]);
    expect(parseCsv('', ',')).toEqual([]);
  });

  it('detects encoding from BOM, valid UTF-8, or falls back to Windows-874', () => {
    expect(detectEncoding(new Uint8Array([0xef, 0xbb, 0xbf, 0x41]))).toBe('utf-8');
    expect(detectEncoding(new Uint8Array([0xff, 0xfe, 0x41, 0x00]))).toBe('utf-16le');
    expect(detectEncoding(new Uint8Array([0xfe, 0xff, 0x00, 0x41]))).toBe('utf-16be');
    expect(detectEncoding(new TextEncoder().encode('ข้าว,10'))).toBe('utf-8');
    // "ข้าว" ใน Windows-874 (TIS-620): ข=0xA2 ้=0xE9 า=0xD2 ว=0xC7
    const tis = new Uint8Array([0xa2, 0xe9, 0xd2, 0xc7, 0x2c, 0x31, 0x30]);
    expect(detectEncoding(tis)).toBe('windows-874');
    expect(decodeBytes(tis, 'windows-874')).toBe('ข้าว,10');
  });

  it('decodes UTF-16 and strips BOMs', () => {
    const le = new Uint8Array([0xff, 0xfe, 0x01, 0x0e, 0x2c, 0x00, 0x31, 0x00]); // "ก,1"
    expect(decodeBytes(le, 'utf-16le')).toBe('ก,1');
    expect(decodeBytes(encodeUtf8WithBom('ก,1'), 'utf-8')).toBe('ก,1');
  });

  it('flags binary content', () => {
    expect(looksBinary(new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00]), 'utf-8')).toBe(true);
    expect(looksBinary(new TextEncoder().encode('a,b'), 'utf-8')).toBe(false);
  });

  it('detects the delimiter from the sample', () => {
    expect(detectDelimiter('a,b,c\n1,2,3')).toBe(',');
    expect(detectDelimiter('a;b;c\n"1,5";2;3')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2')).toBe('\t');
    expect(detectDelimiter('ข้อความเดียว\nอีกบรรทัด')).toBe(',');
  });
});

describe('การแปลงค่าเมื่อนำเข้า', () => {
  const v = (t: string, auto = true) => convertCsvValue(t, auto);

  it('numbers incl. thousands separators, negatives and decimals', () => {
    expect(v('42')).toEqual({ raw: '42' });
    expect(v('-15.25')).toEqual({ raw: '-15.25' });
    expect(v('1,234.50')).toEqual({ raw: '1234.5', fmt: 'number2' });
    expect(v('1,234')).toEqual({ raw: '1234', fmt: 'integer' });
    expect(v('0')).toEqual({ raw: '0' });
    expect(v('0.5')).toEqual({ raw: '0.5' });
  });

  it('leading zeros and long numbers stay text', () => {
    expect(v('0812345678')).toEqual({ raw: "'0812345678" });
    expect(v('007')).toEqual({ raw: "'007" });
    expect(v('1234567890123456')).toEqual({ raw: "'1234567890123456" });
    expect(v('123456789012345')).toEqual({ raw: '123456789012345' });
  });

  it('percent', () => {
    expect(v('5%')).toEqual({ raw: '0.05', fmt: 'percent' });
    expect(v('12.5%')).toEqual({ raw: '0.125', fmt: 'percent' });
  });

  it('dates d/m/yyyy, dd-mm-yyyy, yyyy-mm-dd (day/month/year)', () => {
    const serial = String(dateToSerial(new Date(2026, 9, 1)));
    expect(v('1/10/2026')).toEqual({ raw: serial, fmt: 'date' });
    expect(v('01-10-2026')).toEqual({ raw: serial, fmt: 'date' });
    expect(v('2026-10-01')).toEqual({ raw: serial, fmt: 'date' });
    expect(v('31/2/2026')).toEqual({ raw: '31/2/2026' }); // วันที่ไม่มีจริง → ข้อความ
  });

  it('text starting with "=" is never a formula', () => {
    expect(v('=SUM(A1:A3)')).toEqual({ raw: "'=SUM(A1:A3)" });
    expect(v('=1+1', false)).toEqual({ raw: "'=1+1" });
  });

  it('with auto-convert off everything stays text exactly', () => {
    expect(v('1,234.50', false)).toEqual({ raw: "'1,234.50" });
    expect(v('5%', false)).toEqual({ raw: '5%' });
    expect(v('ข้าว', false)).toEqual({ raw: 'ข้าว' });
    expect(v('', false)).toBeNull();
  });

  it('builds a sheet whose values evaluate as expected', () => {
    const sheet = buildSheetFromCsv(
      [
        ['รหัส', 'จำนวน', 'ส่วนลด'],
        ['007', '1,234.50', '5%'],
        ['=HACK()', '-3', ''],
      ],
      'นำเข้า',
      true
    );
    const ev = new Evaluator(sheet);
    expect(ev.getCellValue('A2')).toBe('007');
    expect(ev.getCellValue('B2')).toBe(1234.5);
    expect(ev.getCellValue('C2')).toBe(0.05);
    expect(ev.getCellValue('A3')).toBe('=HACK()');
    expect(ev.getCellValue('B3')).toBe(-3);
    expect(sheet.cells['C3']).toBeUndefined();
  });

  it('limits: more than 10,000 rows is detected and truncated only when building', () => {
    const rows = Array.from({ length: CSV_MAX_ROWS + 5 }, (_, i) => [String(i)]);
    expect(csvExceedsLimits(rows)).toBe(true);
    const sheet = buildSheetFromCsv(rows, 'ใหญ่', true);
    expect(sheet.cells[`A${CSV_MAX_ROWS}`].raw).toBe(String(CSV_MAX_ROWS - 1));
    expect(sheet.cells[`A${CSV_MAX_ROWS + 1}`]).toBeUndefined();
    expect(csvExceedsLimits([Array(703).fill('x')])).toBe(true);
    expect(csvExceedsLimits([Array(702).fill('x')])).toBe(false);
  });

  it('sheet name from the file name follows sheet-name rules', () => {
    expect(sheetNameFromFileName('C:\\งาน\\ยอดขาย [ต.ค.].csv')).toBe('ยอดขาย ต.ค.');
    expect(sheetNameFromFileName('x/' + 'ก'.repeat(40) + '.csv')).toHaveLength(31);
    expect(sheetNameFromFileName('???.csv')).toBe('ชีต1');
    expect(sheetNameFromFileName('ข้อมูล.csv', ['ข้อมูล'])).toBe('ข้อมูล (2)');
  });
});

describe('round-trip CSV', () => {
  it('export then import gives the same data (formulas become values)', () => {
    const sheet = createEmptySheet('ชีต1');
    sheet.cells['A1'] = { raw: 'สินค้า' };
    sheet.cells['B1'] = { raw: 'ราคา' };
    sheet.cells['A2'] = { raw: 'ข้าว, หอมมะลิ' };
    sheet.cells['B2'] = { raw: '1234.5', fmt: 'number2' };
    sheet.cells['A3'] = { raw: "'0812345678" };
    sheet.cells['B3'] = { raw: '=B2*2', fmt: 'number2' };
    sheet.cells['A4'] = { raw: 'เขา "ว่า"\nขึ้นบรรทัด' };
    sheet.cells['B4'] = { raw: '0.05', fmt: 'percent' };

    const bytes = encodeUtf8WithBom(writeCsv(sheetToCsvRows(sheet, new Evaluator(sheet), 0)));
    const enc = detectEncoding(bytes);
    const text = decodeBytes(bytes, enc);
    const back = buildSheetFromCsv(parseCsv(text, detectDelimiter(text)), 'กลับ', true);
    const ev = new Evaluator(back);

    expect(enc).toBe('utf-8');
    expect(ev.getCellValue('A2')).toBe('ข้าว, หอมมะลิ');
    expect(ev.getCellValue('B2')).toBe(1234.5);
    expect(ev.getCellValue('A3')).toBe('0812345678');
    expect(back.cells['B3'].raw).toBe('2469'); // สูตรกลายเป็นค่า
    expect(ev.getCellValue('A4')).toBe('เขา "ว่า"\nขึ้นบรรทัด');
    expect(ev.getCellValue('B4')).toBe(0.05);
  });
});
