import { describe, it, expect } from 'vitest';
import { tokenize } from '../../src/core/tokenizer';
import { Evaluator } from '../../src/core/evaluator';
import { Workbook, createEmptySheet } from '../../src/core/model';
import { convertShorthand } from '../../src/core/shorthand';

function makeWorkbook(...names: string[]): Workbook {
  return { sheets: names.map((n) => createEmptySheet(n)), activeSheet: 0 };
}

describe('tokenizer: อ้างข้ามชีต', () => {
  it('quoted sheet name with spaces and Thai', () => {
    const [t] = tokenize("='ชีต 2'!A1");
    expect(t).toMatchObject({ type: 'CELL_REF', value: 'A1', sheet: 'ชีต 2', raw: "'ชีต 2'!A1" });
  });

  it("quoted name with '' unescapes to '", () => {
    const [t] = tokenize("='Bob''s'!$B$2");
    expect(t).toMatchObject({ type: 'CELL_REF', value: '$B$2', sheet: "Bob's" });
  });

  it('unquoted Thai / ASCII / digit names and ranges', () => {
    expect(tokenize('=ชีต2!a1:b5')[0]).toMatchObject({ type: 'RANGE', value: 'A1:B5', sheet: 'ชีต2' });
    expect(tokenize('=Data_1!C3')[0]).toMatchObject({ type: 'CELL_REF', sheet: 'Data_1' });
    expect(tokenize('=2024!C3')[0]).toMatchObject({ type: 'CELL_REF', sheet: '2024' });
  });

  it('unquoted names with spaces are not sheet refs', () => {
    const tokens = tokenize('=My Sheet!A1');
    expect(tokens.some((t) => t.sheet !== undefined && t.sheet.includes(' '))).toBe(false);
  });

  it('strings containing ! stay strings, and #REF! is an error literal', () => {
    const tokens = tokenize('="Data!A1"&#REF!');
    expect(tokens[0]).toMatchObject({ type: 'STRING', value: 'Data!A1' });
    expect(tokens[2]).toMatchObject({ type: 'ERROR', value: '#REF!' });
  });

  it('shorthand does not auto-convert text with sheet references', () => {
    expect(convertShorthand('ชีต2!a1+1')).toBeNull();
    expect(convertShorthand('a1+1')).toBe('=A1+1');
  });
});

describe('Evaluator: คำนวณข้ามชีต', () => {
  it('reads values from another sheet (quoted / unquoted / $)', () => {
    const wb = makeWorkbook('ชีต1', 'ชีต 2', 'Data');
    wb.sheets[1].cells['A1'] = { raw: '40' };
    wb.sheets[2].cells['B2'] = { raw: '2' };
    wb.sheets[0].cells['A1'] = { raw: "='ชีต 2'!A1+Data!$B$2" };
    wb.sheets[0].cells['A2'] = { raw: "=data!B2*'ชีต 2'!$A$1" };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe(42);
    expect(ev.getCellValue('A2', 0)).toBe(80);
  });

  it('ranges on another sheet work with SUM / AVERAGE / COUNT', () => {
    const wb = makeWorkbook('สรุป', 'ข้อมูล');
    wb.sheets[1].cells['A1'] = { raw: '10' };
    wb.sheets[1].cells['A2'] = { raw: '20' };
    wb.sheets[1].cells['A3'] = { raw: 'ข้อความ' };
    wb.sheets[0].cells['A1'] = { raw: "=SUM('ข้อมูล'!A1:A3)" };
    wb.sheets[0].cells['A2'] = { raw: '=AVERAGE(ข้อมูล!A1:A3)' };
    wb.sheets[0].cells['A3'] = { raw: '=COUNTA(ข้อมูล!A1:A3)' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe(30);
    expect(ev.getCellValue('A2', 0)).toBe(15);
    expect(ev.getCellValue('A3', 0)).toBe(3);
  });

  it('conditional functions over another sheet (SUMIF, COUNTIF, SUMIFS, AVERAGEIF)', () => {
    const wb = makeWorkbook('สรุป', 'รายการ');
    const items = wb.sheets[1].cells;
    items['A1'] = { raw: 'อาหาร' };
    items['B1'] = { raw: '100' };
    items['A2'] = { raw: 'เดินทาง' };
    items['B2'] = { raw: '50' };
    items['A3'] = { raw: 'อาหาร' };
    items['B3'] = { raw: '30' };
    const s = wb.sheets[0].cells;
    s['A1'] = { raw: `=SUMIF(รายการ!A1:A3,"อาหาร",รายการ!B1:B3)` };
    s['A2'] = { raw: `=COUNTIF('รายการ'!A1:A3,"อาหาร")` };
    s['A3'] = { raw: `=SUMIFS('รายการ'!B1:B3,'รายการ'!A1:A3,"อาหาร",'รายการ'!B1:B3,">50")` };
    s['A4'] = { raw: `=AVERAGEIF(รายการ!A1:A3,"อาหาร",รายการ!B1:B3)` };
    s['B1'] = { raw: 'อาหาร' };
    s['A5'] = { raw: `=SUMIF(รายการ!A1:A3,B1,รายการ!B1:B3)` };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe(130);
    expect(ev.getCellValue('A2', 0)).toBe(2);
    expect(ev.getCellValue('A3', 0)).toBe(100);
    expect(ev.getCellValue('A4', 0)).toBe(65);
    expect(ev.getCellValue('A5', 0)).toBe(130);
  });

  it('circular reference across sheets gives #CYCLE!', () => {
    const wb = makeWorkbook('ชีต1', 'ชีต2');
    wb.sheets[0].cells['A1'] = { raw: '=ชีต2!A1+1' };
    wb.sheets[1].cells['A1'] = { raw: '=ชีต1!A1+1' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe('#CYCLE!');
    expect(ev.getCellValue('A1', 1)).toBe('#CYCLE!');
  });

  it('same address on different sheets is not a cycle', () => {
    const wb = makeWorkbook('ชีต1', 'ชีต2');
    wb.sheets[1].cells['A1'] = { raw: '3' };
    wb.sheets[0].cells['A1'] = { raw: '=ชีต2!A1*2' };
    expect(new Evaluator(wb).getCellValue('A1', 0)).toBe(6);
  });

  it('missing sheet name gives #REF!', () => {
    const wb = makeWorkbook('ชีต1');
    wb.sheets[0].cells['A1'] = { raw: "='ไม่มีชีตนี้'!A1" };
    wb.sheets[0].cells['A2'] = { raw: '=SUM(ไม่มี!A1:A5)' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe('#REF!');
    expect(ev.getCellValue('A2', 0)).toBe('#REF!');
  });

  it('#REF! literal inside a formula evaluates to #REF!', () => {
    const wb = makeWorkbook('ชีต1');
    wb.sheets[0].cells['A1'] = { raw: '=1+#REF!' };
    expect(new Evaluator(wb).getCellValue('A1', 0)).toBe('#REF!');
  });

  it('editing a cell recalculates dependents on other sheets immediately', () => {
    const wb = makeWorkbook('ชีต1', 'ชีต2', 'ชีต3');
    wb.sheets[1].cells['A1'] = { raw: '1' };
    wb.sheets[0].cells['A1'] = { raw: '=ชีต2!A1*100' };
    wb.sheets[2].cells['A1'] = { raw: '=ชีต1!A1+1' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 2)).toBe(101);

    wb.sheets[1].cells['A1'] = { raw: '2' };
    ev.invalidate('A1', 1);
    expect(ev.getCellValue('A1', 0)).toBe(200);
    expect(ev.getCellValue('A1', 2)).toBe(201);
  });

  it('a range dependency on another sheet triggers recalculation', () => {
    const wb = makeWorkbook('ชีต1', 'ชีต2');
    wb.sheets[1].cells['A2'] = { raw: '5' };
    wb.sheets[0].cells['A1'] = { raw: '=SUM(ชีต2!A1:A3)' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe(5);
    wb.sheets[1].cells['A3'] = { raw: '7' };
    ev.invalidate('A3', 1);
    expect(ev.getCellValue('A1', 0)).toBe(12);
  });

  it('hidden sheets still compute', () => {
    const wb = makeWorkbook('ชีต1', 'ซ่อน');
    wb.sheets[1].hidden = 'hidden';
    wb.sheets[1].cells['A1'] = { raw: '9' };
    wb.sheets[0].cells['A1'] = { raw: '=ซ่อน!A1' };
    expect(new Evaluator(wb).getCellValue('A1', 0)).toBe(9);
  });

  it('a sheet created later resolves after setWorkbook', () => {
    const wb = makeWorkbook('ชีต1');
    wb.sheets[0].cells['A1'] = { raw: '=ใหม่!A1' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe('#REF!');
    wb.sheets.push(createEmptySheet('ใหม่'));
    wb.sheets[1].cells['A1'] = { raw: '4' };
    ev.setWorkbook(wb);
    expect(ev.getCellValue('A1', 0)).toBe(4);
  });
});
