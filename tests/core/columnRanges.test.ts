import { describe, it, expect } from 'vitest';
import ExcelJS from 'exceljs';
import { tokenize } from '../../src/core/tokenizer';
import { parseRange, MAX_SHEET_ROWS } from '../../src/core/address';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet, Workbook } from '../../src/core/model';
import {
  shiftFormula,
  shiftFormulaOnInsert,
  shiftFormulaOnDelete,
  renameSheetInFormula,
  moveRefsInFormula,
} from '../../src/core/refshift';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';

const refs = (f: string) =>
  tokenize(f)
    .filter((t) => t.type === 'RANGE' || t.type === 'CELL_REF')
    .map((t) => `${t.type}:${t.sheet ? t.sheet + '!' : ''}${t.value}`);

describe('อ้างทั้งคอลัมน์ A:A (SPEC 4.1): tokenizer และ parseRange', () => {
  it('อ่าน A:A, A:C, $A:$C, ตัวพิมพ์เล็ก และแบบมีชื่อชีตเป็นช่วง', () => {
    expect(refs('=SUM(G:G)')).toEqual(['RANGE:G:G']);
    expect(refs('=sum(g:g)')).toEqual(['RANGE:G:G']);
    expect(refs('=SUM(A:C)+$B:$B')).toEqual(['RANGE:A:C', 'RANGE:$B:$B']);
    expect(refs("=SUM('ข้อมูล ปี 69'!A:A)")).toEqual(['RANGE:ข้อมูล ปี 69!A:A']);
    expect(refs('=SUM(Sheet2!B:D)')).toEqual(['RANGE:Sheet2!B:D']);
    expect(refs('=SUM(A1:B2)+LEN(C1)')).toEqual(['RANGE:A1:B2', 'CELL_REF:C1']);
  });

  it('parseRange: ทั้งคอลัมน์ แถว 1 ถึงแถวสุดท้ายของ .xlsx พร้อม $', () => {
    expect(parseRange('$A:C')).toEqual({
      r1: 0,
      c1: 0,
      r2: MAX_SHEET_ROWS - 1,
      c2: 2,
      startColAbs: true,
      endColAbs: false,
      wholeColumns: true,
    });
    expect(parseRange('C:A')).toMatchObject({ c1: 0, c2: 2, wholeColumns: true });
    expect(parseRange('A1:B2')?.wholeColumns).toBeUndefined();
  });
});

function book(cells: Record<string, string>, other?: Record<string, string>): Workbook {
  const s1 = createEmptySheet('ชีต1');
  for (const [k, v] of Object.entries(cells)) s1.cells[k] = { raw: v };
  const sheets = [s1];
  if (other) {
    const s2 = createEmptySheet('ข้อมูล ปี 69');
    for (const [k, v] of Object.entries(other)) s2.cells[k] = { raw: v };
    sheets.push(s2);
  }
  return { sheets, activeSheet: 0 };
}

describe('อ้างทั้งคอลัมน์: evaluator', () => {
  it('SUM / AVERAGE / COUNT / MAX ทั้งคอลัมน์ ข้ามหัวตารางที่เป็นข้อความ', () => {
    const ev = new Evaluator(book({ G1: 'รวม', G2: '10', G3: '20', G7: '30', J1: '=SUM(G:G)', J2: '=AVERAGE(G:G)', J3: '=COUNT(G:G)', J4: '=MAX(G:H)' }));
    expect(ev.getCellValue('J1')).toBe(60);
    expect(ev.getCellValue('J2')).toBe(20);
    expect(ev.getCellValue('J3')).toBe(3);
    expect(ev.getCellValue('J4')).toBe(30);
  });

  it('รวมแถวที่อยู่ไกลมาก และไม่ไล่ทุก 1,048,576 แถว', () => {
    const wb = book({ A1: '1', A500000: '2', B1: '=SUM(A:A)' });
    const ev = new Evaluator(wb);
    const t0 = Date.now();
    expect(ev.getCellValue('B1')).toBe(3);
    expect(Date.now() - t0).toBeLessThan(200);
  });

  it('เพิ่ม/แก้เซลในคอลัมน์ทีหลัง ผลรวมอัปเดต (dependency ต่อคอลัมน์)', () => {
    const wb = book({ G2: '5', J1: '=SUM(G:G)', K1: '=J1*2' });
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('K1')).toBe(10);
    wb.sheets[0].cells['G40'] = { raw: '7' };
    ev.invalidate('G40', 0);
    expect(ev.getCellValue('J1')).toBe(12);
    expect(ev.getCellValue('K1')).toBe(24);
    wb.sheets[0].cells['G2'] = { raw: '1' };
    ev.invalidate('G2', 0);
    expect(ev.getCellValue('K1')).toBe(16);
    // เซลคอลัมน์อื่นไม่เกี่ยว
    wb.sheets[0].cells['H3'] = { raw: '100' };
    ev.invalidate('H3', 0);
    expect(ev.getCellValue('J1')).toBe(8);
  });

  it("อ้างข้ามชีต 'ชื่อชีต'!A:A และอัปเดตเมื่อชีตนั้นเปลี่ยน", () => {
    const wb = book({ A1: "=SUM('ข้อมูล ปี 69'!B:B)" }, { B1: '4', B9: '6' });
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('A1', 0)).toBe(10);
    wb.sheets[1].cells['B20'] = { raw: '5' };
    ev.invalidate('B20', 1);
    expect(ev.getCellValue('A1', 0)).toBe(15);
  });

  it('สูตรนอกคอลัมน์ที่อ้างทั้งคอลัมน์ไม่เป็นวงกลม; สูตรในคอลัมน์เดียวกันเป็นวงกลม (#CYCLE!)', () => {
    const ok = new Evaluator(book({ G1: '1', G2: '=G1*2', J1: '=SUM(G:G)' }));
    expect(ok.getCellValue('J1')).toBe(3);
    const loop = new Evaluator(book({ G1: '1', G2: '=SUM(G:G)', J1: '=SUM(G:G)' }));
    expect(loop.getCellValue('G2')).toBe('#CYCLE!');
    // J1 รวม G2 ที่เป็นวงกลม จึงได้ error ตามไปด้วย (เหมือนสูตรอ้างเซลที่ error)
    expect(loop.getCellValue('J1')).toBe('#CYCLE!');
  });

  it('SUMIF / COUNTIF จับคู่ช่วงทั้งคอลัมน์ตรงแถวกัน แม้ความยาวข้อมูลแต่ละคอลัมน์ต่างกัน', () => {
    const ev = new Evaluator(
      book({ A1: 'หมวด', A2: 'อาหาร', A3: 'เดินทาง', A4: 'อาหาร', A9: 'อาหาร', B2: '100', B3: '50', B4: '30', D1: '=SUMIF(A:A,"อาหาร",B:B)', D2: '=COUNTIF(A:A,"อาหาร")' })
    );
    expect(ev.getCellValue('D1')).toBe(130);
    expect(ev.getCellValue('D2')).toBe(3);
  });
});

describe('อ้างทั้งคอลัมน์: refshift', () => {
  it('คัดลอก/ลาก: เลื่อนคอลัมน์ตาม ไม่เลื่อนแถว, $ คงที่', () => {
    expect(shiftFormula('=SUM(G:G)', 5, 1)).toBe('=SUM(H:H)');
    expect(shiftFormula('=SUM(G:G)', 5, 0)).toBe('=SUM(G:G)');
    expect(shiftFormula('=SUM($G:$G)+SUM(A:$B)', 3, 2)).toBe('=SUM($G:$G)+SUM(C:$B)');
  });

  it('แทรก/ลบคอลัมน์ ปรับตาม; แทรก/ลบแถวไม่กระทบ', () => {
    expect(shiftFormulaOnInsert('=SUM(G:G)', 'col', 3, 1)).toBe('=SUM(H:H)');
    expect(shiftFormulaOnInsert('=SUM(A:C)', 'col', 1, 2)).toBe('=SUM(A:E)');
    expect(shiftFormulaOnInsert('=SUM(G:G)', 'row', 0, 5)).toBe('=SUM(G:G)');
    expect(shiftFormulaOnDelete('=SUM(G:G)', 'col', 6, 1)).toBe('=SUM(#REF!)');
    expect(shiftFormulaOnDelete('=SUM(A:E)', 'col', 1, 2)).toBe('=SUM(A:C)');
    expect(shiftFormulaOnDelete('=SUM(G:G)', 'row', 0, 3)).toBe('=SUM(G:G)');
  });

  it('เปลี่ยนชื่อชีตและตัด-วางไม่ทำให้รูปแบบ A:A เสีย', () => {
    expect(renameSheetInFormula("=SUM('ชีต1'!A:A)", 'ชีต1', 'สรุป ปี')).toBe("=SUM('สรุป ปี'!A:A)");
    const move = { sourceSheet: 'ชีต1', destSheet: 'ชีต1', sourceRows: [1], c1: 0, c2: 0, destRows: [5], dc: 0 };
    expect(moveRefsInFormula('=SUM(A:A)+A2', 'ชีต1', 'ชีต1', move)).toBe('=SUM(A:A)+A6');
  });
});

describe('อ้างทั้งคอลัมน์: .xlsx', () => {
  it('บันทึกแล้วเปิดกลับได้สูตรเดิม และคำนวณได้', async () => {
    const wb = book({ G2: '3', G3: '4', J1: '=SUM(G:G)', J2: "=SUM('ข้อมูล ปี 69'!A:A)" }, { A1: '10' });
    const back = await importFromXlsxBuffer(await exportToXlsxBuffer(wb));
    expect(back.sheets[0].cells['J1'].raw).toBe('=SUM(G:G)');
    expect(back.sheets[0].cells['J2'].raw).toBe("=SUM('ข้อมูล ปี 69'!A:A)");
    expect(back.sheets[0].cells['J1'].unsupportedFormula).toBeUndefined();
    const ev = new Evaluator(back);
    expect(ev.getCellValue('J1')).toBe(7);
    expect(ev.getCellValue('J2')).toBe(10);
  });

  it('เปิดไฟล์จากโปรแกรมอื่นที่เขียน =sum(G:G) ตัวพิมพ์เล็ก: ไม่ถูกมองว่าเป็นฟังก์ชันที่ไม่รองรับ', async () => {
    const x = new ExcelJS.Workbook();
    const ws = x.addWorksheet('Sheet1');
    ws.getCell('G2').value = 40;
    ws.getCell('G3').value = 2;
    ws.getCell('J1').value = { formula: 'sum(G:G)', result: 42 } as ExcelJS.CellFormulaValue;
    const wb = await importFromXlsxBuffer(new Uint8Array(await x.xlsx.writeBuffer()));
    const j1 = wb.sheets[0].cells['J1'];
    expect(j1.raw.toUpperCase()).toBe('=SUM(G:G)');
    expect(j1.unsupportedFormula).toBeUndefined();
    expect(new Evaluator(wb).getCellValue('J1')).toBe(42);
  });
});
