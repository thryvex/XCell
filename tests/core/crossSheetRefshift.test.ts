import { describe, it, expect } from 'vitest';
import {
  shiftFormula,
  shiftFormulaOnInsert,
  shiftFormulaOnDelete,
  renameSheetInFormula,
  invalidateSheetInFormula,
  normalizeSheetRefsInFormula,
} from '../../src/core/refshift';
import { formatSheetReference } from '../../src/core/sheetName';
import { insertRow, deleteRow, adjustOtherSheetsForRowCol, countFormulasBrokenByDelete } from '../../src/core/sheetOps';
import { duplicateSheet, renameSheet, deleteSheet } from '../../src/core/workbookOps';
import { Evaluator } from '../../src/core/evaluator';
import { Workbook, createEmptySheet } from '../../src/core/model';

const inX = { formulaSheet: 'X', targetSheet: 'X' };
const yToX = { formulaSheet: 'Y', targetSheet: 'X' };

describe('refshift ข้ามชีต (a) คัดลอก/วาง/ลากเติม', () => {
  it('a1. relative part moves, quoted sheet name stays', () => {
    expect(shiftFormula("='ชีต 2'!A1", 1, 0)).toBe("='ชีต 2'!A2");
  });

  it('a2. unquoted Thai sheet name stays, mixed refs both move', () => {
    expect(shiftFormula('=ชีต2!A1+B1', 1, 1)).toBe('=ชีต2!B2+C2');
  });

  it('a3. absolute cross-sheet ref does not move', () => {
    expect(shiftFormula("='ชีต 2'!$A$1*B2", 1, 0)).toBe("='ชีต 2'!$A$1*B3");
  });

  it('a4. cross-sheet range moves', () => {
    expect(shiftFormula("=SUM('ข้อมูล'!A1:A10)", 2, 0)).toBe("=SUM('ข้อมูล'!A3:A12)");
  });

  it('a5. mixed $ in cross-sheet range', () => {
    expect(shiftFormula("='ข้อมูล'!$A1:A$5", 1, 1)).toBe("='ข้อมูล'!$A2:B$5");
  });

  it('a6. moving out of bounds gives #REF!', () => {
    expect(shiftFormula("='ชีต 2'!A1", -1, 0)).toBe('=#REF!');
  });

  it("a7. sheet name with an apostrophe ('') is kept intact", () => {
    expect(shiftFormula("='Bob''s'!A1", 0, 1)).toBe("='Bob''s'!B1");
  });

  it('a8. sheet name spelling/case is preserved', () => {
    expect(shiftFormula('=data!a1', 1, 0)).toBe('=data!A2');
  });

  it('a9. SUMIF over another sheet shifts both ranges', () => {
    expect(shiftFormula(`=SUMIF('รายการ'!C2:C100,"อาหาร",'รายการ'!E2:E100)`, 0, 1)).toBe(
      `=SUMIF('รายการ'!D2:D100,"อาหาร",'รายการ'!F2:F100)`
    );
  });

  it('a10. unqualified ref stays unqualified (points to the destination sheet when pasted elsewhere)', () => {
    expect(shiftFormula("=A1+'ชีต 2'!A1", 1, 0)).toBe("=A2+'ชีต 2'!A2");
  });

  it('a11. text inside strings that looks like a reference is not touched', () => {
    expect(shiftFormula(`="ชีต2!A1"&ชีต2!A1`, 1, 0)).toBe(`="ชีต2!A1"&ชีต2!A2`);
  });
});

describe('refshift ข้ามชีต (b) แทรก/ลบแถว-คอลัมน์ในชีต X', () => {
  it('b1. insert row in X shifts X refs from another sheet', () => {
    expect(shiftFormulaOnInsert('=X!A5', 'row', 2, 1, yToX)).toBe('=X!A6');
  });

  it('b2. insert row in X does not touch unqualified refs of a formula in Y', () => {
    expect(shiftFormulaOnInsert('=A5', 'row', 2, 1, yToX)).toBe('=A5');
  });

  it('b3. insert row in X does NOT touch refs to Y that live in X', () => {
    expect(shiftFormulaOnInsert('=Y!A5', 'row', 2, 1, inX)).toBe('=Y!A5');
  });

  it('b4. formula in X: own refs shift, refs to Y stay', () => {
    expect(shiftFormulaOnInsert('=A5+Y!A5', 'row', 2, 1, inX)).toBe('=A6+Y!A5');
  });

  it('b5. formula in X referencing X by name shifts', () => {
    expect(shiftFormulaOnInsert('=X!A5+A5', 'row', 0, 1, inX)).toBe('=X!A6+A6');
  });

  it("b6. range 'ชีต 2'!A1:B5 spanning the insert point expands", () => {
    const scope = { formulaSheet: 'ชีต1', targetSheet: 'ชีต 2' };
    expect(shiftFormulaOnInsert("=SUM('ชีต 2'!A1:B5)", 'row', 2, 1, scope)).toBe("=SUM('ชีต 2'!A1:B6)");
  });

  it('b7. range entirely below the insert point moves down', () => {
    const scope = { formulaSheet: 'ชีต1', targetSheet: 'ชีต 2' };
    expect(shiftFormulaOnInsert("=SUM('ชีต 2'!A3:A4)", 'row', 1, 1, scope)).toBe("=SUM('ชีต 2'!A4:A5)");
  });

  it('b8. range entirely above the insert point stays', () => {
    const scope = { formulaSheet: 'ชีต1', targetSheet: 'ชีต 2' };
    expect(shiftFormulaOnInsert("=SUM('ชีต 2'!A1:A2)", 'row', 5, 1, scope)).toBe("=SUM('ชีต 2'!A1:A2)");
  });

  it('b9. absolute refs also move on insert (like Google Sheets)', () => {
    const scope = { formulaSheet: 'ชีต1', targetSheet: 'ชีต 2' };
    expect(shiftFormulaOnInsert("='ชีต 2'!$A$5", 'row', 0, 1, scope)).toBe("='ชีต 2'!$A$6");
  });

  it('b10. insert column in X shifts X refs', () => {
    expect(shiftFormulaOnInsert('=X!C1', 'col', 1, 1, yToX)).toBe('=X!D1');
  });

  it('b11. sheet names match case-insensitively', () => {
    expect(shiftFormulaOnInsert('=data!A5', 'row', 0, 1, { formulaSheet: 'Y', targetSheet: 'Data' })).toBe(
      '=data!A6'
    );
  });

  it('b12. delete referenced row in X gives #REF!', () => {
    expect(shiftFormulaOnDelete('=X!A5', 'row', 4, 1, yToX)).toBe('=#REF!');
  });

  it('b13. delete row above shifts X ref up', () => {
    expect(shiftFormulaOnDelete('=X!A5', 'row', 1, 1, yToX)).toBe('=X!A4');
  });

  it('b14. delete row in X does NOT touch refs to Y living in X', () => {
    expect(shiftFormulaOnDelete('=Y!A5+A5', 'row', 4, 1, inX)).toBe('=Y!A5+#REF!');
  });

  it('b15. delete rows cutting the top of a cross-sheet range contracts it', () => {
    const scope = { formulaSheet: 'ชีต1', targetSheet: 'ชีต 2' };
    expect(shiftFormulaOnDelete("=SUM('ชีต 2'!A2:A6)", 'row', 1, 2, scope)).toBe("=SUM('ชีต 2'!A2:A4)");
  });

  it('b16. deleting a whole cross-sheet range gives #REF!', () => {
    expect(shiftFormulaOnDelete('=SUM(X!A2:A3)', 'row', 1, 2, yToX)).toBe('=SUM(#REF!)');
  });

  it('b17. delete column in X', () => {
    expect(shiftFormulaOnDelete('=X!B1', 'col', 1, 1, yToX)).toBe('=#REF!');
    expect(shiftFormulaOnDelete('=X!C1+$X$1', 'col', 1, 1, yToX)).toBe('=X!B1+$X$1');
  });

  it('b18. without a scope all refs are adjusted (single-sheet behavior unchanged)', () => {
    expect(shiftFormulaOnInsert('=A5', 'row', 0, 1)).toBe('=A6');
  });

  it('b19. sheetOps.insertRow on X leaves refs to Y untouched', () => {
    const x = createEmptySheet('X');
    x.cells['A1'] = { raw: '=Y!A5+A5' };
    insertRow(x, 0, 1);
    expect(x.cells['A2'].raw).toBe('=Y!A5+A6');
  });

  it('b20. sheetOps.deleteRow on X leaves refs to Y untouched', () => {
    const x = createEmptySheet('X');
    x.cells['B1'] = { raw: '=Y!A3+A3' };
    deleteRow(x, 1, 1);
    expect(x.cells['B1'].raw).toBe('=Y!A3+A2');
  });

  it('b21. adjustOtherSheetsForRowCol rewrites only formulas pointing at X and returns patches', () => {
    const wb: Workbook = {
      sheets: [createEmptySheet('X'), createEmptySheet('Y'), createEmptySheet('Z')],
      activeSheet: 0,
    };
    wb.sheets[1].cells['A1'] = { raw: '=X!A5+A5' };
    wb.sheets[2].cells['A1'] = { raw: '=Y!A5' };
    const patches = adjustOtherSheetsForRowCol(wb, 0, 'insert', 'row', 0, 1);
    expect(wb.sheets[1].cells['A1'].raw).toBe('=X!A6+A5');
    expect(wb.sheets[2].cells['A1'].raw).toBe('=Y!A5');
    expect(patches).toHaveLength(1);
    expect(patches[0]).toMatchObject({ sheetIndex: 1, key: 'A1', before: { raw: '=X!A5+A5' } });
  });

  it('b22. delete warning counts formulas in other sheets too', () => {
    const wb: Workbook = { sheets: [createEmptySheet('X'), createEmptySheet('Y')], activeSheet: 0 };
    wb.sheets[0].cells['B1'] = { raw: '=A3' };
    wb.sheets[1].cells['A1'] = { raw: '=X!A3' };
    wb.sheets[1].cells['A2'] = { raw: '=A3' };
    expect(countFormulasBrokenByDelete(wb, 0, 'row', 2, 1)).toBe(2);
  });
});

describe('refshift ข้ามชีต (c) เปลี่ยนชื่อชีต', () => {
  it('c1. Thai name gets quoted', () => {
    expect(renameSheetInFormula('=ชีต2!A1', 'ชีต2', 'ยอดขาย')).toBe("='ยอดขาย'!A1");
  });

  it('c2. quotes are removed when the new name is plain', () => {
    expect(renameSheetInFormula("='My Sheet'!A1", 'My Sheet', 'Data')).toBe('=Data!A1');
  });

  it('c3. quotes are added for names with spaces', () => {
    expect(renameSheetInFormula('=Data!A1', 'Data', 'My Data')).toBe("='My Data'!A1");
  });

  it("c4. apostrophe in the new name becomes ''", () => {
    expect(renameSheetInFormula('=Data!A1', 'Data', "Bob's")).toBe("='Bob''s'!A1");
  });

  it('c5. old name matched case-insensitively', () => {
    expect(renameSheetInFormula('=DATA!A1+data!B2', 'Data', 'New')).toBe('=New!A1+New!B2');
  });

  it('c6. other sheets and unqualified refs untouched', () => {
    expect(renameSheetInFormula('=A1+Other!A1', 'Data', 'New')).toBe('=A1+Other!A1');
  });

  it('c7. $ and ranges preserved', () => {
    expect(renameSheetInFormula('=SUM(Data!$A$1:B5)', 'Data', 'New')).toBe('=SUM(New!$A$1:B5)');
  });

  it('c8. names that look like cells or start with a digit are quoted', () => {
    expect(renameSheetInFormula('=Data!B2', 'Data', 'A1')).toBe("='A1'!B2");
    expect(renameSheetInFormula('=Data!B2', 'Data', '2024')).toBe("='2024'!B2");
  });

  it('c9. string literals are not rewritten', () => {
    expect(renameSheetInFormula('="Data!A1"&Data!A1', 'Data', 'New')).toBe('="Data!A1"&New!A1');
  });

  it('c10. renameSheet rewrites formulas in every sheet including the renamed one', () => {
    const wb: Workbook = { sheets: [createEmptySheet('ชีต1'), createEmptySheet('ชีต2')], activeSheet: 0 };
    wb.sheets[0].cells['A1'] = { raw: '=ชีต2!A1*2' };
    wb.sheets[1].cells['A2'] = { raw: '=ชีต2!A1+A1' };
    expect(renameSheet(wb, 1, 'ข้อมูล ปี 2569').ok).toBe(true);
    expect(wb.sheets[0].cells['A1'].raw).toBe("='ข้อมูล ปี 2569'!A1*2");
    expect(wb.sheets[1].cells['A2'].raw).toBe("='ข้อมูล ปี 2569'!A1+A1");
  });
});

describe('refshift ข้ามชีต (d) ลบชีต', () => {
  it('d1. refs to the deleted sheet become #REF!', () => {
    expect(invalidateSheetInFormula('=Data!A1+B1', 'Data')).toBe('=#REF!+B1');
  });

  it('d2. ranges become #REF! (case-insensitive)', () => {
    expect(invalidateSheetInFormula('=SUM(data!A1:A3)', 'Data')).toBe('=SUM(#REF!)');
  });

  it('d3. refs to other sheets are untouched', () => {
    expect(invalidateSheetInFormula("='ชีต 3'!A1+Data!A1", 'Data')).toBe("='ชีต 3'!A1+#REF!");
  });

  it('d4. deleteSheet rewrites other sheets and the result evaluates to #REF!', () => {
    const wb: Workbook = { sheets: [createEmptySheet('ชีต1'), createEmptySheet('ชีต2')], activeSheet: 0 };
    wb.sheets[0].cells['A1'] = { raw: '=ชีต2!A1+1' };
    expect(deleteSheet(wb, 1).ok).toBe(true);
    expect(wb.sheets[0].cells['A1'].raw).toBe('=#REF!+1');
    expect(new Evaluator(wb).getCellValue('A1', 0)).toBe('#REF!');
  });
});

describe('refshift ข้ามชีต (e) ทำสำเนาชีต', () => {
  it('e1. unqualified refs in the copy point to the copy, qualified refs keep pointing to the original', () => {
    const wb: Workbook = { sheets: [createEmptySheet('ชีต1'), createEmptySheet('ข้อมูล')], activeSheet: 0 };
    wb.sheets[0].cells['A1'] = { raw: '5' };
    wb.sheets[0].cells['B1'] = { raw: '=A1*10' };
    wb.sheets[0].cells['C1'] = { raw: '=ข้อมูล!A1' };
    wb.sheets[0].cells['D1'] = { raw: '=ชีต1!A1' };
    wb.sheets[1].cells['A1'] = { raw: '7' };

    const res = duplicateSheet(wb, 0);
    expect(res.index).toBe(1);
    const copy = wb.sheets[1];
    expect(copy.name).toBe('สำเนาของ ชีต1');
    expect(copy.cells['B1'].raw).toBe('=A1*10');
    expect(copy.cells['C1'].raw).toBe('=ข้อมูล!A1');

    copy.cells['A1'] = { raw: '6' };
    const ev = new Evaluator(wb);
    expect(ev.getCellValue('B1', 1)).toBe(60); // ชี้สำเนาเอง
    expect(ev.getCellValue('B1', 0)).toBe(50); // ต้นฉบับไม่เปลี่ยน
    expect(ev.getCellValue('C1', 1)).toBe(7);
    expect(ev.getCellValue('D1', 1)).toBe(5); // อ้างชื่อชีตต้นฉบับ → ยังชี้ต้นฉบับ
  });
});

describe('รูปแบบชื่อชีตในสูตร', () => {
  it('quotes everything except plain A-Z / digits / underscore names', () => {
    expect(formatSheetReference('Sheet2')).toBe('Sheet2');
    expect(formatSheetReference('my_data')).toBe('my_data');
    expect(formatSheetReference('ชีต2')).toBe("'ชีต2'");
    expect(formatSheetReference('My Sheet')).toBe("'My Sheet'");
    expect(formatSheetReference("Bob's")).toBe("'Bob''s'");
    expect(formatSheetReference('2024')).toBe("'2024'");
    expect(formatSheetReference('AB12')).toBe("'AB12'");
    expect(formatSheetReference('TRUE')).toBe("'TRUE'");
  });

  it('normalizeSheetRefsInFormula quotes Thai names for .xlsx', () => {
    expect(normalizeSheetRefsInFormula('=ชีต2!A1+Data!B2+A1')).toBe("='ชีต2'!A1+Data!B2+A1");
  });
});
