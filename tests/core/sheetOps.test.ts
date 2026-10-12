import { describe, it, expect } from 'vitest';
import { insertRow, deleteRow, insertCol, deleteCol } from '../../src/core/sheetOps';
import { createEmptySheet } from '../../src/core/model';

describe('src/core/sheetOps', () => {
  it('inserts row and shifts cells, formulas, and merges', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['A2'] = { raw: '20' };
    sheet.cells['A3'] = { raw: '=A1+A2' };
    sheet.merges = [{ r1: 0, c1: 0, r2: 1, c2: 1 }]; // A1:B2

    // Insert 1 row at row 2 (0-based index 1)
    insertRow(sheet, 1, 1);

    expect(sheet.cells['A1'].raw).toBe('10');
    expect(sheet.cells['A2']).toBeUndefined(); // newly inserted blank row
    expect(sheet.cells['A3'].raw).toBe('20'); // shifted down from A2
    expect(sheet.cells['A4'].raw).toBe('=A1+A3'); // shifted formula from A3

    // Merge A1:B2 expanded to A1:B3
    expect(sheet.merges[0]).toEqual({ r1: 0, c1: 0, r2: 2, c2: 1 });
  });

  it('deletes row and updates formulas to #REF!', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['A2'] = { raw: '20' };
    sheet.cells['A3'] = { raw: '=A1+A2' };

    deleteRow(sheet, 1, 1); // delete row 2 (0-based 1)

    expect(sheet.cells['A1'].raw).toBe('10');
    expect(sheet.cells['A2'].raw).toBe('=A1+#REF!'); // shifted up, A2 ref became #REF!
  });

  it('inserts and deletes columns', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['B1'] = { raw: '=A1*2' };

    insertCol(sheet, 1, 1);
    expect(sheet.cells['C1'].raw).toBe('=A1*2');

    deleteCol(sheet, 1, 1);
    expect(sheet.cells['B1'].raw).toBe('=A1*2');
  });
});
