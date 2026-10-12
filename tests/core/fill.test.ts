import { describe, it, expect } from 'vitest';
import { computeFill, fillDown, fillRight } from '../../src/core/fill';
import { createEmptySheet } from '../../src/core/model';

describe('src/core/fill', () => {
  it('fills formulas downwards with reference shifting', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '=B1*2' };

    const result = computeFill(
      sheet,
      { r1: 0, c1: 0, r2: 0, c2: 0 },
      { r1: 0, c1: 0, r2: 2, c2: 0 }
    );

    expect(result.cellsToSet['A2']?.raw).toBe('=B2*2');
    expect(result.cellsToSet['A3']?.raw).toBe('=B3*2');
  });

  it('fills number sequence (1, 2 -> 3, 4, 5)', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '1' };
    sheet.cells['A2'] = { raw: '2' };

    const result = computeFill(
      sheet,
      { r1: 0, c1: 0, r2: 1, c2: 0 },
      { r1: 0, c1: 0, r2: 4, c2: 0 }
    );

    expect(result.cellsToSet['A3']?.raw).toBe('3');
    expect(result.cellsToSet['A4']?.raw).toBe('4');
    expect(result.cellsToSet['A5']?.raw).toBe('5');
  });

  it('repeats single number or text', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '5' };
    sheet.cells['B1'] = { raw: 'รายการ' };

    const res1 = computeFill(
      sheet,
      { r1: 0, c1: 0, r2: 0, c2: 0 },
      { r1: 0, c1: 0, r2: 2, c2: 0 }
    );
    expect(res1.cellsToSet['A2']?.raw).toBe('5');
    expect(res1.cellsToSet['A3']?.raw).toBe('5');

    const res2 = computeFill(
      sheet,
      { r1: 0, c1: 1, r2: 0, c2: 1 },
      { r1: 0, c1: 1, r2: 2, c2: 1 }
    );
    expect(res2.cellsToSet['B2']?.raw).toBe('รายการ');
    expect(res2.cellsToSet['B3']?.raw).toBe('รายการ');
  });

  it('fills dates incrementing day by day', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '46296', fmt: 'date' }; // serial number

    const res = computeFill(
      sheet,
      { r1: 0, c1: 0, r2: 0, c2: 0 },
      { r1: 0, c1: 0, r2: 2, c2: 0 }
    );
    expect(res.cellsToSet['A2']?.raw).toBe('46297');
    expect(res.cellsToSet['A3']?.raw).toBe('46298');
  });

  it('performs Ctrl+D (fillDown) and Ctrl+R (fillRight)', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '=B1+10' };
    sheet.cells['B1'] = { raw: '100' };

    const downRes = fillDown(sheet, { r1: 0, c1: 0, r2: 2, c2: 0 });
    expect(downRes['A2']?.raw).toBe('=B2+10');
    expect(downRes['A3']?.raw).toBe('=B3+10');

    const rightRes = fillRight(sheet, { r1: 0, c1: 1, r2: 0, c2: 3 });
    expect(rightRes['C1']?.raw).toBe('100');
    expect(rightRes['D1']?.raw).toBe('100');
  });
});
