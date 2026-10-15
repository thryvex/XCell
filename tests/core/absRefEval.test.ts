import { describe, it, expect } from 'vitest';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet } from '../../src/core/model';

describe('absolute single-cell references evaluate', () => {
  it('$A$1, A$1 and $A1 read the same cell as A1', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '7' };
    sheet.cells['B1'] = { raw: '=$A$1+A$1+$A1' };
    expect(new Evaluator(sheet).getCellValue('B1')).toBe(21);
  });
});
