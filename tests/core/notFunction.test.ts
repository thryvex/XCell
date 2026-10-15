import { describe, it, expect } from 'vitest';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet } from '../../src/core/model';

describe('NOT', () => {
  it('treats non-zero numbers as TRUE and zero as FALSE', () => {
    const ev = new Evaluator(createEmptySheet());
    expect(ev.evaluateFormula('=NOT(5)')).toBe(false);
    expect(ev.evaluateFormula('=NOT(-1)')).toBe(false);
    expect(ev.evaluateFormula('=NOT(0)')).toBe(true);
    expect(ev.evaluateFormula('=NOT(TRUE)')).toBe(false);
    expect(ev.evaluateFormula('=NOT(FALSE)')).toBe(true);
  });
});
