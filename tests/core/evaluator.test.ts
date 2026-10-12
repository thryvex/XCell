import { describe, it, expect } from 'vitest';
import { Evaluator, normalizeFloat } from '../../src/core/evaluator';
import { createEmptySheet } from '../../src/core/model';

describe('src/core/evaluator', () => {
  it('normalizes float precision (0.1 + 0.2 = 0.3)', () => {
    expect(normalizeFloat(0.1 + 0.2)).toBe(0.3);
  });

  it('evaluates basic arithmetic and operator precedence', () => {
    const sheet = createEmptySheet();
    const evaluator = new Evaluator(sheet);

    expect(evaluator.evaluateFormula('=1 + 2 * 3')).toBe(7);
    expect(evaluator.evaluateFormula('=(1 + 2) * 3')).toBe(9);
    expect(evaluator.evaluateFormula('=2 ^ 3')).toBe(8);
    expect(evaluator.evaluateFormula('=10%')).toBe(0.1);
    expect(evaluator.evaluateFormula('=50 * 10%')).toBe(5);
    expect(evaluator.evaluateFormula('="Hello " & "World"')).toBe('Hello World');
  });

  it('evaluates cell references and ranges', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['A2'] = { raw: '20' };
    sheet.cells['A3'] = { raw: '=SUM(A1:A2)' };
    sheet.cells['A4'] = { raw: '=A3 * 2' };

    const evaluator = new Evaluator(sheet);
    expect(evaluator.getCellValue('A3')).toBe(30);
    expect(evaluator.getCellValue('A4')).toBe(60);
  });

  it('handles division by zero (#DIV/0!) and invalid values (#VALUE!)', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '=10 / 0' };
    sheet.cells['A2'] = { raw: '="text" * 5' };

    const evaluator = new Evaluator(sheet);
    expect(evaluator.getCellValue('A1')).toBe('#DIV/0!');
    expect(evaluator.getCellValue('A2')).toBe('#VALUE!');
  });

  it('detects circular references and returns #CYCLE!', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '=B1 + 1' };
    sheet.cells['B1'] = { raw: '=A1 + 1' };

    const evaluator = new Evaluator(sheet);
    expect(evaluator.getCellValue('A1')).toBe('#CYCLE!');
    expect(evaluator.getCellValue('B1')).toBe('#CYCLE!');
  });

  it('invalidates cache when cell changes', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['A2'] = { raw: '=A1 * 2' };

    const evaluator = new Evaluator(sheet);
    expect(evaluator.getCellValue('A2')).toBe(20);

    // Update A1
    sheet.cells['A1'] = { raw: '50' };
    evaluator.invalidate('A1');
    expect(evaluator.getCellValue('A2')).toBe(100);
  });
});
