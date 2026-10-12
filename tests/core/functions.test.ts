import { describe, it, expect } from 'vitest';
import { FUNCTIONS } from '../../src/core/functions';

describe('src/core/functions', () => {
  it('SUM adds numbers and ignores text/blanks in ranges', () => {
    expect(FUNCTIONS.SUM(1, 2, 3)).toBe(6);
    expect(FUNCTIONS.SUM([1, 'hello', '', 2, 3])).toBe(6);
  });

  it('AVERAGE calculates mean and returns #DIV/0! when no numbers', () => {
    expect(FUNCTIONS.AVERAGE(10, 20, 30)).toBe(20);
    expect(FUNCTIONS.AVERAGE(['text', ''])).toBe('#DIV/0!');
    // Alias MEAN
    expect(FUNCTIONS.MEAN(10, 20)).toBe(15);
  });

  it('MEDIAN handles odd, even counts and empty sets', () => {
    expect(FUNCTIONS.MEDIAN(1, 3, 2)).toBe(2);
    expect(FUNCTIONS.MEDIAN(1, 2, 3, 4)).toBe(2.5);
    expect(FUNCTIONS.MEDIAN([])).toBe('#DIV/0!');
  });

  it('MIN, MAX, COUNT, COUNTA', () => {
    expect(FUNCTIONS.MIN(5, 2, 8)).toBe(2);
    expect(FUNCTIONS.MAX(5, 2, 8)).toBe(8);
    expect(FUNCTIONS.COUNT([1, 'a', 2, '', false])).toBe(2);
    expect(FUNCTIONS.COUNTA([1, 'a', 2, '', false])).toBe(4);
  });

  it('PRODUCT multiplies numbers and skips text', () => {
    expect(FUNCTIONS.PRODUCT(2, 3, 4)).toBe(24);
    expect(FUNCTIONS.PRODUCT([2, 'text', 3])).toBe(6);
  });

  it('ROUND, ROUNDUP, ROUNDDOWN', () => {
    expect(FUNCTIONS.ROUND(2.55, 1)).toBe(2.6);
    expect(FUNCTIONS.ROUND(2.54, 1)).toBe(2.5);
    expect(FUNCTIONS.ROUNDUP(2.11, 1)).toBe(2.2);
    expect(FUNCTIONS.ROUNDDOWN(2.99, 1)).toBe(2.9);
  });

  it('ABS, INT, SQRT, POWER, MOD', () => {
    expect(FUNCTIONS.ABS(-15)).toBe(15);
    expect(FUNCTIONS.INT(4.9)).toBe(4);
    expect(FUNCTIONS.SQRT(16)).toBe(4);
    expect(FUNCTIONS.SQRT(-4)).toBe('#VALUE!');
    expect(FUNCTIONS.POWER(2, 3)).toBe(8);
    expect(FUNCTIONS.MOD(10, 3)).toBe(1);
    expect(FUNCTIONS.MOD(10, 0)).toBe('#DIV/0!');
    // Excel-style MOD with negative divisor
    expect(FUNCTIONS.MOD(-10, 3)).toBe(2);
  });

  it('IF, IFERROR, AND, OR, NOT', () => {
    expect(FUNCTIONS.IF(true, 'Yes', 'No')).toBe('Yes');
    expect(FUNCTIONS.IF(false, 'Yes', 'No')).toBe('No');
    expect(FUNCTIONS.IFERROR(10, 'Fallback')).toBe(10);
    expect(FUNCTIONS.IFERROR('#DIV/0!', 'Fallback')).toBe('Fallback');
    expect(FUNCTIONS.AND(true, 1, 'text')).toBe(true);
    expect(FUNCTIONS.AND(true, 0)).toBe(false);
    expect(FUNCTIONS.OR(false, 0, 1)).toBe(true);
    expect(FUNCTIONS.OR(false, 0)).toBe(false);
    expect(FUNCTIONS.NOT(true)).toBe(false);
    expect(FUNCTIONS.NOT(false)).toBe(true);
  });
});
