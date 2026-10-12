import { describe, it, expect } from 'vitest';
import { convertShorthand } from '../../src/core/shorthand';

describe('src/core/shorthand', () => {
  it('converts basic expressions with cell references', () => {
    expect(convertShorthand('a1+c1')).toBe('=A1+C1');
    expect(convertShorthand('b1*c1')).toBe('=B1*C1');
    expect(convertShorthand('a1-b1')).toBe('=A1-B1');
    expect(convertShorthand('a1/b1')).toBe('=A1/B1');
    expect(convertShorthand('a1+5')).toBe('=A1+5');
  });

  it('converts English function shorthands', () => {
    expect(convertShorthand('sum a1:a10')).toBe('=SUM(A1:A10)');
    expect(convertShorthand('avg a1,b1,c1')).toBe('=AVERAGE(A1,B1,C1)');
    expect(convertShorthand('mean a1:a5')).toBe('=AVERAGE(A1:A5)');
    expect(convertShorthand('min a1,a2')).toBe('=MIN(A1,A2)');
    expect(convertShorthand('max a1,a2')).toBe('=MAX(A1,A2)');
    expect(convertShorthand('count a1:a5')).toBe('=COUNT(A1:A5)');
    expect(convertShorthand('median a1:a5')).toBe('=MEDIAN(A1:A5)');
  });

  it('converts Thai function shorthands', () => {
    expect(convertShorthand('รวม a1:a10')).toBe('=SUM(A1:A10)');
    expect(convertShorthand('เฉลี่ย a1:a10')).toBe('=AVERAGE(A1:A10)');
    expect(convertShorthand('มัธยฐาน a1:a10')).toBe('=MEDIAN(A1:A10)');
    expect(convertShorthand('ต่ำสุด a1,b1')).toBe('=MIN(A1,B1)');
    expect(convertShorthand('สูงสุด a1,b1')).toBe('=MAX(A1,B1)');
    expect(convertShorthand('นับ a1:a10')).toBe('=COUNT(A1:A10)');
  });

  it('does NOT convert general text or invalid inputs', () => {
    // Normal Thai words starting with alias root
    expect(convertShorthand('นับถือ')).toBeNull();
    expect(convertShorthand('รวมกัน')).toBeNull();
    expect(convertShorthand('เฉลี่ยงาน')).toBeNull();

    // Plain text without formula meaning
    expect(convertShorthand('a1')).toBeNull();
    expect(convertShorthand('123')).toBeNull();
    expect(convertShorthand('เงินเดือน')).toBeNull();

    // Already formula
    expect(convertShorthand('=A1+B1')).toBeNull();

    // Escaped text
    expect(convertShorthand("'a1+b1")).toBeNull();
  });
});
