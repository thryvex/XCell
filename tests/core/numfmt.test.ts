import { describe, it, expect } from 'vitest';
import {
  formatDisplayValue,
  parseDateString,
  formatDate,
  serialToDate,
  dateToSerial,
} from '../../src/core/numfmt';

describe('src/core/numfmt', () => {
  it('formats baht currency with commas and ฿ symbol', () => {
    expect(formatDisplayValue(1234.5, 'baht')).toBe('฿1,234.50');
    expect(formatDisplayValue(25000, 'baht')).toBe('฿25,000.00');
    expect(formatDisplayValue(0, 'baht')).toBe('฿0.00');
  });

  it('formats percent with 2 decimals', () => {
    expect(formatDisplayValue(0.125, 'percent')).toBe('12.50%');
    expect(formatDisplayValue(1, 'percent')).toBe('100.00%');
  });

  it('formats numbers with 2 decimals and integer format', () => {
    expect(formatDisplayValue(1234.567, 'number2')).toBe('1,234.57');
    expect(formatDisplayValue(1234.567, 'integer')).toBe('1,235');
  });

  it('converts error codes to Thai user-friendly text', () => {
    expect(formatDisplayValue('#DIV/0!')).toBe('#หารศูนย์');
    expect(formatDisplayValue('#VALUE!')).toBe('#ค่าผิด');
    expect(formatDisplayValue('#REF!')).toBe('#อ้างอิงหาย');
    expect(formatDisplayValue('#CYCLE!')).toBe('#วนซ้ำ');
    expect(formatDisplayValue('#NAME?')).toBe('#ไม่รู้จักชื่อ');
  });

  it('parses and formats dates (Day/Month/Year)', () => {
    const serial = parseDateString('01/10/2026');
    expect(serial).not.toBeNull();
    expect(formatDate(serial!)).toBe('01/10/2026');
    expect(formatDisplayValue(serial!, 'date')).toBe('01/10/2026');

    // Hyphen delimiter
    const serial2 = parseDateString('15-05-2025');
    expect(formatDate(serial2!)).toBe('15/05/2025');

    // Short year or omitted year
    const serial3 = parseDateString('1/10');
    expect(serial3).not.toBeNull();
  });
});
