import { describe, it, expect } from 'vitest';
import {
  formatDisplayValue,
  parseDateString,
  formatDate,
  serialToDate,
  dateToSerial,
  ERROR_TOOLTIP_MAP,
  THAI_TO_ERROR_MAP,
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

  it('converts error codes to Thai user-friendly text and tooltips', () => {
    expect(formatDisplayValue('#DIV/0!')).toBe('#หารศูนย์');
    expect(formatDisplayValue('#VALUE!')).toBe('#ค่าผิด');
    expect(formatDisplayValue('#REF!')).toBe('#อ้างอิงหาย');
    expect(formatDisplayValue('#NAME?')).toBe('#ไม่รู้จักชื่อ');
    expect(formatDisplayValue('#CYCLE!')).toBe('#วนซ้ำ');
    expect(formatDisplayValue('#ERR')).toBe('#สูตรผิด');

    // Tooltip mapping for both English code and Thai display name
    expect(ERROR_TOOLTIP_MAP['#DIV/0!']).toContain('หารด้วยศูนย์ไม่ได้');
    expect(ERROR_TOOLTIP_MAP['#หารศูนย์']).toContain('หารด้วยศูนย์ไม่ได้');
    expect(ERROR_TOOLTIP_MAP['#REF!']).toContain('เซลที่สูตรอ้างถึงถูกลบไปแล้ว');
    expect(ERROR_TOOLTIP_MAP['#อ้างอิงหาย']).toContain('เซลที่สูตรอ้างถึงถูกลบไปแล้ว');
    expect(ERROR_TOOLTIP_MAP['#CYCLE!']).toContain('สูตรอ้างอิงวนกลับมาหาตัวเอง');
    expect(ERROR_TOOLTIP_MAP['#วนซ้ำ']).toContain('สูตรอ้างอิงวนกลับมาหาตัวเอง');

    // Reverse map
    expect(THAI_TO_ERROR_MAP['#หารศูนย์']).toBe('#DIV/0!');
    expect(THAI_TO_ERROR_MAP['#ค่าผิด']).toBe('#VALUE!');
    expect(THAI_TO_ERROR_MAP['#อ้างอิงหาย']).toBe('#REF!');
    expect(THAI_TO_ERROR_MAP['#ไม่รู้จักชื่อ']).toBe('#NAME?');
    expect(THAI_TO_ERROR_MAP['#วนซ้ำ']).toBe('#CYCLE!');
    expect(THAI_TO_ERROR_MAP['#สูตรผิด']).toBe('#ERR');
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
