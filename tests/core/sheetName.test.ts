import { describe, it, expect } from 'vitest';
import {
  validateSheetName,
  generateNextSheetName,
  generateDuplicateSheetName,
} from '../../src/core/sheetName';
import { th } from '../../src/i18n/th';

describe('validateSheetName (SPEC 5.13 / ข้อตกลง 16.6)', () => {
  const existing = ['ชีต1', 'Data', 'รายรับ'];

  it('accepts a valid Thai name', () => {
    expect(validateSheetName('สรุป', existing).valid).toBe(true);
  });

  it('rejects empty and whitespace-only names', () => {
    expect(validateSheetName('', existing)).toEqual({ valid: false, error: th.sheetTabs.nameEmpty });
    expect(validateSheetName('   ', existing)).toEqual({ valid: false, error: th.sheetTabs.nameEmpty });
  });

  it('accepts exactly 31 characters and rejects 32', () => {
    expect(validateSheetName('a'.repeat(31), existing).valid).toBe(true);
    expect(validateSheetName('a'.repeat(32), existing)).toEqual({
      valid: false,
      error: th.sheetTabs.nameTooLong,
    });
  });

  it.each(['[', ']', ':', '*', '?', '/', '\\'])('rejects forbidden character %s', (ch) => {
    expect(validateSheetName(`ชีต${ch}ใหม่`, existing)).toEqual({
      valid: false,
      error: th.sheetTabs.nameInvalidChars,
    });
  });

  it('rejects duplicates case-insensitively', () => {
    const res = validateSheetName('DATA', existing);
    expect(res.valid).toBe(false);
    expect(res.error).toBe(th.sheetTabs.nameDuplicate.replace('{name}', 'DATA'));
    expect(validateSheetName('ชีต1', existing).valid).toBe(false);
  });

  it('allows renaming a sheet to a different case of its own name', () => {
    expect(validateSheetName('data', existing, 'Data').valid).toBe(true);
  });

  it('allows spaces and apostrophes inside the name', () => {
    expect(validateSheetName("ยอด ขาย's", existing).valid).toBe(true);
  });
});

describe('default and duplicate sheet names', () => {
  it('generates ชีตN that is not taken', () => {
    expect(generateNextSheetName(['ชีต1'])).toBe('ชีต2');
    expect(generateNextSheetName(['ชีต1', 'ชีต2', 'ชีต3'])).toBe('ชีต4');
    expect(generateNextSheetName(['ชีต1', 'ชีต3'])).toBe('ชีต4');
    expect(generateNextSheetName(['ชีต2', 'ชีต3'])).toBe('ชีต4');
    expect(generateNextSheetName(['Data'])).toBe('ชีต2');
  });

  it('generates สำเนาของ X and numbers further copies', () => {
    expect(generateDuplicateSheetName('รายรับ', ['รายรับ'])).toBe('สำเนาของ รายรับ');
    expect(generateDuplicateSheetName('รายรับ', ['รายรับ', 'สำเนาของ รายรับ'])).toBe('สำเนาของ รายรับ (2)');
  });

  it('keeps duplicate names within 31 characters', () => {
    const longName = 'ก'.repeat(31);
    const first = generateDuplicateSheetName(longName, [longName]);
    expect(first.length).toBeLessThanOrEqual(31);
    const second = generateDuplicateSheetName(longName, [longName, first]);
    expect(second.length).toBeLessThanOrEqual(31);
    expect(second).not.toBe(first);
  });
});
