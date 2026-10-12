import { describe, it, expect } from 'vitest';
import { detectSigmaRange } from '../../src/core/sigma';
import { createEmptySheet } from '../../src/core/model';

describe('src/core/sigma', () => {
  it('detects range vertically above active cell', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['A2'] = { raw: '20' };
    sheet.cells['A3'] = { raw: '30' };

    // In A4, should detect A1:A3
    const range = detectSigmaRange(sheet, 3, 0);
    expect(range).toBe('A1:A3');
  });

  it('detects range horizontally to the left when no numbers above', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: '10' };
    sheet.cells['B1'] = { raw: '20' };
    sheet.cells['C1'] = { raw: '30' };

    // In D1, should detect A1:C1
    const range = detectSigmaRange(sheet, 0, 3);
    expect(range).toBe('A1:C1');
  });

  it('returns null when empty or no adjacent numbers', () => {
    const sheet = createEmptySheet();
    const range = detectSigmaRange(sheet, 5, 5);
    expect(range).toBeNull();
  });
});
