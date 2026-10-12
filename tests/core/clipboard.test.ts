import { describe, it, expect } from 'vitest';
import { rangeToTsv, rangeToHtml, parseTsv } from '../../src/core/clipboard';
import { createEmptySheet } from '../../src/core/model';

describe('src/core/clipboard', () => {
  it('formats range to TSV and HTML', () => {
    const sheet = createEmptySheet();
    sheet.cells['A1'] = { raw: 'วันที่' };
    sheet.cells['B1'] = { raw: '100' };

    const tsv = rangeToTsv(sheet, { r1: 0, c1: 0, r2: 0, c2: 1 }, (r, c) =>
      r === 0 && c === 0 ? 'วันที่' : '100'
    );
    expect(tsv).toBe('วันที่\t100');

    const html = rangeToHtml(sheet, { r1: 0, c1: 0, r2: 0, c2: 1 }, (r, c) =>
      r === 0 && c === 0 ? 'วันที่' : '100'
    );
    expect(html).toContain('<td>วันที่</td>');
    expect(html).toContain('<td>100</td>');
  });

  it('parses TSV string from clipboard into 2D array', () => {
    const text = 'A\tB\tC\n1\t2\t3';
    const grid = parseTsv(text);
    expect(grid).toEqual([
      ['A', 'B', 'C'],
      ['1', '2', '3'],
    ]);
  });
});
