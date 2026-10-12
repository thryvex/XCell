import { describe, it, expect } from 'vitest';
import {
  colToStr,
  strToCol,
  cellToA1,
  parseA1,
  parseRange,
  rangeToA1,
  isCellInRange,
  cycleReference,
} from '../../src/core/address';

describe('src/core/address', () => {
  it('converts column index to letters and back', () => {
    expect(colToStr(0)).toBe('A');
    expect(colToStr(25)).toBe('Z');
    expect(colToStr(26)).toBe('AA');
    expect(colToStr(27)).toBe('AB');
    expect(colToStr(51)).toBe('AZ');
    expect(colToStr(52)).toBe('BA');
    expect(colToStr(701)).toBe('ZZ');

    expect(strToCol('A')).toBe(0);
    expect(strToCol('Z')).toBe(25);
    expect(strToCol('AA')).toBe(26);
    expect(strToCol('AB')).toBe(27);
    expect(strToCol('AZ')).toBe(51);
    expect(strToCol('BA')).toBe(52);
    expect(strToCol('ZZ')).toBe(701);
  });

  it('formats cell coordinates to A1 notation', () => {
    expect(cellToA1(0, 0)).toBe('A1');
    expect(cellToA1(2, 3)).toBe('D3');
    expect(cellToA1(0, 0, true, false)).toBe('$A1');
    expect(cellToA1(0, 0, false, true)).toBe('A$1');
    expect(cellToA1(0, 0, true, true)).toBe('$A$1');
  });

  it('parses A1 notation with absolute/relative signs', () => {
    expect(parseA1('A1')).toEqual({ r: 0, c: 0, rowAbs: false, colAbs: false });
    expect(parseA1('b12')).toEqual({ r: 11, c: 1, rowAbs: false, colAbs: false });
    expect(parseA1('$A1')).toEqual({ r: 0, c: 0, rowAbs: false, colAbs: true });
    expect(parseA1('A$1')).toEqual({ r: 0, c: 0, rowAbs: true, colAbs: false });
    expect(parseA1('$A$1')).toEqual({ r: 0, c: 0, rowAbs: true, colAbs: true });
    expect(parseA1('$ZZ$1000')).toEqual({ r: 999, c: 701, rowAbs: true, colAbs: true });
    expect(parseA1('invalid')).toBeNull();
  });

  it('parses ranges', () => {
    expect(parseRange('A1:C5')).toEqual({
      r1: 0,
      c1: 0,
      r2: 4,
      c2: 2,
      startColAbs: false,
      startRowAbs: false,
      endColAbs: false,
      endRowAbs: false,
    });
    // Inverted ranges should normalize to top-left / bottom-right
    expect(parseRange('C5:A1')).toEqual({
      r1: 0,
      c1: 0,
      r2: 4,
      c2: 2,
      startColAbs: false,
      startRowAbs: false,
      endColAbs: false,
      endRowAbs: false,
    });
  });

  it('checks isCellInRange correctly', () => {
    const range = { r1: 1, c1: 1, r2: 5, c2: 5 };
    expect(isCellInRange(1, 1, range)).toBe(true);
    expect(isCellInRange(3, 3, range)).toBe(true);
    expect(isCellInRange(5, 5, range)).toBe(true);
    expect(isCellInRange(0, 1, range)).toBe(false);
    expect(isCellInRange(6, 2, range)).toBe(false);
  });

  it('cycles references with F4 (cycleReference)', () => {
    expect(cycleReference('A1')).toBe('$A$1');
    expect(cycleReference('$A$1')).toBe('A$1');
    expect(cycleReference('A$1')).toBe('$A1');
    expect(cycleReference('$A1')).toBe('A1');
    expect(cycleReference('A1:B10')).toBe('$A$1:$B$10');
  });
});
