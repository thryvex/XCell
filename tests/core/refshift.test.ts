import { describe, it, expect } from 'vitest';
import {
  shiftFormula,
  shiftFormulaOnInsert,
  shiftFormulaOnDelete,
} from '../../src/core/refshift';

describe('src/core/refshift (≥ 30 test cases)', () => {
  // 1-18: Copy/Drag/Fill tests
  it('1. drags relative formula down by 1 row (=D2-E2 -> =D3-E3)', () => {
    expect(shiftFormula('=D2-E2', 1, 0)).toBe('=D3-E3');
  });

  it('2. drags relative formula down 50 rows (=D2-E2 -> =D52-E52)', () => {
    expect(shiftFormula('=D2-E2', 50, 0)).toBe('=D52-E52');
  });

  it('3. drags absolute $A$1 down 1 row (remains $A$1)', () => {
    expect(shiftFormula('=$A$1*B2', 1, 0)).toBe('=$A$1*B3');
  });

  it('4. drags absolute $A$1 right 1 col (remains $A$1)', () => {
    expect(shiftFormula('=$A$1*B2', 0, 1)).toBe('=$A$1*C2');
  });

  it('5. drags row-absolute A$1 down 1 row (row fixed)', () => {
    expect(shiftFormula('=A$1*B2', 1, 0)).toBe('=A$1*B3');
  });

  it('6. drags row-absolute A$1 right 1 col (col shifts)', () => {
    expect(shiftFormula('=A$1*B2', 0, 1)).toBe('=B$1*C2');
  });

  it('7. drags col-absolute $A1 down 1 row (row shifts)', () => {
    expect(shiftFormula('=$A1*B2', 1, 0)).toBe('=$A2*B3');
  });

  it('8. drags col-absolute $A1 right 1 col (col fixed)', () => {
    expect(shiftFormula('=$A1*B2', 0, 1)).toBe('=$A1*C2');
  });

  it('9. drags out of bounds above row 1 yields #REF!', () => {
    expect(shiftFormula('=A1', -1, 0)).toBe('=#REF!');
  });

  it('10. drags out of bounds left of col A yields #REF!', () => {
    expect(shiftFormula('=A1', 0, -1)).toBe('=#REF!');
  });

  it('11. drags relative range down 2 rows', () => {
    expect(shiftFormula('=SUM(A1:B10)', 2, 0)).toBe('=SUM(A3:B12)');
  });

  it('12. drags fully absolute range down 2 rows', () => {
    expect(shiftFormula('=SUM($A$1:$B$10)', 2, 0)).toBe('=SUM($A$1:$B$10)');
  });

  it('13. drags row-absolute range down 2 rows', () => {
    expect(shiftFormula('=SUM(A$1:B$10)', 2, 0)).toBe('=SUM(A$1:B$10)');
  });

  it('14. drags col-absolute range down 2 rows', () => {
    expect(shiftFormula('=SUM($A1:$B10)', 2, 0)).toBe('=SUM($A3:$B12)');
  });

  it('15. drags col-absolute range right 1 col', () => {
    expect(shiftFormula('=SUM($A1:$B10)', 0, 1)).toBe('=SUM($A1:$B10)');
  });

  it('16. drags diagonally (dr=3, dc=2)', () => {
    expect(shiftFormula('=A1+B1', 3, 2)).toBe('=C4+D4');
  });

  it('17. preserves formula whitespace and operators', () => {
    expect(shiftFormula('=  A1  +  $B$2  * 3', 1, 0)).toBe('=  A2  +  $B$2  * 3');
  });

  it('18. preserves string literals containing cell-like text', () => {
    expect(shiftFormula('="A1" & A1', 1, 0)).toBe('="A1" & A2');
  });

  // 19-25: Insert row/col tests
  it('19. inserts row above reference shifts reference down', () => {
    expect(shiftFormulaOnInsert('=A2', 'row', 1, 1)).toBe('=A3');
  });

  it('20. inserts row below reference leaves reference unchanged', () => {
    expect(shiftFormulaOnInsert('=A2', 'row', 2, 1)).toBe('=A2');
  });

  it('21. inserts row inside range expands range', () => {
    expect(shiftFormulaOnInsert('=SUM(A1:A5)', 'row', 2, 1)).toBe('=SUM(A1:A6)');
  });

  it('22. inserts row above range shifts range', () => {
    expect(shiftFormulaOnInsert('=SUM(A3:A5)', 'row', 1, 1)).toBe('=SUM(A4:A6)');
  });

  it('23. inserts row below range leaves range unchanged', () => {
    expect(shiftFormulaOnInsert('=SUM(A1:A5)', 'row', 6, 1)).toBe('=SUM(A1:A5)');
  });

  it('24. inserts col to the left shifts column', () => {
    expect(shiftFormulaOnInsert('=B1', 'col', 1, 1)).toBe('=C1');
  });

  it('25. inserts col inside range expands range', () => {
    expect(shiftFormulaOnInsert('=SUM(A1:C1)', 'col', 1, 1)).toBe('=SUM(A1:D1)');
  });

  // 26-34: Delete row/col tests
  it('26. deletes referenced row turns into #REF!', () => {
    expect(shiftFormulaOnDelete('=A2', 'row', 1, 1)).toBe('=#REF!');
  });

  it('27. deletes row above shifts reference up', () => {
    expect(shiftFormulaOnDelete('=A3', 'row', 1, 1)).toBe('=A2');
  });

  it('28. deletes row below leaves reference unchanged', () => {
    expect(shiftFormulaOnDelete('=A1', 'row', 2, 1)).toBe('=A1');
  });

  it('29. deletes row inside range contracts range', () => {
    expect(shiftFormulaOnDelete('=SUM(A1:A5)', 'row', 2, 1)).toBe('=SUM(A1:A4)');
  });

  it('30. deletes whole range turns into #REF!', () => {
    expect(shiftFormulaOnDelete('=SUM(A2:A3)', 'row', 1, 2)).toBe('=SUM(#REF!)');
  });

  it('31. deletes referenced column turns into #REF!', () => {
    expect(shiftFormulaOnDelete('=B1', 'col', 1, 1)).toBe('=#REF!');
  });

  it('32. deletes column inside range contracts range', () => {
    expect(shiftFormulaOnDelete('=SUM(A1:D1)', 'col', 1, 1)).toBe('=SUM(A1:C1)');
  });

  it('33. shifts multiple references in formula', () => {
    expect(shiftFormula('=A1+B2+C3', 1, 0)).toBe('=A2+B3+C4');
  });

  it('34. shifts formula inside complex function calls', () => {
    expect(shiftFormula('=IF(A1>0, $B$1, C1)', 1, 0)).toBe('=IF(A2>0, $B$1, C2)');
  });
});
