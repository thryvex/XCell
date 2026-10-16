import { describe, it, expect } from 'vitest';
import { moveRefsInFormula, CutMove } from '../../src/core/refshift';

/** ย้ายช่วง (r1..r2, c1..c2) ไป (dr, dc) แบบไม่มีแถวซ่อน */
function block(
  r1: number,
  c1: number,
  r2: number,
  c2: number,
  dr: number,
  dc: number,
  sourceSheet = 'ชีต1',
  destSheet = sourceSheet
): CutMove {
  const sourceRows = Array.from({ length: r2 - r1 + 1 }, (_, i) => r1 + i);
  return { sourceSheet, destSheet, sourceRows, c1, c2, destRows: sourceRows.map((r) => r + dr), dc };
}

const same = (f: string, m: CutMove) => moveRefsInFormula(f, 'ชีต1', 'ชีต1', m);

describe('moveRefsInFormula (ตัด-วาง SPEC 4.5 ข้อ 2)', () => {
  // A2 → A6
  const a2toA6 = block(1, 0, 1, 0, 4, 0);

  it('reference ที่ชี้เซลที่ถูกย้ายตามไปตำแหน่งใหม่', () => {
    expect(same('=A2*10', a2toA6)).toBe('=A6*10');
    expect(same('=A2+A2-B2', a2toA6)).toBe('=A6+A6-B2');
  });

  it('reference ที่ไม่เกี่ยวกับเซลที่ย้ายคงเดิม (ข้อความเดิมทุกตัวอักษร)', () => {
    expect(same('=a1 + b2 * 3', a2toA6)).toBe('=a1 + b2 * 3');
    expect(same('=SUM(A1:A3)', a2toA6)).toBe('=SUM(A1:A3)');
  });

  it('$ (absolute) ก็ตามไปเหมือน Google Sheets และคงเครื่องหมาย $', () => {
    expect(same('=$A$2+A$2+$A2', a2toA6)).toBe('=$A$6+A$6+$A6');
  });

  it('ช่วงที่ถูกย้ายทั้งก้อนตามไปทั้งช่วง; ย้ายแค่บางส่วนคงเดิม', () => {
    const m = block(0, 0, 2, 1, 0, 3); // A1:B3 → D1:E3
    expect(same('=SUM(A1:B3)', m)).toBe('=SUM(D1:E3)');
    expect(same('=SUM(A2:B2)', m)).toBe('=SUM(D2:E2)');
    expect(same('=SUM(A1:B4)', m)).toBe('=SUM(A1:B4)');
    expect(same('=SUM(A1:C3)', m)).toBe('=SUM(A1:C3)');
  });

  it('ย้ายช่วงหลายเซล: ทุกเซลในช่วงตามไปตามตำแหน่งในช่วง', () => {
    const m = block(0, 0, 2, 0, 5, 2); // A1:A3 → C6:C8
    expect(same('=A1+A2+A3+A4', m)).toBe('=C6+C7+C8+A4');
  });

  it('สูตรในเซลที่ย้ายเอง: อ้างเซลอื่นไม่เปลี่ยน, อ้างเซลในช่วงที่ย้ายด้วยกันตามไป', () => {
    const m = block(0, 2, 1, 2, 4, 0); // C1:C2 → C5:C6
    expect(same('=A1*$B$1', m)).toBe('=A1*$B$1');
    expect(same('=C1*2', m)).toBe('=C5*2');
  });

  it('ข้ามชีต: สูตรชีตอื่นที่อ้าง \'ชื่อชีต\'!A1 ตามไป และใส่ชื่อชีตปลายทางเมื่อย้ายข้ามชีต', () => {
    const within = block(1, 0, 1, 0, 4, 0, 'ข้อมูล ปี 69');
    expect(moveRefsInFormula("='ข้อมูล ปี 69'!A2*2", 'สรุป', 'สรุป', within)).toBe("='ข้อมูล ปี 69'!A6*2");
    expect(moveRefsInFormula('=A2*2', 'สรุป', 'สรุป', within)).toBe('=A2*2'); // A2 ของชีตสรุปไม่เกี่ยว

    const across = block(1, 0, 1, 0, 0, 1, 'ชีต1', 'ชีต2'); // ชีต1!A2 → ชีต2!B2
    expect(moveRefsInFormula('=A2+1', 'ชีต1', 'ชีต1', across)).toBe("='ชีต2'!B2+1");
    expect(moveRefsInFormula("='ชีต1'!A2+1", 'ชีต3', 'ชีต3', across)).toBe("='ชีต2'!B2+1");
    expect(moveRefsInFormula("='ชีต1'!A2+1", 'ชีต2', 'ชีต2', across)).toBe("='ชีต2'!B2+1"); // คงรูปแบบมีชื่อชีต
    // ชื่อชีตตัวพิมพ์ต่างกันถือเป็นชีตเดียวกัน
    expect(moveRefsInFormula("=sheet1!A2", 'X', 'X', block(1, 0, 1, 0, 1, 0, 'Sheet1'))).toBe('=sheet1!A3');
  });

  it('สูตรที่ถูกย้ายข้ามชีตเอง: อ้างเซลชีตเดิมที่ไม่ได้ย้ายจะได้ชื่อชีตเดิม', () => {
    const across = block(0, 1, 0, 1, 0, 0, 'ชีต1', 'ชีต2'); // ชีต1!B1 (=A1*2) → ชีต2!B1
    expect(moveRefsInFormula('=A1*2', 'ชีต1', 'ชีต2', across)).toBe("='ชีต1'!A1*2");
    expect(moveRefsInFormula('=B1+1', 'ชีต1', 'ชีต2', across)).toBe('=B1+1'); // อ้างตัวเองที่ย้ายไปด้วย
  });

  it('แถวที่ตัวกรองซ่อน: แถวที่ไม่ได้ถูกตัดคงเดิม แถวที่ตัดตามไปแถวปลายทางของมัน', () => {
    const m: CutMove = { sourceSheet: 'ชีต1', destSheet: 'ชีต1', sourceRows: [0, 2], c1: 0, c2: 0, destRows: [9, 11], dc: 0 };
    expect(same('=A1+A2+A3', m)).toBe('=A10+A2+A12');
    expect(same('=SUM(A1:A3)', m)).toBe('=SUM(A1:A3)'); // A2 ในช่วงไม่ได้ถูกย้าย
  });

  it('ไม่ใช่สูตร / ไม่มี reference คืนค่าเดิม', () => {
    expect(same('A2', a2toA6)).toBe('A2');
    expect(same('="A2"&1', a2toA6)).toBe('="A2"&1');
  });
});
