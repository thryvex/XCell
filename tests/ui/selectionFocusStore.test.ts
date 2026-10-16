import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';

const S = () => useSheetStore.getState();

describe('การเลือก: anchor (เซลที่ใช้งานอยู่) กับ focus (ขอบที่กำลังขยาย)', () => {
  beforeEach(() => {
    S().newWorkbook();
    S().setActiveCell(4, 2); // C5
  });

  it('ขยายการเลือกโดยเซลที่ใช้งานอยู่ไม่ย้าย และ focus อยู่ที่ปลายที่ขยาย', () => {
    S().extendSelectionTo(9, 4); // E10
    expect(S().activeCell).toEqual({ r: 4, c: 2 });
    expect(S().selection).toEqual({ r1: 4, c1: 2, r2: 9, c2: 4 });
    expect(S().getSelectionFocus()).toEqual({ r: 9, c: 4 });
  });

  it('หดการเลือกกลับได้ และขยายข้ามเซลที่ใช้งานอยู่ไปอีกฝั่งได้', () => {
    S().extendSelectionTo(9, 4);
    S().extendSelectionTo(6, 4);
    expect(S().selection).toEqual({ r1: 4, c1: 2, r2: 6, c2: 4 });
    S().extendSelectionTo(1, 0);
    expect(S().selection).toEqual({ r1: 1, c1: 0, r2: 4, c2: 2 });
    expect(S().activeCell).toEqual({ r: 4, c: 2 });
  });

  it('ไม่ขยายเกินขอบตารางที่มีอยู่', () => {
    S().extendSelectionTo(5000, 500);
    expect(S().selection.r2).toBe(S().rowCount - 1);
    expect(S().selection.c2).toBe(S().colCount - 1);
  });

  it('เลือกเซลใหม่ / เลือกด้วยเมาส์ ล้าง focus ที่เก็บไว้ แล้วใช้มุมตรงข้ามแทน', () => {
    S().extendSelectionTo(9, 4);
    S().setActiveCell(0, 0);
    expect(S().getSelectionFocus()).toEqual({ r: 0, c: 0 });
    S().setSelection(0, 0, 3, 3);
    expect(S().getSelectionFocus()).toEqual({ r: 3, c: 3 });
  });

  it('scrollToCell ตั้งเป้าหมายการเลื่อนและเพิ่ม token', () => {
    const before = S().scrollToActiveToken;
    S().scrollToCell(50, 3);
    expect(S().scrollTarget).toEqual({ r: 50, c: 3 });
    expect(S().scrollToActiveToken).toBe(before + 1);
  });
});
