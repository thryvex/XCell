import { describe, it, expect, beforeEach } from 'vitest';
import { useSheetStore } from '../../src/ui/store';
import { parseA1 } from '../../src/core/address';

const S = () => useSheetStore.getState();
const pos = (a1: string) => parseA1(a1)!;
const type = (a1: string, text: string) => {
  const p = pos(a1);
  S().setActiveCell(p.r, p.c);
  S().startEditing(text);
  S().commitEditing(0, 0);
};
const raw = (a1: string, sheet = S().activeSheetIndex) => S().workbook.sheets[sheet].cells[a1]?.raw;
const value = (a1: string) => {
  const p = pos(a1);
  return S().getComputedValue(p.r, p.c);
};
const select = (from: string, to = from) => {
  const a = pos(from);
  const b = pos(to);
  S().setActiveCell(a.r, a.c);
  S().setSelection(a.r, a.c, b.r, b.c);
};
const cutTo = async (from: string, to: string, dest: string) => {
  select(from, to);
  S().cutSelection();
  select(dest);
  await S().paste();
};

describe('ตัดแล้ววาง: reference ตามไป (SPEC 4.5 ข้อ 2)', () => {
  beforeEach(() => S().newWorkbook());

  it('ชีตเดียว: สูตรที่อ้างเซลที่ย้ายตามไป และค่ายังถูกต้อง', async () => {
    type('A2', '7');
    type('G1', '=A2*10');
    type('C1', '=A2+$A$2');
    await cutTo('A2', 'A2', 'A6');
    expect(raw('A2')).toBeUndefined();
    expect(raw('A6')).toBe('7');
    expect(raw('G1')).toBe('=A6*10');
    expect(raw('C1')).toBe('=A6+$A$6');
    expect(value('G1')).toBe(70);
    expect(value('C1')).toBe(14);
  });

  it('ตัดเป็นช่วง: สูตรที่อ้างทั้งช่วงย้ายช่วงตาม, อ้างบางเซลย้ายตามเซล, อ้างช่วงที่ย้ายแค่บางส่วนคงเดิม', async () => {
    type('A1', '1');
    type('A2', '2');
    type('A3', '3');
    type('A4', '4');
    type('E1', '=SUM(A1:A3)');
    type('E2', '=A2');
    type('E3', '=SUM(A1:A4)');
    await cutTo('A1', 'A3', 'C5');
    expect(raw('E1')).toBe('=SUM(C5:C7)');
    expect(raw('E2')).toBe('=C6');
    expect(raw('E3')).toBe('=SUM(A1:A4)');
    expect(value('E1')).toBe(6);
    expect(value('E2')).toBe(2);
    expect(value('E3')).toBe(4); // A1:A3 ว่างแล้ว เหลือ A4
  });

  it('สูตรในเซลที่ถูกย้ายเอง: อ้างเซลอื่นไม่เปลี่ยน (ต่างจากคัดลอก)', async () => {
    type('A1', '5');
    type('B1', '2');
    type('C1', '=A1*$B$1');
    await cutTo('C1', 'C1', 'C5');
    expect(raw('C5')).toBe('=A1*$B$1');
    expect(value('C5')).toBe(10);

    // เทียบ: คัดลอกแล้ววาง เลื่อนตามตำแหน่ง
    select('C5');
    S().copySelection();
    select('C6');
    await S().paste();
    expect(raw('C6')).toBe('=A2*$B$1');
  });

  it('สูตรในช่วงที่ย้ายด้วยกัน อ้างกันเองตามไปด้วย', async () => {
    type('A1', '3');
    type('A2', '=A1*2');
    await cutTo('A1', 'A2', 'D1');
    expect(raw('D2')).toBe('=D1*2');
    expect(value('D2')).toBe(6);
  });

  it("ข้ามชีต: สูตรชีตอื่นที่อ้าง 'ชื่อชีต'!A1 ตามไป", async () => {
    S().renameSheet(0, 'ข้อมูล ปี 69');
    type('A2', '100');
    S().addSheet(); // ชีตที่ 2 เปิดอยู่
    type('B1', "='ข้อมูล ปี 69'!A2*2");
    S().switchSheet(0);
    await cutTo('A2', 'A2', 'C9');
    expect(raw('B1', 1)).toBe("='ข้อมูล ปี 69'!C9*2");
    S().switchSheet(1);
    expect(value('B1')).toBe(200);
  });

  it('ข้ามชีต: ตัดจากชีตหนึ่งไปวางอีกชีต สูตรทั้งสองชีตยังชี้เซลเดิม', async () => {
    type('A2', '8');
    type('B1', '=A2+1');
    S().addSheet();
    const second = S().workbook.sheets[1].name;
    S().switchSheet(0);
    select('A2');
    S().cutSelection();
    S().switchSheet(1);
    select('D4');
    await S().paste();
    expect(raw('D4', 1)).toBe('8');
    expect(raw('A2', 0)).toBeUndefined();
    expect(raw('B1', 0)).toBe(`='${second}'!D4+1`);
    S().switchSheet(0);
    expect(value('B1')).toBe(9);
  });

  it('Undo คืนทั้งตำแหน่งและสูตรที่อ้างอิง (รวมชีตอื่น), Redo ทำซ้ำได้', async () => {
    type('A2', '7');
    type('G1', '=A2*10');
    S().addSheet();
    type('A1', "='ชีต1'!A2");
    S().switchSheet(0);
    await cutTo('A2', 'A2', 'A6');
    expect(raw('G1')).toBe('=A6*10');
    expect(raw('A1', 1)).toBe("='ชีต1'!A6");

    S().undo();
    expect(raw('A2', 0)).toBe('7');
    expect(raw('A6', 0)).toBeUndefined();
    expect(raw('G1', 0)).toBe('=A2*10');
    expect(raw('A1', 1)).toBe("='ชีต1'!A2");
    expect(value('G1')).toBe(70);

    S().redo();
    expect(raw('A6', 0)).toBe('7');
    expect(raw('G1', 0)).toBe('=A6*10');
    expect(raw('A1', 1)).toBe("='ชีต1'!A6");
  });

  it('ตัดเซลเดียวแล้ววางบนช่วงหลายเซล: ย้ายครั้งเดียวที่เซลปัจจุบัน (ไม่เติมซ้ำ)', async () => {
    type('A1', '4');
    type('B1', '=A1');
    select('A1');
    S().cutSelection();
    S().setActiveCell(4, 0);
    S().setSelection(4, 0, 6, 1);
    await S().paste();
    expect(raw('A5')).toBe('4');
    expect(raw('A6')).toBeUndefined();
    expect(raw('B6')).toBeUndefined();
    expect(raw('B1')).toBe('=A5');
  });

  it('คัดลอกแล้ววาง (ไม่ใช่ตัด) ไม่เปลี่ยนสูตรที่อ้างต้นทาง', async () => {
    type('A2', '7');
    type('G1', '=A2*10');
    select('A2');
    S().copySelection();
    select('A6');
    await S().paste();
    expect(raw('G1')).toBe('=A2*10');
    expect(raw('A2')).toBe('7');
  });
});
