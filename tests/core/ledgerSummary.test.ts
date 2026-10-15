import { describe, it, expect, beforeEach } from 'vitest';
import {
  createLedgerWorkbook,
  SUMMARY_LAYOUT as L,
  SUMMARY_CATEGORIES,
  THAI_MONTHS,
} from '../../src/templates/ledger';
import { Evaluator } from '../../src/core/evaluator';
import { Workbook } from '../../src/core/model';
import { dateToSerial } from '../../src/core/numfmt';
import { renameSheet } from '../../src/core/workbookOps';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import { useSheetStore } from '../../src/ui/store';

interface Entry {
  y: number;
  m: number; // 1-12
  d: number;
  cat: string;
  income: number;
  expense: number;
  typed?: boolean; // พิมพ์วันที่เป็นข้อความ (เหมือนผู้ใช้พิมพ์ 31/1/2026)
}

const ENTRIES: Entry[] = [
  { y: 2025, m: 12, d: 31, cat: 'อาหาร', income: 0, expense: 999 }, // ปีก่อน
  { y: 2026, m: 1, d: 1, cat: 'อาหาร', income: 0, expense: 120 }, // วันแรกของปี/เดือน
  { y: 2026, m: 1, d: 31, cat: 'เดินทาง', income: 0, expense: 45, typed: true }, // วันสุดท้ายของเดือน
  { y: 2026, m: 2, d: 1, cat: 'ที่พัก', income: 0, expense: 3000 },
  { y: 2026, m: 2, d: 28, cat: 'อื่นๆ', income: 30000, expense: 0, typed: true },
  { y: 2026, m: 3, d: 15, cat: 'อาหาร', income: 0, expense: 80.5 },
  { y: 2026, m: 3, d: 31, cat: 'ของใช้', income: 0, expense: 250 }, // หมวดที่ผู้ใช้เพิ่มเอง
  { y: 2026, m: 12, d: 31, cat: 'อื่นๆ', income: 500, expense: 0 }, // วันสุดท้ายของปี
  { y: 2027, m: 1, d: 1, cat: 'อาหาร', income: 0, expense: 777 }, // ปีถัดไป
  { y: 2026, m: 6, d: 10, cat: '', income: 0, expense: 60 }, // ไม่ระบุหมวด
];

function sumWhere(pred: (e: Entry) => boolean, field: 'income' | 'expense') {
  return ENTRIES.filter(pred).reduce((s, e) => s + e[field], 0);
}

function fillLedger(wb: Workbook, entries: Entry[]) {
  const ledger = wb.sheets[0];
  for (let r = 2; r <= 4; r++) {
    for (const c of 'ABCDEF') delete ledger.cells[`${c}${r}`];
  }
  entries.forEach((e, i) => {
    const r = i + 2;
    ledger.cells[`A${r}`] = {
      raw: e.typed ? `${e.d}/${e.m}/${e.y}` : String(dateToSerial(new Date(e.y, e.m - 1, e.d))),
      fmt: 'date',
    };
    if (e.cat) ledger.cells[`C${r}`] = { raw: e.cat };
    ledger.cells[`D${r}`] = { raw: String(e.income), fmt: 'baht' };
    ledger.cells[`E${r}`] = { raw: String(e.expense), fmt: 'baht' };
  });
}

function setYear(wb: Workbook, year: number) {
  wb.sheets[1].cells['B1'] = { ...wb.sheets[1].cells['B1'], raw: String(year) };
}

const v = (ev: Evaluator, a1: string) => ev.getCellValue(a1, 1);
const monthRow = (m: number) => L.monthFirstRow + m - 1;

describe('เทมเพลตรายรับรายจ่าย: ชีตสรุป', () => {
  let wb: Workbook;
  beforeEach(() => {
    wb = createLedgerWorkbook(new Date(2026, 9, 16));
    fillLedger(wb, ENTRIES);
    setYear(wb, 2026);
  });

  it('has the expected layout, formats and quoted references to บันทึก', () => {
    const s = wb.sheets[1];
    expect(s.name).toBe('สรุป');
    expect(s.frozenRows).toBe(1);
    expect(createLedgerWorkbook().sheets[1].cells['B1'].raw).toBe('=YEAR(TODAY())');
    expect(s.cells['A3'].raw).toBe('หมวด');
    expect(s.cells['A3'].style).toMatchObject({ bold: true, fill: '#e2e8f0' });
    expect(SUMMARY_CATEGORIES.map((_, i) => s.cells[`A${4 + i}`].raw)).toEqual(['อาหาร', 'เดินทาง', 'ที่พัก', 'อื่นๆ']);
    expect(s.cells['B4'].raw).toBe(
      "=IF(A4=\"\",\"\",SUMIFS('บันทึก'!$D$2:$D$1000,'บันทึก'!$C$2:$C$1000,A4))"
    );
    expect(s.cells['B4'].fmt).toBe('baht');
    expect(L.categoryLastRow - L.categoryFirstRow + 1).toBe(10); // 4 หมวด + 6 แถวว่าง
    expect(THAI_MONTHS.map((_, i) => s.cells[`B${monthRow(i + 1)}`].raw)).toEqual(THAI_MONTHS);
    expect(s.cells[`C${monthRow(1)}`].raw).toBe(
      "=SUMIFS('บันทึก'!$D$2:$D$1000,'บันทึก'!$A$2:$A$1000,\">=\"&DATE($B$1,A18,1),'บันทึก'!$A$2:$A$1000,\"<\"&DATE($B$1,A18+1,1))"
    );
    expect(s.cells[`B${L.monthTotalRow}`].raw).toBe('รวมทั้งปี');
  });

  it('category table matches independently computed totals (all years, blank category not counted)', () => {
    const ev = new Evaluator(wb);
    for (const [i, cat] of SUMMARY_CATEGORIES.entries()) {
      const r = L.categoryFirstRow + i;
      expect(v(ev, `B${r}`)).toBeCloseTo(sumWhere((e) => e.cat === cat, 'income'));
      expect(v(ev, `C${r}`)).toBeCloseTo(sumWhere((e) => e.cat === cat, 'expense'));
    }
    // แถวว่างที่เตรียมไว้: ผลลัพธ์ว่าง ไม่รวมแถวที่หมวดว่าง
    expect(v(ev, 'B8')).toBe('');
    expect(v(ev, 'C8')).toBe('');
    const known = (e: Entry) => SUMMARY_CATEGORIES.includes(e.cat);
    expect(v(ev, `C${L.categoryTotalRow}`)).toBeCloseTo(sumWhere(known, 'expense'));
  });

  it('a category the user types into a spare row gets its totals', () => {
    wb.sheets[1].cells['A8'] = { raw: 'ของใช้' };
    const ev = new Evaluator(wb);
    expect(v(ev, 'C8')).toBe(250);
  });

  it('monthly table matches independent totals; first and last days land in the right month', () => {
    const ev = new Evaluator(wb);
    for (let m = 1; m <= 12; m++) {
      const inMonth = (e: Entry) => e.y === 2026 && e.m === m;
      const income = sumWhere(inMonth, 'income');
      const expense = sumWhere(inMonth, 'expense');
      expect(v(ev, `C${monthRow(m)}`)).toBeCloseTo(income);
      expect(v(ev, `D${monthRow(m)}`)).toBeCloseTo(expense);
      expect(v(ev, `E${monthRow(m)}`)).toBeCloseTo(income - expense);
    }
    expect(v(ev, `D${monthRow(1)}`)).toBe(165); // 1 ม.ค. + 31 ม.ค. (พิมพ์เป็นข้อความ)
    expect(v(ev, `C${monthRow(12)}`)).toBe(500); // 31 ธ.ค.
    const in2026 = (e: Entry) => e.y === 2026;
    expect(v(ev, `C${L.monthTotalRow}`)).toBeCloseTo(sumWhere(in2026, 'income'));
    expect(v(ev, `D${L.monthTotalRow}`)).toBeCloseTo(sumWhere(in2026, 'expense'));
  });

  it('changing the year in B1 recalculates the monthly table', () => {
    const ev = new Evaluator(wb);
    expect(v(ev, `D${monthRow(12)}`)).toBe(0);
    setYear(wb, 2025);
    ev.invalidate('B1', 1);
    expect(v(ev, `D${monthRow(12)}`)).toBe(999);
    expect(v(ev, `D${monthRow(1)}`)).toBe(0);
    setYear(wb, 2027);
    ev.invalidate('B1', 1);
    expect(v(ev, `D${monthRow(1)}`)).toBe(777);
  });

  it('an empty ledger shows zeros, never errors', () => {
    wb = createLedgerWorkbook(new Date(2026, 9, 16));
    fillLedger(wb, []); // ลบแถวตัวอย่าง ชีตบันทึกว่าง
    setYear(wb, 2026);
    const ev = new Evaluator(wb);
    const values = Object.keys(wb.sheets[1].cells)
      .filter((k) => wb.sheets[1].cells[k].raw.startsWith('='))
      .map((k) => v(ev, k));
    expect(values.filter((x) => !(x === 0 || x === "" || x === 2026))).toEqual([]);
    expect(v(ev, `E${L.monthTotalRow}`)).toBe(0);
  });

  it('renaming the บันทึก sheet keeps the summary working', () => {
    expect(renameSheet(wb, 0, 'รายการ ปี 69').ok).toBe(true);
    expect(wb.sheets[1].cells['B4'].raw).toContain("'รายการ ปี 69'!$D$2:$D$1000");
    expect(wb.sheets[1].cells[`C${monthRow(2)}`].raw).not.toContain('บันทึก');
    const ev = new Evaluator(wb);
    expect(v(ev, 'C6')).toBe(3000);
    expect(v(ev, `C${monthRow(2)}`)).toBe(30000);
  });

  it('round-trips through .xlsx with the same results', async () => {
    const reloaded = await importFromXlsxBuffer(await exportToXlsxBuffer(wb));
    expect(reloaded.sheets.map((s) => s.name)).toEqual(['บันทึก', 'สรุป']);
    expect(reloaded.sheets[1].cells['B4'].raw).toBe(wb.sheets[1].cells['B4'].raw);
    const ev = new Evaluator(reloaded);
    expect(v(ev, `C${monthRow(2)}`)).toBe(30000);
    expect(v(ev, 'C4')).toBeCloseTo(999 + 120 + 80.5 + 777);
    expect(reloaded.sheets[1].frozenRows).toBe(1);
  });
});

describe('เทมเพลตรายรับรายจ่ายเมื่อเปิดใช้', () => {
  it('sample data is in the current month and shows up in the summary straight away', () => {
    const today = new Date();
    const wb = createLedgerWorkbook(today);
    const ev = new Evaluator(wb);
    expect(v(ev, 'B1')).toBe(today.getFullYear());
    const r = monthRow(today.getMonth() + 1);
    expect(v(ev, `C${r}`)).toBe(25000);
    expect(v(ev, `D${r}`)).toBe(230);
    expect(v(ev, 'C4')).toBe(150); // อาหาร
    expect(v(ev, 'B7')).toBe(25000); // อื่นๆ (เงินเดือน)
  });

  it('renaming บันทึก in the app updates summary formulas (and undo restores them)', () => {
    const S = () => useSheetStore.getState();
    S().newWorkbook();
    S().loadTemplate(createLedgerWorkbook(), 'ตารางรายรับรายจ่าย');
    const before = S().workbook.sheets[1].cells['B4'].raw;
    expect(S().renameSheet(0, 'Ledger').success).toBe(true);
    expect(S().workbook.sheets[1].cells['B4'].raw).toBe(
      '=IF(A4="","",SUMIFS(Ledger!$D$2:$D$1000,Ledger!$C$2:$C$1000,A4))'
    );
    S().switchSheet(1);
    expect(S().getComputedValue(3, 2)).toBe(150);
    S().undo();
    expect(S().workbook.sheets[1].cells['B4'].raw).toBe(before);
  });
});
