import { describe, it, expect, afterEach } from 'vitest';
import {
  buildChartFromRange,
  resolveChartData,
  workbookChartContext,
  newChart,
  makeChartRef,
  parseChartRef,
  isLostRef,
  shiftChartsOnInsert,
  shiftChartsOnDelete,
  snapshotCharts,
  restoreCharts,
  anchorToPx,
  pxToAnchor,
  chartToTwoCellAnchor,
  twoCellAnchorToChart,
  niceTicks,
  paletteColor,
  CHART_PALETTE,
  MAX_CHART_SERIES,
  MAX_CHART_POINTS,
  EMU_PER_PX,
  countChartsWithLostData,
} from '../../src/core/chart';
import { renameSheet, deleteSheet, duplicateSheet } from '../../src/core/workbookOps';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet, Sheet, SheetChart, Workbook, Cell } from '../../src/core/model';
import { setDisplayYearMode, dateToSerial } from '../../src/core/numfmt';

function sheetOf(name: string, cells: Record<string, Cell | string>): Sheet {
  const s = createEmptySheet(name);
  for (const [k, v] of Object.entries(cells)) s.cells[k] = typeof v === 'string' ? { raw: v } : v;
  return s;
}

function ctxOf(wb: Workbook, forFile = false) {
  return workbookChartContext(wb, new Evaluator(wb), forFile);
}

/** ข้อมูลตัวอย่าง: หัวตาราง + ป้ายข้อความ + 2 ชุด (มี error, ช่องว่าง, ข้อความ) */
function dataSheet(name = 'ข้อมูล'): Sheet {
  return sheetOf(name, {
    A1: 'หมวด', B1: 'ม.ค.', C1: 'ก.พ.',
    A2: 'อาหาร', B2: '100', C2: '120',
    A3: 'เดินทาง', B3: '=B2/2', C3: '=1/0',
    A4: 'ที่พัก', C4: 'ไม่ทราบ',
  });
}

const R = (r1: number, c1: number, r2: number, c2: number) => ({ r1, c1, r2, c2 });

function built(wb: Workbook, sheet: string, range: ReturnType<typeof R>, opts = {}) {
  const res = buildChartFromRange(ctxOf(wb), sheet, range, opts);
  if (!res.ok) throw new Error(res.error);
  return res;
}

afterEach(() => setDisplayYearMode('ce'));

// ---------------------------------------------------------------------------

describe('reference ของกราฟ', () => {
  it('makeChartRef ใช้กติกา 4B และ parseChartRef อ่านกลับได้', () => {
    expect(makeChartRef('Data', 1, 1, 4, 1)).toBe('Data!B2:B5');
    expect(makeChartRef('ข้อมูล ปี 69', 0, 0)).toBe("'ข้อมูล ปี 69'!A1");
    expect(makeChartRef("ร้าน'ป้า", 0, 0, 0, 2)).toBe("'ร้าน''ป้า'!A1:C1");
    expect(parseChartRef("'ร้าน''ป้า'!A1:C1")).toEqual({ sheet: "ร้าน'ป้า", r1: 0, c1: 0, r2: 0, c2: 2 });
    expect(parseChartRef('$C$5:A1', 'S')).toEqual({ sheet: 'S', r1: 0, c1: 0, r2: 4, c2: 2 });
    expect(parseChartRef('A1:B2')).toBeNull(); // ไม่มีชื่อชีตและไม่มีค่าเริ่มต้น
    expect(parseChartRef('A1:B2,C1')).toBeNull();
    expect(isLostRef('#REF!')).toBe(true);
    expect(isLostRef('Data!A1')).toBe(false);
  });
});

describe('สร้างกราฟจากช่วงที่เลือก', () => {
  const wb = (): Workbook => ({ sheets: [dataSheet()], activeSheet: 0 });

  it('ตรวจแถวหัวข้อ + คอลัมน์ป้ายชื่อ และสร้างชุดจากแต่ละคอลัมน์ตัวเลข', () => {
    const r = built(wb(), 'ข้อมูล', R(0, 0, 3, 2));
    expect(r.headerRow).toBe(true);
    expect(r.labelCol).toBe(true);
    expect(r.byRow).toBe(false);
    expect(r.suggestedType).toBe('bar');
    expect(r.series).toEqual([
      { name: { ref: "'ข้อมูล'!B1" }, categories: "'ข้อมูล'!A2:A4", values: "'ข้อมูล'!B2:B4", color: CHART_PALETTE[0] },
      { name: { ref: "'ข้อมูล'!C1" }, categories: "'ข้อมูล'!A2:A4", values: "'ข้อมูล'!C2:C4", color: CHART_PALETTE[1] },
    ]);
  });

  it('สลับแถว/คอลัมน์: ชุดจากแต่ละแถว ชื่อจากคอลัมน์แรก ป้ายจากแถวแรก', () => {
    const r = built(wb(), 'ข้อมูล', R(0, 0, 3, 2), { byRow: true });
    // แถว "ที่พัก" ไม่มีตัวเลข (ว่าง + ข้อความ) จึงไม่เป็นชุดข้อมูล
    expect(r.series.map((s) => s.values)).toEqual(["'ข้อมูล'!B2:C2", "'ข้อมูล'!B3:C3"]);
    expect(r.series[0].name).toEqual({ ref: "'ข้อมูล'!A2" });
    expect(r.series[0].categories).toBe("'ข้อมูล'!B1:C1");
  });

  it('ตัวเลือกบังคับ: ไม่ใช้แถวแรกเป็นหัวข้อ / ไม่ใช้คอลัมน์แรกเป็นป้าย', () => {
    const r = built(wb(), 'ข้อมูล', R(0, 0, 3, 2), { headerRow: false, labelCol: false });
    expect(r.series[0]).toMatchObject({ name: { text: '' }, values: "'ข้อมูล'!B1:B4" });
    expect(r.series[0].categories).toBeUndefined();
    // คอลัมน์ A ไม่มีตัวเลข จึงไม่เป็นชุดข้อมูล
    expect(r.series).toHaveLength(2);
  });

  it('แถวแรกไม่ใช่หัวข้อเมื่อแถวถัดไปไม่มีตัวเลข; คอลัมน์ตัวเลขล้วนไม่ใช่ป้าย', () => {
    const s = sheetOf('S', { A1: 'ก', B1: 'ข', A2: 'ค', B2: 'ง', A3: '1', B3: '5' });
    const r = built({ sheets: [s], activeSheet: 0 }, 'S', R(0, 0, 2, 1));
    expect(r.headerRow).toBe(false);
    const n = sheetOf('N', { A1: 'ลำดับ', B1: 'ยอด', A2: '1', B2: '10', A3: '2', B3: '20' });
    const r2 = built({ sheets: [n], activeSheet: 0 }, 'N', R(0, 0, 2, 1));
    expect(r2.headerRow).toBe(true);
    expect(r2.labelCol).toBe(false);
    expect(r2.series.map((x) => x.values)).toEqual(["N!A2:A3", 'N!B2:B3']);
  });

  it('ชนิดแนะนำ: ป้ายวันที่ → เส้น, ชุดเดียว ≤ 8 จุด → วงกลม, 9 จุด → แท่ง', () => {
    const d = createEmptySheet('D');
    d.cells.A1 = { raw: 'วันที่' };
    d.cells.B1 = { raw: 'ยอด' };
    for (let i = 0; i < 3; i++) {
      d.cells[`A${i + 2}`] = { raw: String(dateToSerial(new Date(2026, 0, i + 1))), fmt: 'date' };
      d.cells[`B${i + 2}`] = { raw: String(i * 10) };
    }
    const dr = built({ sheets: [d], activeSheet: 0 }, 'D', R(0, 0, 3, 1));
    expect(dr.labelCol).toBe(true);
    expect(dr.suggestedType).toBe('line');

    const p = createEmptySheet('P');
    for (let i = 0; i < 9; i++) {
      p.cells[`A${i + 1}`] = { raw: `หมวด${i}` };
      p.cells[`B${i + 1}`] = { raw: String(i + 1) };
    }
    const ctx = { sheets: [p], activeSheet: 0 };
    expect(built(ctx, 'P', R(0, 0, 7, 1)).suggestedType).toBe('pie');
    expect(built(ctx, 'P', R(0, 0, 8, 1)).suggestedType).toBe('bar');
  });

  it('ช่วงไม่มีตัวเลข และเกิน 10 ชุด แจ้งข้อผิดพลาด', () => {
    const t = sheetOf('T', { A1: 'ก', B1: 'ข', A2: 'ค', B2: 'ง' });
    expect(buildChartFromRange(ctxOf({ sheets: [t], activeSheet: 0 }), 'T', R(0, 0, 1, 1))).toEqual({ ok: false, error: 'noNumbers' });
    const w = createEmptySheet('W');
    for (let c = 0; c <= MAX_CHART_SERIES; c++) w.cells[`${String.fromCharCode(66 + c)}1`] = { raw: String(c) };
    w.cells.A1 = { raw: 'ป้าย' };
    expect(buildChartFromRange(ctxOf({ sheets: [w], activeSheet: 0 }), 'W', R(0, 0, 0, MAX_CHART_SERIES + 1))).toEqual({
      ok: false,
      error: 'tooManySeries',
    });
  });
});

describe('ค่าที่ใช้วาดกราฟ', () => {
  function chartOn(wb: Workbook, opts = {}): SheetChart {
    const r = built(wb, 'ข้อมูล', R(0, 0, 3, 2), opts);
    return newChart('bar', r.series, { row: 0, col: 5, dx: 0, dy: 0 });
  }

  it('ข้อความ = 0, ช่องว่าง = ไม่มีค่า, error = ไม่มีค่า + ธง; ค่าจากสูตร', () => {
    const wb: Workbook = { sheets: [dataSheet()], activeSheet: 0 };
    const d = resolveChartData(chartOn(wb), ctxOf(wb));
    expect(d.status).toBe('ok');
    expect(d.categories).toEqual(['อาหาร', 'เดินทาง', 'ที่พัก']);
    expect(d.series.map((s) => s.name)).toEqual(['ม.ค.', 'ก.พ.']);
    expect(d.series[0].values).toEqual([100, 50, null]);
    expect(d.series[1].values).toEqual([120, null, 0]);
    expect(d.series[1].hasError).toBe(true);
    expect(d.hasError).toBe(true);
  });

  it('ข้ามแถวที่ซ่อนด้วยตัวกรอง (ไฟล์ไม่ข้าม) และชุดตามแถวที่ซ่อนไม่แสดง', () => {
    const s = dataSheet();
    s.hiddenRows = [2];
    const wb: Workbook = { sheets: [s], activeSheet: 0 };
    const c = chartOn(wb);
    expect(resolveChartData(c, ctxOf(wb)).categories).toEqual(['อาหาร', 'ที่พัก']);
    expect(resolveChartData(c, ctxOf(wb), false).categories).toHaveLength(3);
    const byRow = chartOn(wb, { byRow: true });
    expect(resolveChartData(byRow, ctxOf(wb)).series.map((x) => x.name)).toEqual(['อาหาร']);
  });

  it('วงกลมใช้ชุดแรกชุดเดียว', () => {
    const wb: Workbook = { sheets: [dataSheet()], activeSheet: 0 };
    const c = { ...chartOn(wb), type: 'pie' as const };
    expect(resolveChartData(c, ctxOf(wb)).series).toHaveLength(1);
  });

  it('ไม่มีป้าย = 1, 2, 3; ชื่อชุดข้อความว่าง = "ชุดข้อมูล N"', () => {
    const wb: Workbook = { sheets: [dataSheet()], activeSheet: 0 };
    const c = newChart('line', [{ name: { text: '' }, values: "'ข้อมูล'!B2:B4", color: '#000000' }], { row: 0, col: 0, dx: 0, dy: 0 });
    const d = resolveChartData(c, ctxOf(wb));
    expect(d.categories).toEqual(['1', '2', '3']);
    expect(d.series[0].name).toBe('ชุดข้อมูล 1');
  });

  it('สถานะ: ไม่มีข้อมูลที่จะแสดง / ข้อมูลหาย / เกินขีดจำกัด', () => {
    const wb: Workbook = { sheets: [dataSheet()], activeSheet: 0 };
    const mk = (values: string) => newChart('bar', [{ name: { text: 'x' }, values, color: '#000000' }], { row: 0, col: 0, dx: 0, dy: 0 });
    expect(resolveChartData(mk("'ข้อมูล'!E1:E5"), ctxOf(wb)).status).toBe('empty');
    expect(resolveChartData(mk('#REF!'), ctxOf(wb)).status).toBe('lost');
    expect(resolveChartData(mk("'ไม่มีชีตนี้'!A1:A3"), ctxOf(wb)).status).toBe('lost');
    expect(resolveChartData(mk(`'ข้อมูล'!A1:A${MAX_CHART_POINTS + 1}`), ctxOf(wb)).status).toBe('tooManyPoints');
    const many = mk("'ข้อมูล'!B2:B4");
    many.series = Array.from({ length: MAX_CHART_SERIES + 1 }, () => ({ ...many.series[0] }));
    expect(resolveChartData(many, ctxOf(wb)).status).toBe('tooManySeries');
  });

  it('อ้างข้อมูลข้ามชีต และป้ายวันที่ตามโหมดปีในแอป แต่เป็น ค.ศ. ในไฟล์', () => {
    const d = createEmptySheet('ข้อมูล ปี 69');
    d.cells.A1 = { raw: String(dateToSerial(new Date(2026, 0, 15))), fmt: 'date' };
    d.cells.B1 = { raw: '5', fmt: 'baht' };
    const host = createEmptySheet('สรุป');
    const wb: Workbook = { sheets: [host, d], activeSheet: 0 };
    const c = newChart('line', [{ name: { text: 'ยอด' }, categories: "'ข้อมูล ปี 69'!A1", values: "'ข้อมูล ปี 69'!B1", color: '#000000' }], {
      row: 0, col: 0, dx: 0, dy: 0,
    });
    setDisplayYearMode('be');
    const app = resolveChartData(c, ctxOf(wb));
    expect(app.categories[0]).toContain('2569');
    expect(app.series[0].values).toEqual([5]);
    expect(app.numFmt).toBe('baht');
    expect(resolveChartData(c, ctxOf(wb, true)).categories[0]).toContain('2026');
  });

  it('จานสี 8 สีวนซ้ำ', () => {
    expect(CHART_PALETTE).toHaveLength(8);
    expect(new Set(CHART_PALETTE).size).toBe(8);
    expect(paletteColor(8)).toBe(CHART_PALETTE[0]);
  });
});

// ---------------------------------------------------------------------------
// refshift ของกราฟ
// ---------------------------------------------------------------------------

/** สมุดงาน: ชีต "ข้อมูล" มีกราฟอ้างตัวเอง, ชีต "สรุป" มีกราฟอ้าง "ข้อมูล" */
function refWorkbook(): Workbook {
  const data = dataSheet();
  const own: SheetChart = newChart(
    'bar',
    [{ name: { ref: "'ข้อมูล'!B2" }, categories: "'ข้อมูล'!A3:A6", values: "'ข้อมูล'!B3:B6", color: '#111111' }],
    { row: 4, col: 5, dx: 3, dy: 7 }
  );
  data.charts = [own];
  const summary = createEmptySheet('สรุป');
  summary.charts = [
    newChart('line', [{ name: { text: 'ยอด' }, categories: "'ข้อมูล'!A3:A6", values: "'ข้อมูล'!B3:B6", color: '#222222' }], {
      row: 2, col: 1, dx: 0, dy: 0,
    }),
  ];
  return { sheets: [data, summary], activeSheet: 0 };
}

const own = (wb: Workbook) => wb.sheets.find((s) => s.name.startsWith('ข้อมูล') || s.name === 'Data')!.charts![0];
const other = (wb: Workbook) => wb.sheets.find((s) => s.name === 'สรุป')!.charts![0];
const vals = (c: SheetChart) => c.series[0].values;

describe('refshift ของกราฟ: แทรก/ลบแถว (ช่วงค่า B3:B6 = แถว 2–5)', () => {
  const cases: Array<[string, (wb: Workbook) => void, string]> = [
    ['แทรกแถวก่อนช่วง', (wb) => shiftChartsOnInsert(wb, 0, 'row', 0, 1), "'ข้อมูล'!B4:B7"],
    ['แทรกที่แถวแรกของช่วง (ช่วงเลื่อนลง)', (wb) => shiftChartsOnInsert(wb, 0, 'row', 2, 1), "'ข้อมูล'!B4:B7"],
    ['แทรกในช่วง (ขยาย)', (wb) => shiftChartsOnInsert(wb, 0, 'row', 4, 1), "'ข้อมูล'!B3:B7"],
    ['แทรก 3 แถวในช่วง', (wb) => shiftChartsOnInsert(wb, 0, 'row', 3, 3), "'ข้อมูล'!B3:B9"],
    ['แทรกถัดจากแถวสุดท้าย (ไม่เปลี่ยน)', (wb) => shiftChartsOnInsert(wb, 0, 'row', 6, 1), "'ข้อมูล'!B3:B6"],
    ['แทรกหลังช่วง', (wb) => shiftChartsOnInsert(wb, 0, 'row', 20, 2), "'ข้อมูล'!B3:B6"],
    ['ลบแถวก่อนช่วง', (wb) => shiftChartsOnDelete(wb, 0, 'row', 0, 1), "'ข้อมูล'!B2:B5"],
    ['ลบแถวในช่วง', (wb) => shiftChartsOnDelete(wb, 0, 'row', 3, 1), "'ข้อมูล'!B3:B5"],
    ['ลบแถวแรกของช่วง', (wb) => shiftChartsOnDelete(wb, 0, 'row', 2, 1), "'ข้อมูล'!B3:B5"],
    ['ลบแถวสุดท้ายของช่วง', (wb) => shiftChartsOnDelete(wb, 0, 'row', 5, 1), "'ข้อมูล'!B3:B5"],
    ['ลบคร่อมขอบบน', (wb) => shiftChartsOnDelete(wb, 0, 'row', 1, 2), "'ข้อมูล'!B2:B4"],
    ['ลบคร่อมขอบล่าง', (wb) => shiftChartsOnDelete(wb, 0, 'row', 4, 4), "'ข้อมูล'!B3:B4"],
    ['ลบหลังช่วง', (wb) => shiftChartsOnDelete(wb, 0, 'row', 10, 3), "'ข้อมูล'!B3:B6"],
    ['ลบทั้งช่วง = reference หาย', (wb) => shiftChartsOnDelete(wb, 0, 'row', 2, 4), '#REF!'],
    ['ลบเกินทั้งช่วง = reference หาย', (wb) => shiftChartsOnDelete(wb, 0, 'row', 0, 10), '#REF!'],
  ];
  for (const [name, op, expected] of cases) {
    it(name, () => {
      const wb = refWorkbook();
      op(wb);
      expect(vals(own(wb))).toBe(expected);
      // กราฟบนชีตอื่นที่อ้างชีตเดียวกันเปลี่ยนเหมือนกัน
      expect(vals(other(wb))).toBe(expected);
    });
  }
});

describe('refshift ของกราฟ: แทรก/ลบคอลัมน์ และชีตอื่น', () => {
  it('แทรกคอลัมน์ก่อนช่วง: ทั้งป้ายและค่าเลื่อน', () => {
    const wb = refWorkbook();
    shiftChartsOnInsert(wb, 0, 'col', 0, 1);
    expect(own(wb).series[0]).toMatchObject({ categories: "'ข้อมูล'!B3:B6", values: "'ข้อมูล'!C3:C6", name: { ref: "'ข้อมูล'!C2" } });
  });

  it('แทรกคอลัมน์ระหว่างป้ายกับค่า', () => {
    const wb = refWorkbook();
    shiftChartsOnInsert(wb, 0, 'col', 1, 1);
    expect(own(wb).series[0]).toMatchObject({ categories: "'ข้อมูล'!A3:A6", values: "'ข้อมูล'!C3:C6" });
  });

  it('แทรกคอลัมน์หลังช่วงไม่เปลี่ยน', () => {
    const wb = refWorkbook();
    shiftChartsOnInsert(wb, 0, 'col', 3, 1);
    expect(vals(own(wb))).toBe("'ข้อมูล'!B3:B6");
  });

  it('ลบคอลัมน์ป้าย: ป้ายหาย (ใช้ 1, 2, 3) แต่กราฟยังแสดงค่า', () => {
    const wb = refWorkbook();
    shiftChartsOnDelete(wb, 0, 'col', 0, 1);
    expect(own(wb).series[0]).toMatchObject({ categories: '#REF!', values: "'ข้อมูล'!A3:A6" });
    const d = resolveChartData(own(wb), ctxOf(wb));
    expect(d.status).toBe('ok');
    expect(d.categories[0]).toBe('1');
  });

  it('ลบคอลัมน์ค่า: ข้อมูลหาย แต่กราฟไม่ถูกลบ', () => {
    const wb = refWorkbook();
    shiftChartsOnDelete(wb, 0, 'col', 1, 1);
    expect(vals(own(wb))).toBe('#REF!');
    expect(wb.sheets[0].charts).toHaveLength(1);
    expect(resolveChartData(own(wb), ctxOf(wb)).status).toBe('lost');
  });

  it('แทรก/ลบแถวในชีตอื่นไม่กระทบ reference ที่ชี้ "ข้อมูล"', () => {
    const wb = refWorkbook();
    shiftChartsOnInsert(wb, 1, 'row', 0, 5);
    shiftChartsOnDelete(wb, 1, 'col', 0, 2);
    expect(vals(own(wb))).toBe("'ข้อมูล'!B3:B6");
    expect(vals(other(wb))).toBe("'ข้อมูล'!B3:B6");
  });

  it('ตำแหน่งยึดเลื่อนเฉพาะกราฟบนชีตที่แทรก/ลบ', () => {
    const wb = refWorkbook();
    shiftChartsOnInsert(wb, 0, 'row', 0, 2);
    expect(own(wb).anchor).toEqual({ row: 6, col: 5, dx: 3, dy: 7 });
    expect(other(wb).anchor).toEqual({ row: 2, col: 1, dx: 0, dy: 0 });
    shiftChartsOnInsert(wb, 1, 'col', 0, 1);
    expect(other(wb).anchor.col).toBe(2);
    expect(vals(other(wb))).toBe("'ข้อมูล'!B5:B8");
  });

  it('แทรกใต้ตำแหน่งยึดไม่เลื่อนกราฟ; ขนาดคงที่', () => {
    const wb = refWorkbook();
    shiftChartsOnInsert(wb, 0, 'row', 10, 3);
    shiftChartsOnInsert(wb, 0, 'col', 9, 1);
    expect(own(wb).anchor).toEqual({ row: 4, col: 5, dx: 3, dy: 7 });
    expect([own(wb).width, own(wb).height]).toEqual([480, 300]);
  });

  it('ลบแถว/คอลัมน์ที่ยึดกราฟ: ยึดกับแถว/คอลัมน์ที่มาแทน ระยะเลื่อนเป็น 0', () => {
    const wb = refWorkbook();
    shiftChartsOnDelete(wb, 0, 'row', 3, 2);
    expect(own(wb).anchor).toEqual({ row: 3, col: 5, dx: 3, dy: 0 });
    shiftChartsOnDelete(wb, 0, 'col', 5, 1);
    expect(own(wb).anchor).toEqual({ row: 3, col: 5, dx: 0, dy: 0 });
    shiftChartsOnDelete(wb, 0, 'col', 0, 1);
    expect(own(wb).anchor.col).toBe(4);
  });
});

describe('refshift ของกราฟ: เปลี่ยนชื่อ/ลบ/ทำสำเนาชีต', () => {
  it('เปลี่ยนชื่อเป็นชื่อไทยมีเว้นวรรค', () => {
    const wb = refWorkbook();
    expect(renameSheet(wb, 0, 'ข้อมูล ปี 69').ok).toBe(true);
    expect(vals(own(wb))).toBe("'ข้อมูล ปี 69'!B3:B6");
    expect(own(wb).series[0].name).toEqual({ ref: "'ข้อมูล ปี 69'!B2" });
    expect(vals(other(wb))).toBe("'ข้อมูล ปี 69'!B3:B6");
  });

  it("เปลี่ยนชื่อที่มี ' และชื่ออังกฤษล้วน", () => {
    const wb = refWorkbook();
    renameSheet(wb, 0, "ร้าน'ป้า");
    expect(vals(other(wb))).toBe("'ร้าน''ป้า'!B3:B6");
    renameSheet(wb, 0, 'Data');
    expect(vals(other(wb))).toBe('Data!B3:B6');
  });

  it('เปลี่ยนชื่อชีตที่ไม่เกี่ยวไม่กระทบ', () => {
    const wb = refWorkbook();
    renameSheet(wb, 1, 'หน้าสรุป');
    expect(vals(wb.sheets[1].charts![0])).toBe("'ข้อมูล'!B3:B6");
  });

  it('ลบชีตที่กราฟอ้างถึง: กราฟยังอยู่ แสดง "ข้อมูลหาย"', () => {
    const wb = refWorkbook();
    expect(deleteSheet(wb, 0).ok).toBe(true);
    expect(wb.sheets).toHaveLength(1);
    expect(wb.sheets[0].charts).toHaveLength(1);
    expect(vals(wb.sheets[0].charts![0])).toBe('#REF!');
    expect(resolveChartData(wb.sheets[0].charts![0], ctxOf(wb)).status).toBe('lost');
  });

  it('ลบชีตที่กราฟอยู่: กราฟถูกลบพร้อมชีต ไม่กระทบกราฟชีตอื่น', () => {
    const wb = refWorkbook();
    deleteSheet(wb, 1);
    expect(wb.sheets.map((s) => s.name)).toEqual(['ข้อมูล']);
    expect(vals(own(wb))).toBe("'ข้อมูล'!B3:B6");
  });

  it('ทำสำเนาชีตข้อมูล: กราฟในสำเนาชี้สำเนา id ใหม่ ต้นฉบับไม่เปลี่ยน', () => {
    const wb = refWorkbook();
    const res = duplicateSheet(wb, 0);
    const copy = wb.sheets[res.index!];
    expect(copy.name).toBe('สำเนาของ ข้อมูล');
    const c = copy.charts![0];
    expect(c.id).not.toBe(own(wb).id);
    expect(c.series[0]).toMatchObject({
      name: { ref: "'สำเนาของ ข้อมูล'!B2" },
      categories: "'สำเนาของ ข้อมูล'!A3:A6",
      values: "'สำเนาของ ข้อมูล'!B3:B6",
    });
    expect(vals(own(wb))).toBe("'ข้อมูล'!B3:B6");
  });

  it('ทำสำเนาชีตสรุป: กราฟที่อ้างชีตอื่นยังชี้ชีตเดิม', () => {
    const wb = refWorkbook();
    const res = duplicateSheet(wb, 1);
    expect(vals(wb.sheets[res.index!].charts![0])).toBe("'ข้อมูล'!B3:B6");
  });
});

describe('Undo ของกราฟ (snapshot)', () => {
  it('คืนกราฟทุกชีตเหมือนเดิมหลังแทรก/ลบหลายครั้ง', () => {
    const wb = refWorkbook();
    const before = JSON.parse(JSON.stringify(wb.sheets.map((s) => s.charts)));
    const snap = snapshotCharts(wb);
    shiftChartsOnInsert(wb, 0, 'row', 3, 2);
    shiftChartsOnDelete(wb, 0, 'col', 1, 1);
    renameSheet(wb, 0, 'อื่น');
    restoreCharts(wb, snap);
    expect(wb.sheets.map((s) => s.charts)).toEqual(before);
  });

  it('ชีตที่ไม่มีกราฟใน snapshot ถูกล้างกราฟ', () => {
    const wb = refWorkbook();
    wb.sheets[1].charts = [];
    const snap = snapshotCharts(wb);
    wb.sheets[1].charts = [own(wb)];
    restoreCharts(wb, snap);
    expect(wb.sheets[1].charts).toBeUndefined();
  });
});

// ---------------------------------------------------------------------------

describe('ตำแหน่งกราฟ', () => {
  const s = createEmptySheet('S');
  s.colWidths = { 0: 100, 2: 50 };
  s.rowHeights = { 1: 40 };

  it('anchor <-> px ไปกลับได้ และแถวที่ซ่อนไม่กินที่', () => {
    const a = { row: 3, col: 3, dx: 7, dy: 9 };
    const p = anchorToPx(s, a);
    expect(p).toEqual({ x: 100 + 110 + 50 + 7, y: 28 + 40 + 28 + 9 });
    expect(pxToAnchor(s, p.x, p.y)).toEqual(a);
    expect(anchorToPx(s, a, new Set([1])).y).toBe(28 + 28 + 9);
  });

  it('two-cell anchor ในไฟล์ตรงกับขนาดกราฟเมื่อคอลัมน์กว้างไม่เท่ากัน', () => {
    const c = newChart('bar', [], { row: 1, col: 1, dx: 10, dy: 0 });
    c.width = 200;
    c.height = 130;
    const { from, to } = chartToTwoCellAnchor(s, c);
    expect(from).toEqual({ col: 1, colOff: 10 * EMU_PER_PX, row: 1, rowOff: 0 });
    // 110-10 = 100 ในคอลัมน์ B, 50 ในคอลัมน์ C, เหลือ 50 ในคอลัมน์ D
    // แถว 2 สูง 40, แถว 3–5 สูง 28 → เหลือ 6 ในแถว 6
    expect(to).toEqual({ col: 3, colOff: 50 * EMU_PER_PX, row: 5, rowOff: 6 * EMU_PER_PX });
    expect(twoCellAnchorToChart(s, from, to)).toEqual({ anchor: c.anchor, width: 200, height: 130 });
  });
});

describe('niceTicks', () => {
  it('รวมศูนย์ และใช้ขั้น 1/2/5 × 10^n', () => {
    const t = niceTicks(0, 8000);
    expect(t.min).toBe(0);
    expect(t.max).toBeGreaterThanOrEqual(8000);
    expect([1000, 2000, 5000]).toContain(t.step);
    for (const [lo, hi] of [[0, 7], [0, 0.37], [-30, 45], [12, 9876543], [-5, -1]]) {
      const r = niceTicks(lo, hi);
      const m = r.step / Math.pow(10, Math.floor(Math.log10(r.step) + 1e-9));
      expect([1, 2, 5]).toContain(Math.round(m * 1e6) / 1e6);
      expect(r.min).toBeLessThanOrEqual(Math.min(0, lo));
      expect(r.max).toBeGreaterThanOrEqual(Math.max(0, hi));
      expect(r.ticks).toContain(0);
      expect(r.ticks.length).toBeLessThanOrEqual(12);
    }
    expect(niceTicks(0, 0).max).toBeGreaterThan(0);
    expect(niceTicks(0, 0.3).ticks).toEqual([0, 0.1, 0.2, 0.3]);
  });
});

describe('กราฟที่ข้อมูลหาย (เตือนก่อนบันทึก)', () => {
  it('นับเฉพาะกราฟที่ช่วงค่าหายทุกชุด หรือชีตที่อ้างถูกลบ', () => {
    const wb = refWorkbook();
    expect(countChartsWithLostData(wb)).toBe(0);
    shiftChartsOnDelete(wb, 0, 'col', 0, 1); // ป้ายหาย แต่ค่ายังอยู่
    expect(countChartsWithLostData(wb)).toBe(0);
    shiftChartsOnDelete(wb, 0, 'col', 0, 1); // ค่าหาย
    expect(countChartsWithLostData(wb)).toBe(2);
    const wb2 = refWorkbook();
    deleteSheet(wb2, 0);
    expect(countChartsWithLostData(wb2)).toBe(1);
  });
});
