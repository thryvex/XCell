import { describe, it, expect } from 'vitest';
import {
  getChartData,
  hasNumericData,
  defaultChartAnchor,
  shiftChartsOnInsert,
  shiftChartsOnDelete,
  anchorToPx,
  pxToAnchor,
  chartToTwoCellAnchor,
  twoCellAnchorToChart,
  niceTicks,
  MAX_CHART_SERIES,
  EMU_PER_PX,
} from '../../src/core/chart';
import { createEmptySheet, SheetChart, CellValue } from '../../src/core/model';

function grid(rows: CellValue[][]): (key: string) => CellValue | undefined {
  const map: Record<string, CellValue> = {};
  rows.forEach((row, r) =>
    row.forEach((v, c) => {
      map[String.fromCharCode(65 + c) + (r + 1)] = v;
    })
  );
  return (k) => map[k];
}
const label = (get: (k: string) => CellValue | undefined) => (k: string) => String(get(k) ?? '');

function chart(range: SheetChart['range'], anchor = { row: 0, col: 5, dx: 0, dy: 0 }): SheetChart {
  return { id: 'x', type: 'bar', title: '', range, anchor, width: 480, height: 300 };
}

describe('getChartData', () => {
  const get = grid([
    ['หมวด', 'ม.ค.', 'ก.พ.'],
    ['อาหาร', 100, 120],
    ['เดินทาง', 50, '#DIV/0!'],
    ['ที่พัก', '', 300],
  ]);

  it('first column = categories, header row = series names, non-numbers = no value', () => {
    const d = getChartData({ r1: 0, c1: 0, r2: 3, c2: 2 }, get, label(get));
    expect(d.hasHeader).toBe(true);
    expect(d.hasCategoryColumn).toBe(true);
    expect(d.categories).toEqual(['อาหาร', 'เดินทาง', 'ที่พัก']);
    expect(d.series.map((s) => s.name)).toEqual(['ม.ค.', 'ก.พ.']);
    expect(d.series[0].values).toEqual([100, 50, null]);
    expect(d.series[1].values).toEqual([120, null, 300]);
    expect(hasNumericData(d)).toBe(true);
  });

  it('without a header row, series get default names', () => {
    const d = getChartData({ r1: 1, c1: 0, r2: 3, c2: 1 }, get, label(get));
    expect(d.hasHeader).toBe(false);
    expect(d.series[0].name).toBe('ชุดข้อมูล 1');
    expect(d.series[0].values).toEqual([100, 50, null]);
  });

  it('a single column is one series with 1, 2, 3… as categories', () => {
    const d = getChartData({ r1: 0, c1: 1, r2: 3, c2: 1 }, get, label(get));
    expect(d.hasCategoryColumn).toBe(false);
    expect(d.hasHeader).toBe(true);
    expect(d.categories).toEqual(['1', '2', '3']);
    expect(d.series.map((s) => s.values)).toEqual([[100, 50, null]]);
  });

  it('skips rows hidden by a filter', () => {
    const d = getChartData({ r1: 0, c1: 0, r2: 3, c2: 1 }, get, label(get), new Set([2]));
    expect(d.categories).toEqual(['อาหาร', 'ที่พัก']);
    expect(d.rows).toEqual([1, 3]);
  });

  it('uses at most 6 series', () => {
    const wide = grid([['หมวด', ...Array.from({ length: 9 }, (_, i) => `s${i}`)], ['a', 1, 2, 3, 4, 5, 6, 7, 8, 9]]);
    const d = getChartData({ r1: 0, c1: 0, r2: 1, c2: 9 }, wide, label(wide));
    expect(d.series).toHaveLength(MAX_CHART_SERIES);
  });

  it('detects ranges with no numbers', () => {
    const text = grid([['a', 'b'], ['c', 'd']]);
    expect(hasNumericData(getChartData({ r1: 0, c1: 0, r2: 1, c2: 1 }, text, label(text)))).toBe(false);
    expect(hasNumericData(getChartData(null, text, label(text)))).toBe(false);
  });

  it('places a new chart one column right of the data', () => {
    expect(defaultChartAnchor({ r1: 2, c1: 0, r2: 8, c2: 3 })).toEqual({ row: 2, col: 5, dx: 0, dy: 0 });
  });
});

describe('charts follow inserted/deleted rows and columns', () => {
  it('insert above/inside the range', () => {
    const s = createEmptySheet('S');
    s.charts = [chart({ r1: 2, c1: 0, r2: 5, c2: 1 }, { row: 2, col: 3, dx: 4, dy: 5 })];
    shiftChartsOnInsert(s, 'row', 0, 2); // เหนือช่วง
    expect(s.charts[0].range).toEqual({ r1: 4, c1: 0, r2: 7, c2: 1 });
    expect(s.charts[0].anchor.row).toBe(4);
    shiftChartsOnInsert(s, 'row', 5, 1); // กลางช่วง: ขยาย
    expect(s.charts[0].range).toEqual({ r1: 4, c1: 0, r2: 8, c2: 1 });
    shiftChartsOnInsert(s, 'col', 0, 1);
    expect(s.charts[0].range).toEqual({ r1: 4, c1: 1, r2: 8, c2: 2 });
    expect(s.charts[0].anchor.col).toBe(4);
  });

  it('delete shrinks the range, and deleting all of it leaves no data', () => {
    const s = createEmptySheet('S');
    s.charts = [chart({ r1: 2, c1: 0, r2: 5, c2: 1 }, { row: 3, col: 3, dx: 4, dy: 5 })];
    shiftChartsOnDelete(s, 'row', 0, 1);
    expect(s.charts[0].range).toEqual({ r1: 1, c1: 0, r2: 4, c2: 1 });
    shiftChartsOnDelete(s, 'row', 3, 3); // ลบท้ายช่วงเกินขอบ
    expect(s.charts[0].range).toEqual({ r1: 1, c1: 0, r2: 2, c2: 1 });
    // แถวที่ยึดกราฟถูกลบ: ยึดกับแถวที่มาแทน
    expect(s.charts[0].anchor).toEqual({ row: 2, col: 3, dx: 4, dy: 5 });
    shiftChartsOnDelete(s, 'row', 2, 1);
    expect(s.charts[0].anchor).toEqual({ row: 2, col: 3, dx: 4, dy: 0 });
    shiftChartsOnDelete(s, 'col', 0, 2);
    expect(s.charts[0].range).toBeNull();
    expect(s.charts[0].anchor.col).toBe(1);
  });
});

describe('chart position', () => {
  const s = createEmptySheet('S');
  s.colWidths = { 0: 100, 2: 50 };
  s.rowHeights = { 1: 40 };

  it('anchor <-> px round-trips, hidden rows take no space', () => {
    const a = { row: 3, col: 3, dx: 7, dy: 9 };
    const p = anchorToPx(s, a);
    expect(p).toEqual({ x: 100 + 110 + 50 + 7, y: 28 + 40 + 28 + 9 });
    expect(pxToAnchor(s, p.x, p.y)).toEqual(a);
    expect(anchorToPx(s, a, new Set([1])).y).toBe(28 + 28 + 9);
  });

  it('two-cell anchor in the file matches the chart size', () => {
    const c = chart({ r1: 0, c1: 0, r2: 1, c2: 1 }, { row: 1, col: 1, dx: 10, dy: 0 });
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
  it('includes zero and uses round steps', () => {
    const t = niceTicks(0, 8000);
    expect(t.min).toBe(0);
    expect(t.max).toBeGreaterThanOrEqual(8000);
    expect([1000, 2000, 2500, 5000]).toContain(t.step);
    expect(niceTicks(-30, 45).ticks).toContain(0);
    expect(niceTicks(5, 5).max).toBeGreaterThanOrEqual(5);
    expect(niceTicks(0, 0).max).toBeGreaterThan(0);
  });
});
