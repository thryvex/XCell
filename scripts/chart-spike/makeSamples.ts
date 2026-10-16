/**
 * สร้างไฟล์ตัวอย่างกราฟด้วยโค้ดจริงของแอป (src/core) สำหรับเปิดใน Google Sheets / Excel
 * รัน: npx vite-node scripts/chart-spike/makeSamples.ts
 * ผลลัพธ์: scripts/chart-spike/out/sample-*.xlsx และพิมพ์ค่าที่แอปคำนวณไว้เทียบ
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { exportToXlsxBuffer, importFromXlsxBuffer } from '../../src/core/xlsx';
import { buildChartFromRange, newChart, resolveChartData, workbookChartContext } from '../../src/core/chart';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet, Sheet, SheetChart, Workbook, ChartType, NumFmt } from '../../src/core/model';
import { dateToSerial } from '../../src/core/numfmt';

const OUT = join(__dirname, 'out');
mkdirSync(OUT, { recursive: true });

function put(s: Sheet, key: string, raw: string | number, fmt?: NumFmt) {
  s.cells[key] = fmt ? { raw: String(raw), fmt } : { raw: String(raw) };
}

function chartFrom(
  wb: Workbook,
  sheetName: string,
  range: { r1: number; c1: number; r2: number; c2: number },
  type: ChartType,
  anchor: SheetChart['anchor'],
  extra: Partial<SheetChart> = {}
): SheetChart {
  const ctx = workbookChartContext(wb, new Evaluator(wb));
  const res = buildChartFromRange(ctx, sheetName, range);
  if (!res.ok) throw new Error(`${sheetName}: ${res.error}`);
  return { ...newChart(type, res.series, anchor), ...extra };
}

function report(name: string, wb: Workbook) {
  const ctx = workbookChartContext(wb, new Evaluator(wb));
  console.log(`\n=== ${name} ===`);
  for (const s of wb.sheets) {
    for (const c of s.charts ?? []) {
      const d = resolveChartData(c, ctx);
      console.log(`[${s.name}] ${c.type} "${c.title}" สถานะ=${d.status} ป้าย=${JSON.stringify(d.categories)}`);
      for (const ser of d.series) console.log(`    ${ser.name} (${ser.color}): ${JSON.stringify(ser.values)}${ser.hasError ? ' (มี error)' : ''}`);
    }
  }
}

async function save(file: string, wb: Workbook) {
  const bytes = await exportToXlsxBuffer(wb);
  writeFileSync(join(OUT, file), bytes);
  // อ่านกลับด้วยแอปเองเพื่อยืนยันว่าเปิดได้และกราฟครบ
  const back = await importFromXlsxBuffer(bytes);
  const count = (w: Workbook) => w.sheets.reduce((n, s) => n + (s.charts?.length ?? 0), 0);
  console.log(`${file}: เขียนกราฟ ${count(wb)} อัน, แอปอ่านกลับได้ ${count(back)} อัน, แถบเตือน=${JSON.stringify(back.unsupportedFeatures ?? [])}`);
  report(file, wb);
}

// (ก) กราฟแท่ง เส้น วงกลม ข้อมูลภาษาไทย
async function sampleA() {
  const s = createEmptySheet('ข้อมูล');
  s.colWidths = { 0: 120 };
  const rows: Array<[string, number | string, number | string]> = [
    ['หมวด', 'ม.ค.', 'ก.พ.'],
    ['อาหาร', 4500, 4800],
    ['เดินทาง', 1200, '=B3*1.1'],
    ['ที่พัก', 8000, 8000],
    ['อื่นๆ', 650, 'ไม่ทราบ'], // ข้อความในช่องค่า: แอปนับเป็น 0
    ['ของขวัญ', '', 300], // ช่องว่าง: ไม่มีค่า
  ];
  rows.forEach((r, i) => {
    put(s, `A${i + 1}`, r[0]);
    if (r[1] !== '') put(s, `B${i + 1}`, r[1], i > 0 ? 'baht' : undefined);
    put(s, `C${i + 1}`, r[2], i > 0 && typeof r[2] === 'number' ? 'baht' : undefined);
  });
  put(s, 'E1', 'วันที่');
  put(s, 'F1', 'ยอดขาย');
  const sales = [320, 410, '', 380, 520, 610, 450]; // วันที่ 3 ว่าง → เส้นขาดช่วง
  sales.forEach((v, i) => {
    put(s, `E${i + 2}`, dateToSerial(new Date(2026, 0, i + 1)), 'date');
    if (v !== '') put(s, `F${i + 2}`, v);
  });

  const wb: Workbook = { sheets: [s], activeSheet: 0 };
  s.charts = [
    chartFrom(wb, 'ข้อมูล', { r1: 0, c1: 0, r2: 5, c2: 2 }, 'bar', { row: 0, col: 7, dx: 0, dy: 0 }, {
      title: 'รายจ่ายรายเดือน',
      xTitle: 'หมวด',
      yTitle: 'บาท',
      legend: 'bottom',
    }),
    chartFrom(wb, 'ข้อมูล', { r1: 0, c1: 4, r2: 7, c2: 5 }, 'line', { row: 12, col: 7, dx: 0, dy: 0 }, {
      title: 'ยอดขายรายวัน (มกราคม)',
      dataLabels: true,
      legend: 'none',
    }),
    chartFrom(wb, 'ข้อมูล', { r1: 0, c1: 0, r2: 5, c2: 1 }, 'pie', { row: 24, col: 7, dx: 0, dy: 0 }, {
      title: 'สัดส่วนรายจ่าย ม.ค.',
      dataLabels: true,
      legend: 'right',
    }),
  ];
  await save('sample-a-types.xlsx', wb);
}

// (ข) ชีตชื่อไทยมีเว้นวรรค และกราฟอ้างข้อมูลข้ามชีต
async function sampleB() {
  const data = createEmptySheet('ข้อมูล ปี 69');
  const months = ['มกราคม', 'กุมภาพันธ์', 'มีนาคม', 'เมษายน'];
  put(data, 'A1', 'เดือน');
  put(data, 'B1', 'รายรับ');
  put(data, 'C1', 'รายจ่าย');
  months.forEach((m, i) => {
    put(data, `A${i + 2}`, m);
    put(data, `B${i + 2}`, 20000 + i * 1500, 'baht');
    put(data, `C${i + 2}`, 15000 + i * 900, 'baht');
  });
  const shop = createEmptySheet("ร้าน'ป้า");
  put(shop, 'A1', 'สินค้า');
  put(shop, 'B1', 'จำนวน');
  ['กาแฟ', 'ชาเย็น', 'ขนมปัง'].forEach((p, i) => {
    put(shop, `A${i + 2}`, p);
    put(shop, `B${i + 2}`, [42, 35, 18][i]);
  });
  const summary = createEmptySheet('สรุป กราฟ');
  put(summary, 'A1', 'กราฟในชีตนี้อ้างข้อมูลจากชีต "ข้อมูล ปี 69" และ "ร้าน\'ป้า"');

  const wb: Workbook = { sheets: [summary, data, shop], activeSheet: 0 };
  summary.charts = [
    chartFrom(wb, 'ข้อมูล ปี 69', { r1: 0, c1: 0, r2: 4, c2: 2 }, 'bar', { row: 2, col: 0, dx: 0, dy: 0 }, {
      title: 'รายรับ-รายจ่ายรายเดือน',
      yTitle: 'บาท',
    }),
    chartFrom(wb, "ร้าน'ป้า", { r1: 0, c1: 0, r2: 3, c2: 1 }, 'pie', { row: 2, col: 5, dx: 0, dy: 0 }, {
      title: 'ยอดขายร้านป้า',
      dataLabels: true,
    }),
  ];
  await save('sample-b-cross-sheet.xlsx', wb);
}

// (ค) คอลัมน์กว้างไม่เท่ากัน และหลายกราฟในชีตเดียว (รวมกราฟที่มีแถวซ่อนด้วยตัวกรอง)
async function sampleC() {
  const s = createEmptySheet('หลายกราฟ');
  s.colWidths = { 0: 64, 1: 200, 2: 80, 3: 160, 5: 64, 6: 240, 7: 96 };
  s.rowHeights = { 0: 40, 6: 60 };
  put(s, 'A1', 'สาขา');
  put(s, 'B1', 'ไตรมาส 1');
  put(s, 'C1', 'ไตรมาส 2');
  put(s, 'D1', 'ไตรมาส 3');
  ['บางนา', 'รังสิต', 'เชียงใหม่', 'ขอนแก่น', 'หาดใหญ่'].forEach((b, i) => {
    put(s, `A${i + 2}`, b);
    put(s, `B${i + 2}`, 100 + i * 20);
    put(s, `C${i + 2}`, 90 + i * 25);
    put(s, `D${i + 2}`, 120 - i * 10);
  });
  // ตัวกรองซ่อนแถว "เชียงใหม่" (แถว 4): แอปไม่พล็อตแถวนี้ ไฟล์เขียนเป็นแถวซ่อน + plotVisOnly
  s.filter = { r1: 0, c1: 0, r2: 5, c2: 3, columns: {} };
  s.hiddenRows = [3];

  const wb: Workbook = { sheets: [s], activeSheet: 0 };
  s.charts = [
    chartFrom(wb, 'หลายกราฟ', { r1: 0, c1: 0, r2: 5, c2: 3 }, 'bar', { row: 7, col: 0, dx: 10, dy: 5 }, {
      title: 'ยอดขายตามสาขา (ซ่อนเชียงใหม่ด้วยตัวกรอง)',
      width: 520,
      height: 300,
    }),
    chartFrom(wb, 'หลายกราฟ', { r1: 0, c1: 0, r2: 5, c2: 3 }, 'line', { row: 7, col: 5, dx: 30, dy: 0 }, {
      title: 'แนวโน้ม',
      width: 400,
      height: 260,
      legend: 'right',
    }),
    chartFrom(wb, 'หลายกราฟ', { r1: 0, c1: 0, r2: 5, c2: 1 }, 'pie', { row: 20, col: 2, dx: 0, dy: 0 }, {
      title: 'ไตรมาส 1',
      width: 300,
      height: 300,
    }),
  ];
  await save('sample-c-widths-multi.xlsx', wb);
}

await sampleA();
await sampleB();
await sampleC();
