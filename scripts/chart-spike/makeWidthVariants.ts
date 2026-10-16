/**
 * ไฟล์ทดลองเรื่องความกว้างคอลัมน์ใน Google Sheets (ต่อจากผลตรวจ sample-c)
 * รัน: npx vite-node scripts/chart-spike/makeWidthVariants.ts
 *
 * สมมติฐาน: Google Sheets ใช้ความกว้างที่ตั้งไว้เฉพาะคอลัมน์ในช่วง <dimension> และคอลัมน์ที่ไม่ได้ตั้ง
 * ใช้ค่าเริ่มต้นของตัวเอง (~60 px) ไม่ใช่ 110 px ของแอป
 *   V1: แก้ <dimension> ให้ครอบทุกคอลัมน์/แถวที่ตั้งขนาดหรือมีกราฟ (แก้ XML หลังเขียน — ExcelJS ไม่มี API ตั้ง dimension)
 *   V2: ตั้ง sheetFormatPr defaultColWidth / defaultRowHeight ผ่าน ws.properties ของ ExcelJS ให้ตรงกับแอป
 *   V3: ทั้งสองอย่าง
 *   V4: V2 + มีค่าในเซล G1/H1 (คอลัมน์ G/H ไม่ว่าง)    — แยกสาเหตุที่ G/H ถูกเมิน
 *   V5: V2 + ไม่มีกราฟเส้น (ไม่มีกราฟทับ G/H)         — แยกสาเหตุที่ G/H ถูกเมิน
 *   รันเฉพาะ V4/V5: npx vite-node scripts/chart-spike/makeWidthVariants.ts --gh
 * ทุกไฟล์มีข้อความ "ไม่มีข้อมูล" ใน C3 (ช่วงค่าของกราฟเส้น ชุด "ไตรมาส 2") เพื่อดูว่า Google Sheets วาดเป็น 0 หรือช่องว่าง
 */
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import JSZip from 'jszip';
import { workbookToExcelJS } from '../../src/core/xlsx';
import { injectChartsIntoXlsx } from '../../src/core/chartXlsx';
import {
  buildChartFromRange,
  newChart,
  workbookChartContext,
  DEFAULT_COL_WIDTH_PX,
  DEFAULT_ROW_HEIGHT_PX,
} from '../../src/core/chart';
import { Evaluator } from '../../src/core/evaluator';
import { createEmptySheet, Sheet, SheetChart, Workbook, ChartType } from '../../src/core/model';

const OUT = join(__dirname, 'out');

function workbook(opts: { ghValues?: boolean; noLine?: boolean } = {}): Workbook {
  const s: Sheet = createEmptySheet('หลายกราฟ');
  s.colWidths = { 0: 64, 1: 200, 2: 80, 3: 160, 5: 64, 6: 240, 7: 96 };
  s.rowHeights = { 0: 40, 6: 60 };
  const put = (k: string, v: string | number) => (s.cells[k] = { raw: String(v) });
  put('A1', 'สาขา');
  put('B1', 'ไตรมาส 1');
  put('C1', 'ไตรมาส 2');
  put('D1', 'ไตรมาส 3');
  ['บางนา', 'รังสิต', 'เชียงใหม่', 'ขอนแก่น', 'หาดใหญ่'].forEach((b, i) => {
    put(`A${i + 2}`, b);
    put(`B${i + 2}`, 100 + i * 20);
    put(`C${i + 2}`, 90 + i * 25);
    put(`D${i + 2}`, 120 - i * 10);
  });
  put('C3', 'ไม่มีข้อมูล'); // ข้อความในช่วงค่า: แอปวาดเป็น 0
  if (opts.ghValues) {
    put('G1', 'ค่า G');
    put('H1', 'ค่า H');
  }
  s.filter = { r1: 0, c1: 0, r2: 5, c2: 3, columns: {} };
  s.hiddenRows = [3];

  const wb: Workbook = { sheets: [s], activeSheet: 0 };
  const ctx = workbookChartContext(wb, new Evaluator(wb));
  const chart = (range: { r1: number; c1: number; r2: number; c2: number }, type: ChartType, anchor: SheetChart['anchor'], extra: Partial<SheetChart>) => {
    const res = buildChartFromRange(ctx, s.name, range);
    if (!res.ok) throw new Error(res.error);
    return { ...newChart(type, res.series, anchor), ...extra };
  };
  s.charts = [
    chart({ r1: 0, c1: 0, r2: 5, c2: 3 }, 'bar', { row: 7, col: 0, dx: 10, dy: 5 }, { title: 'ยอดขายตามสาขา', width: 520, height: 300 }),
    chart({ r1: 0, c1: 0, r2: 5, c2: 3 }, 'line', { row: 7, col: 5, dx: 30, dy: 0 }, {
      title: 'แนวโน้ม (C3 เป็นข้อความ)',
      width: 400,
      height: 260,
      legend: 'right',
      dataLabels: true,
    }),
    chart({ r1: 0, c1: 0, r2: 5, c2: 1 }, 'pie', { row: 20, col: 2, dx: 0, dy: 0 }, { title: 'ไตรมาส 1', width: 300, height: 300 }),
  ];
  if (opts.noLine) s.charts = s.charts.filter((c) => c.type !== 'line');
  return wb;
}

function colLetters(c: number): string {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

async function build(file: string, opts: { dimension: boolean; defaults: boolean; ghValues?: boolean; noLine?: boolean }) {
  const wb = workbook(opts);
  const excel = await workbookToExcelJS(wb);
  if (opts.defaults) {
    // ค่าเริ่มต้นของแอป: คอลัมน์ 110 px (หน่วยละ 8 px ตามที่แอปแปลงความกว้าง) แถว 28 px = 21 pt
    const ws = excel.worksheets[0];
    ws.properties.defaultColWidth = DEFAULT_COL_WIDTH_PX / 8;
    ws.properties.defaultRowHeight = DEFAULT_ROW_HEIGHT_PX * 0.75;
  }
  const raw = new Uint8Array(await excel.xlsx.writeBuffer());
  let bytes = await injectChartsIntoXlsx(raw, wb.sheets, workbookChartContext(wb, new Evaluator(wb), true));

  const zip = await JSZip.loadAsync(bytes);
  const sheetPath = 'xl/worksheets/sheet1.xml';
  let xml = await zip.file(sheetPath)!.async('string');
  if (opts.dimension) {
    // ครอบคอลัมน์/แถวที่ตั้งขนาด และมุมขวาล่างของกราฟทุกอัน (จาก drawing ที่เขียนแล้ว)
    const drawing = await zip.file('xl/drawings/drawing1.xml')!.async('string');
    const tos = [...drawing.matchAll(/<xdr:to><xdr:col>(\d+)<\/xdr:col><xdr:colOff>\d+<\/xdr:colOff><xdr:row>(\d+)<\/xdr:row>/g)];
    const s = wb.sheets[0];
    const maxCol = Math.max(...Object.keys(s.colWidths).map(Number), ...tos.map((m) => Number(m[1])));
    const maxRow = Math.max(...Object.keys(s.rowHeights).map(Number), ...tos.map((m) => Number(m[2])));
    xml = xml.replace(/<dimension ref="[^"]*"\/>/, `<dimension ref="A1:${colLetters(maxCol)}${maxRow + 1}"/>`);
    zip.file(sheetPath, xml);
    bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  }
  writeFileSync(join(OUT, file), bytes);

  const check = await (await JSZip.loadAsync(bytes)).file(sheetPath)!.async('string');
  console.log(`\n${file}`);
  console.log('  ' + check.match(/<dimension[^>]*>/)![0]);
  console.log('  ' + check.match(/<sheetFormatPr[^>]*>/)![0]);
  console.log('  ' + check.match(/<cols>.*?<\/cols>/s)![0]);
}

if (process.argv.includes('--gh')) {
  await build('sample-c-v4-gh-values.xlsx', { dimension: false, defaults: true, ghValues: true });
  await build('sample-c-v5-no-line.xlsx', { dimension: false, defaults: true, noLine: true });
} else {
  await build('sample-c-v1-dimension.xlsx', { dimension: true, defaults: false });
  await build('sample-c-v2-defaults.xlsx', { dimension: false, defaults: true });
  await build('sample-c-v3-both.xlsx', { dimension: true, defaults: true });
}
