/**
 * V6: หากฎที่ Google Sheets ใช้/เมินความกว้างคอลัมน์ (ต่อจาก V4/V5)
 * รัน: npx vite-node scripts/chart-spike/makeV6.ts
 *
 * ไฟล์เดียว 4 ชีต ทุกชีตมีข้อมูลเฉพาะ A1:B3 และ defaultColWidth = 13.75 (คอลัมน์ที่ถูกเมิน ≈ 95 px)
 *   T1: ตั้งความกว้างเฉพาะ K (30) และ L (12) ไม่มีเซล              — แยกกฎ (ก) คอลัมน์ข้อมูลสุดท้าย + 2 กับ (ข) <col> ไม่กี่รายการแรก
 *   T2: K (30) มีเซลว่างสีพื้นใน K1, L (12) มีเซลว่างมีขอบใน L1, M (20) ไม่มีเซล (ตัวเทียบ)
 *   T3: K–M เป็นกลุ่มเดียว <col min="11" max="13" width="30"> ไม่มีเซล
 *   T4: เหมือน T3 + เซลว่างสีพื้นใน K1 เท่านั้น
 * ใช้ ExcelJS ตรงๆ (ไม่ผ่านโค้ดแอป) เพื่อคุม XML ได้ครบ; กลุ่มใน T3/T4 แก้ XML หลังเขียน
 */
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const OUT = join(__dirname, 'out');
const FILL: ExcelJS.Fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE68A' } };
const BORDER: Partial<ExcelJS.Borders> = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };

const wb = new ExcelJS.Workbook();
function sheet(name: string): ExcelJS.Worksheet {
  const ws = wb.addWorksheet(name);
  ws.properties.defaultColWidth = 13.75;
  ws.properties.defaultRowHeight = 21;
  ws.getCell('A1').value = 'หัว';
  ws.getCell('B1').value = 'ค่า';
  ws.getCell('A2').value = 'ก';
  ws.getCell('B2').value = 1;
  ws.getCell('A3').value = 'ข';
  ws.getCell('B3').value = 2;
  return ws;
}

const t1 = sheet('T1');
t1.getColumn('K').width = 30;
t1.getColumn('L').width = 12;

const t2 = sheet('T2');
t2.getColumn('K').width = 30;
t2.getColumn('L').width = 12;
t2.getColumn('M').width = 20;
t2.getCell('K1').fill = FILL;
t2.getCell('L1').border = BORDER;

sheet('T3');
const t4 = sheet('T4');
t4.getCell('K1').fill = FILL;

const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
// T3 = sheet3, T4 = sheet4: ใส่ <cols> กลุ่มเดียว K–M (ก่อน <sheetData>)
for (const n of [3, 4]) {
  const p = `xl/worksheets/sheet${n}.xml`;
  let xml = await zip.file(p)!.async('string');
  if (xml.includes('<cols>')) throw new Error(`${p} มี <cols> อยู่แล้ว`);
  xml = xml.replace('<sheetData', '<cols><col min="11" max="13" width="30" customWidth="1"/></cols><sheetData');
  zip.file(p, xml);
}
const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
writeFileSync(join(OUT, 'sample-v6-width-rules.xlsx'), bytes);

const check = await JSZip.loadAsync(bytes);
for (let n = 1; n <= 4; n++) {
  const xml = await check.file(`xl/worksheets/sheet${n}.xml`)!.async('string');
  console.log(`T${n}`);
  console.log('  ' + (xml.match(/<dimension[^>]*>/)?.[0] ?? '-'));
  console.log('  ' + (xml.match(/<cols>.*?<\/cols>/s)?.[0] ?? 'no cols'));
  console.log('  cells: ' + [...xml.matchAll(/<c r="([A-Z]+\d+)"( s="\d+")?/g)].map((m) => m[1] + (m[2] ? '(s)' : '')).join(','));
}
