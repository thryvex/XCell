/**
 * V7: เซลว่างแบบไหนที่ทำให้ Google Sheets ใช้ความกว้างคอลัมน์ โดยผู้ใช้มองไม่เห็นอะไร (ต่อจาก V6)
 * รัน: npx vite-node scripts/chart-spike/makeV7.ts
 *
 * ไฟล์เดียว 5 ชีต; defaultColWidth = 13.75 (คอลัมน์ที่ถูกเมิน ≈ 96 px)
 *   U1: K (30) มี <c r="K1"/> ไม่มีสไตล์; L (12) ไม่มีเซล (ตัวเทียบ)
 *   U2: K (30) มี <c r="K1" s="0"/>; L (12) ไม่มีเซล
 *   U3: K (30) มีเซลว่างที่มีแค่ alignment vertical="bottom" (ค่าเริ่มต้นอยู่แล้ว มองไม่เห็น); L (12) ไม่มีเซล
 *   U4: เหมือน U3 + ตัวกรอง A1:B3 — ดูผลข้างเคียง: Ctrl+End, ช่วงตัวกรอง, พื้นที่พิมพ์
 *   U5: ตัวกรอง A1:B5 ซ่อนแถว 3; K (30) มีเซลแบบ U3 ใน K3 (แถวที่ซ่อน);
 *       M1:N1 ผสาน (M 20, N 12) มีเซลแบบ U3 ใน M1; P (30) ไม่มีเซล (ตัวเทียบ)
 */
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';

const OUT = join(__dirname, 'out');
const INVISIBLE: Partial<ExcelJS.Alignment> = { vertical: 'bottom' };

const wb = new ExcelJS.Workbook();
function sheet(name: string, rows = 3): ExcelJS.Worksheet {
  const ws = wb.addWorksheet(name);
  ws.properties.defaultColWidth = 13.75;
  ws.properties.defaultRowHeight = 21;
  ws.getCell('A1').value = 'หัว';
  ws.getCell('B1').value = 'ค่า';
  for (let r = 2; r <= rows; r++) {
    ws.getCell(`A${r}`).value = `แถว${r}`;
    ws.getCell(`B${r}`).value = r;
  }
  return ws;
}

for (const name of ['U1', 'U2']) {
  const ws = sheet(name);
  ws.getColumn('K').width = 30;
  ws.getColumn('L').width = 12;
}

const u3 = sheet('U3');
u3.getColumn('K').width = 30;
u3.getColumn('L').width = 12;
u3.getCell('K1').alignment = INVISIBLE;

const u4 = sheet('U4');
u4.getColumn('K').width = 30;
u4.getColumn('L').width = 12;
u4.getCell('K1').alignment = INVISIBLE;
u4.autoFilter = 'A1:B3';

const u5 = sheet('U5', 5);
u5.getColumn('K').width = 30;
u5.getColumn('M').width = 20;
u5.getColumn('N').width = 12;
u5.getColumn('P').width = 30;
u5.autoFilter = 'A1:B5';
u5.getRow(3).hidden = true;
u5.getCell('K3').alignment = INVISIBLE;
u5.getCell('M1').alignment = INVISIBLE;
u5.mergeCells('M1:N1');

const zip = await JSZip.loadAsync(await wb.xlsx.writeBuffer());
// U1, U2: ใส่เซลว่างลงแถว 1 ตรงๆ (ExcelJS ไม่เขียนเซลที่ไม่มีค่าและไม่มีสไตล์)
const addRaw = async (n: number, cell: string) => {
  const p = `xl/worksheets/sheet${n}.xml`;
  let xml = await zip.file(p)!.async('string');
  const row1 = /<row r="1"[^>]*>.*?<\/row>/s.exec(xml)![0];
  xml = xml.replace(row1, row1.replace('</row>', `${cell}</row>`));
  zip.file(p, xml);
};
await addRaw(1, '<c r="K1"/>');
await addRaw(2, '<c r="K1" s="0"/>');

const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
writeFileSync(join(OUT, 'sample-v7-invisible-cells.xlsx'), bytes);

const check = await JSZip.loadAsync(bytes);
const styles = await check.file('xl/styles.xml')!.async('string');
console.log('cellXfs: ' + (styles.match(/<cellXfs.*?<\/cellXfs>/s)?.[0] ?? '-'));
for (let n = 1; n <= 5; n++) {
  const xml = await check.file(`xl/worksheets/sheet${n}.xml`)!.async('string');
  console.log(`U${n}`);
  console.log('  ' + (xml.match(/<dimension[^>]*>/)?.[0] ?? '-'));
  console.log('  ' + (xml.match(/<cols>.*?<\/cols>/s)?.[0] ?? 'no cols'));
  console.log('  cells: ' + [...xml.matchAll(/<c r="([A-Z]+\d+)"([^>]*)\/?>/g)].map((m) => m[1] + m[2]).join(','));
  console.log('  ' + [xml.match(/<autoFilter[^>]*>/)?.[0], xml.match(/<mergeCells.*?<\/mergeCells>/s)?.[0], xml.match(/<row r="3"[^>]*>/)?.[0]].filter(Boolean).join(' '));
}
