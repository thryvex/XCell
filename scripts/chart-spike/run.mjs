// Spike กราฟใน .xlsx — รัน: node scripts/chart-spike/run.mjs [ไฟล์กราฟจาก Excel/Google Sheets ...]
// ขั้นที่ 1: ตรวจความสามารถของ ExcelJS
// ขั้นที่ 2: สร้าง .xlsx ด้วย ExcelJS แล้วแทรกกราฟแท่ง/เส้น/วงกลมเอง แล้วอ่าน-เขียนกลับด้วย ExcelJS
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { SaxesParser } from 'saxes';
import { injectChart } from './injectChart.mjs';
import { ref } from './chartXml.mjs';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '../..');
const outDir = path.join(here, 'out');
fs.mkdirSync(outDir, { recursive: true });

const log = (...a) => console.log(...a);
const hr = (t) => log(`\n=== ${t} ===`);

// ---------------------------------------------------------------- ขั้นที่ 1
hr('ขั้นที่ 1: ExcelJS ' + JSON.parse(fs.readFileSync(path.join(root, 'node_modules/exceljs/package.json'), 'utf8')).version);
{
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('x');
  log('worksheet.addChart:', typeof ws.addChart, '| workbook.addChart:', typeof wb.addChart);
  log('worksheet.addImage:', typeof ws.addImage, '(มีแค่รูปภาพ)');
  const dts = fs.readFileSync(path.join(root, 'node_modules/exceljs/index.d.ts'), 'utf8');
  log('index.d.ts กล่าวถึง "chart":', (dts.match(/chart/gi) || []).length, 'ครั้ง');
  const xlsxJs = fs.readFileSync(path.join(root, 'node_modules/exceljs/lib/xlsx/xlsx.js'), 'utf8');
  log('lib/xlsx/xlsx.js กล่าวถึง "xl/charts":', (xlsxJs.match(/xl\/charts/g) || []).length, 'ครั้ง');
  const anchor = fs.readFileSync(path.join(root, 'node_modules/exceljs/lib/xlsx/xform/drawing/two-cell-anchor-xform.js'), 'utf8');
  log('two-cell-anchor รู้จัก xdr:pic:', anchor.includes("'xdr:pic'"), '| xdr:graphicFrame:', anchor.includes('graphicFrame'));
}

// ---------------------------------------------------------------- ขั้นที่ 2
hr('ขั้นที่ 2: สร้างข้อมูลด้วย ExcelJS แล้วแทรกกราฟ');
const SHEET = 'ข้อมูล';
const categories = ['อาหาร', 'เดินทาง', 'ที่พัก', 'อื่นๆ'];
const amounts = [4500, 1200, 8000, 650];

async function baseWorkbookBytes() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(SHEET);
  ws.getCell('A1').value = 'หมวด';
  ws.getCell('B1').value = 'จำนวนเงิน';
  ws.getRow(1).font = { bold: true };
  categories.forEach((c, i) => {
    ws.getCell(`A${i + 2}`).value = c;
    ws.getCell(`B${i + 2}`).value = amounts[i];
    ws.getCell(`B${i + 2}`).numFmt = '#,##0';
  });
  ws.getColumn(1).width = 14;
  ws.getColumn(2).width = 14;
  return Buffer.from(await wb.xlsx.writeBuffer());
}

const series = {
  nameRef: ref(SHEET, '$B$1'), name: 'จำนวนเงิน',
  catRef: ref(SHEET, '$A$2:$A$5'), categories,
  valRef: ref(SHEET, '$B$2:$B$5'), values: amounts,
};
const anchor = { fromCol: 3, fromRow: 1, toCol: 11, toRow: 18 }; // D2 ถึง K18 โดยประมาณ

const outputs = [];
for (const [type, title] of [['bar', 'ค่าใช้จ่ายตามหมวด (แท่ง)'], ['line', 'ค่าใช้จ่ายตามหมวด (เส้น)'], ['pie', 'สัดส่วนค่าใช้จ่าย (วงกลม)']]) {
  const bytes = await injectChart(await baseWorkbookBytes(), { sheetPath: 'xl/worksheets/sheet1.xml', type, title, series, anchor });
  const file = path.join(outDir, `chart-${type}.xlsx`);
  fs.writeFileSync(file, bytes);
  outputs.push(file);
  log('เขียน', path.relative(root, file), `(${bytes.length} ไบต์)`);
}

// ---------------------------------------------------------------- ตรวจโครงสร้าง
async function inspect(file) {
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
  const charts = names.filter((n) => /^xl\/charts\//.test(n));
  const drawings = names.filter((n) => /^xl\/drawings\/[^/]+\.xml$/.test(n));
  const sheets = names.filter((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
  let sheetDrawingTags = 0;
  for (const s of sheets) sheetDrawingTags += ((await zip.file(s).async('string')).match(/<drawing\b/g) || []).length;
  let graphicFrames = 0;
  for (const d of drawings) graphicFrames += ((await zip.file(d).async('string')).match(/graphicFrame/g) || []).length / 2;
  const ct = await zip.file('[Content_Types].xml').async('string');
  // XML ทุกไฟล์ต้อง well-formed
  const bad = [];
  for (const n of names.filter((x) => /\.(xml|rels)$/.test(x))) {
    try {
      const p = new SaxesParser({ xmlns: true });
      p.write(await zip.file(n).async('string')).close();
    } catch (e) {
      bad.push(`${n}: ${e.message}`);
    }
  }
  return {
    charts: charts.length,
    drawings: drawings.length,
    graphicFrames,
    sheetDrawingTags,
    chartContentType: (ct.match(/drawingml\.chart\+xml/g) || []).length,
    xmlErrors: bad,
  };
}

async function roundTripExcelJS(file) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(fs.readFileSync(file));
  const ws = wb.worksheets[0];
  const modelHasDrawing = Boolean(ws.model?.drawing);
  const out = file.replace(/\.xlsx$/, '.exceljs-roundtrip.xlsx');
  await wb.xlsx.writeFile(out);
  return { out, modelHasDrawing, images: ws.getImages().length };
}

hr('ตรวจโครงสร้างไฟล์ที่แทรกกราฟ');
for (const f of outputs) log(path.basename(f), JSON.stringify(await inspect(f)));

hr('อ่านแล้วเขียนใหม่ด้วย ExcelJS (load → writeFile)');
for (const f of outputs) {
  const rt = await roundTripExcelJS(f);
  log(path.basename(f), '→', path.basename(rt.out), '| ExcelJS เห็น drawing ในชีต:', rt.modelHasDrawing, '| images:', rt.images);
  log('   หลังเขียนใหม่:', JSON.stringify(await inspect(rt.out)));
}

// ไฟล์ที่ผู้ใช้ส่งมา (สร้างจาก Excel/Google Sheets จริง)
const extra = process.argv.slice(2);
if (extra.length) {
  hr('ไฟล์จากภายนอก');
  for (const f of extra) {
    log(f, JSON.stringify(await inspect(f)));
    const rt = await roundTripExcelJS(f);
    log('   หลัง ExcelJS เขียนใหม่ →', path.basename(rt.out), JSON.stringify(await inspect(rt.out)));
  }
}
