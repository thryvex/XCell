// Spike: ไฟล์ที่มีกราฟผ่านเส้นทางเปิด/บันทึกของแอปจริง (src/core/xlsx.ts — ไม่ได้แก้โค้ดแอป)
// รัน: npx vite-node scripts/chart-spike/appRoundtrip.ts [ไฟล์ .xlsx ...]
import fs from 'node:fs';
import path from 'node:path';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { importFromXlsxBuffer, exportToXlsxBuffer, detectUnsupportedFeatures } from '../../src/core/xlsx';

const outDir = path.resolve(__dirname, 'out');
const files = process.argv.slice(2).length
  ? process.argv.slice(2)
  : ['chart-bar.xlsx', 'chart-line.xlsx', 'chart-pie.xlsx'].map((f) => path.join(outDir, f));

async function countCharts(bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  return Object.keys(zip.files).filter((n) => n.startsWith('xl/charts/chart')).length;
}

for (const file of files) {
  const bytes = new Uint8Array(fs.readFileSync(file));
  const excel = new ExcelJS.Workbook();
  await excel.xlsx.load(bytes as any);
  const banner = detectUnsupportedFeatures(excel);
  const model = await importFromXlsxBuffer(bytes);
  const saved = await exportToXlsxBuffer(model);
  const out = file.replace(/\.xlsx$/, '.app-saved.xlsx');
  fs.writeFileSync(out, saved);
  console.log(
    path.basename(file),
    '| กราฟก่อน:', await countCharts(bytes),
    '| แถบเตือนส่วนที่ไม่รองรับ:', JSON.stringify(banner),
    '| กราฟหลังแอปบันทึก:', await countCharts(saved),
    '| ข้อมูล A2/B2:', model.sheets[0].cells['A2']?.raw, model.sheets[0].cells['B2']?.raw
  );
}
