// Spike: แทรกกราฟหนึ่งกราฟลงไฟล์ .xlsx ที่ ExcelJS เขียนออกมา (post-process zip)
// ใช้ jszip ที่ติดมากับ exceljs (ไม่ได้เป็น dependency ตรงของโปรเจกต์)
import JSZip from 'jszip';
import { chartXml, drawingXml, relsXml, REL, CT } from './chartXml.mjs';

// ลำดับ element ใน <worksheet> (ECMA-376 CT_Worksheet): <drawing> ต้องอยู่ก่อน element เหล่านี้
const AFTER_DRAWING = [
  'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls',
  'webPublishItems', 'tableParts', 'extLst',
];

function nextFreeNumber(zip, prefix, ext) {
  let n = 1;
  while (zip.file(`${prefix}${n}${ext}`)) n++;
  return n;
}

function nextRelId(relsText) {
  const ids = [...relsText.matchAll(/Id="rId(\d+)"/g)].map((m) => Number(m[1]));
  return `rId${(ids.length ? Math.max(...ids) : 0) + 1}`;
}

/** sheetPath เช่น 'xl/worksheets/sheet1.xml' — หาได้จาก workbook.xml + rels (ที่นี่รับตรงๆ เพื่อความง่าย) */
export async function injectChart(xlsxBytes, { sheetPath, type, title, series, anchor }) {
  const zip = await JSZip.loadAsync(xlsxBytes);

  const sheetXml = await zip.file(sheetPath)?.async('string');
  if (!sheetXml) throw new Error('ไม่พบชีต ' + sheetPath);
  if (/<drawing\b/.test(sheetXml)) {
    // ชีตมี drawing อยู่แล้ว (เช่น รูปภาพ) ต้องเพิ่ม anchor เข้า drawing เดิมแทน — ยังไม่ทำใน spike
    throw new Error('ชีตมี <drawing> อยู่แล้ว: spike ยังไม่รองรับการรวมเข้า drawing เดิม');
  }

  const chartNo = nextFreeNumber(zip, 'xl/charts/chart', '.xml');
  const drawingNo = nextFreeNumber(zip, 'xl/drawings/drawing', '.xml');
  const chartPath = `xl/charts/chart${chartNo}.xml`;
  const drawingPath = `xl/drawings/drawing${drawingNo}.xml`;

  // 1) กราฟ
  zip.file(chartPath, chartXml(type, title, series));

  // 2) drawing + rels ของ drawing → กราฟ
  zip.file(drawingPath, drawingXml(anchor, 'rId1', `Chart ${chartNo}`));
  zip.file(`xl/drawings/_rels/drawing${drawingNo}.xml.rels`, relsXml([
    { id: 'rId1', type: REL.chart, target: `../charts/chart${chartNo}.xml` },
  ]));

  // 3) rels ของชีต → drawing (สร้างใหม่หรือเติมเข้าไฟล์เดิม)
  const sheetName = sheetPath.split('/').pop();
  const sheetRelsPath = sheetPath.replace(sheetName, `_rels/${sheetName}.rels`);
  let sheetRels = await zip.file(sheetRelsPath)?.async('string');
  let drawingRelId;
  const drawingRel = `Type="${REL.drawing}" Target="../drawings/drawing${drawingNo}.xml"`;
  if (sheetRels) {
    drawingRelId = nextRelId(sheetRels);
    sheetRels = sheetRels.replace('</Relationships>', `<Relationship Id="${drawingRelId}" ${drawingRel}/></Relationships>`);
  } else {
    drawingRelId = 'rId1';
    sheetRels = relsXml([{ id: drawingRelId, type: REL.drawing, target: `../drawings/drawing${drawingNo}.xml` }]);
  }
  zip.file(sheetRelsPath, sheetRels);

  // 4) <drawing r:id> ในชีต ณ ตำแหน่งที่ถูกต้องตามลำดับ schema
  if (!/xmlns:r="/.test(sheetXml.slice(0, 1000))) throw new Error('ชีตไม่มี xmlns:r');
  const tag = `<drawing r:id="${drawingRelId}"/>`;
  let newSheet = null;
  for (const el of AFTER_DRAWING) {
    const idx = sheetXml.search(new RegExp(`<${el}[\\s/>]`));
    if (idx >= 0) { newSheet = sheetXml.slice(0, idx) + tag + sheetXml.slice(idx); break; }
  }
  if (!newSheet) newSheet = sheetXml.replace('</worksheet>', `${tag}</worksheet>`);
  zip.file(sheetPath, newSheet);

  // 5) [Content_Types].xml
  let ct = await zip.file('[Content_Types].xml').async('string');
  const addOverride = (part, contentType) => {
    if (!ct.includes(`PartName="/${part}"`)) {
      ct = ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${contentType}"/></Types>`);
    }
  };
  addOverride(chartPath, CT.chart);
  addOverride(drawingPath, CT.drawing);
  zip.file('[Content_Types].xml', ct);

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}
