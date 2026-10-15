import { Workbook, PrintSettings, CellStyle } from './model';
import { Evaluator } from './evaluator';
import { cellToA1 } from './address';
import { formatDisplayValue } from './numfmt';
import { isCellError } from './functions';
import { PrintPage, pageGeometry, colWidthMm, rowHeightMm, PX_TO_MM } from './paginate';

/** escape ข้อความทุกชนิดก่อนใส่ใน HTML (ทั้งเนื้อหาและค่า attribute) */
export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/** สีต้องเป็น #RRGGBB เท่านั้น (กันการแทรก CSS) */
function safeColor(color: string | undefined): string | undefined {
  return color && /^#[0-9a-fA-F]{6}$/.test(color) ? color : undefined;
}

/** ชื่อฟอนต์จากไฟล์ภายนอก: อนุญาตเฉพาะตัวอักษร ตัวเลข ช่องว่าง ขีด (กันการแทรก CSS) */
function safeFontFamily(name: string | undefined): string | undefined {
  return name && /^[\w\s฀-๿-]{1,64}$/.test(name) ? name : undefined;
}

const BASE_FONT_PX = 14; // เท่ากับตัวอักษรในตาราง
const CELL_PADDING_PX = 6;
const ERROR_COLOR = '#dc2626';

export interface PrintMeta {
  fileName: string;
  printedAt: string; // วันที่พิมพ์ (จัดรูปแบบตามโหมดปีแล้ว)
}

function mm(n: number): string {
  return `${Math.round(n * 1000) / 1000}mm`;
}

function cellCss(style: CellStyle | undefined, settings: PrintSettings, scale: number, isNumber: boolean, isError: boolean): string {
  const parts: string[] = [];
  const s = style ?? {};
  if (s.bold) parts.push('font-weight:bold');
  if (s.italic) parts.push('font-style:italic');
  if (s.underline) parts.push('text-decoration:underline');
  const color = isError ? ERROR_COLOR : safeColor(s.color);
  if (color) parts.push(`color:${color}`);
  const fill = settings.fills ? safeColor(s.fill) : undefined;
  if (fill) parts.push(`background-color:${fill}`);
  parts.push(`text-align:${s.align ?? (isNumber ? 'right' : 'left')}`);
  if (s.wrap) parts.push('white-space:normal;word-break:break-word');
  if (s.border && s.border !== 'none') parts.push('border:0.3mm solid #000000');
  if (s.fontSize) parts.push(`font-size:${mm(s.fontSize * (25.4 / 72) * scale)}`);
  const family = safeFontFamily(s.fontFamily);
  if (family) parts.push(`font-family:'${family}','Sarabun',sans-serif`);
  return parts.join(';');
}

/**
 * HTML ของหนึ่งหน้า (ส่วน <div class="page">) จากโมเดลข้อมูล ใช้ค่าที่แสดงเหมือนในตาราง
 */
export function renderPrintPage(
  wb: Workbook,
  evaluator: Evaluator,
  page: PrintPage,
  pageNumber: number,
  totalPages: number,
  settings: PrintSettings,
  meta: PrintMeta,
  labels: { pageNumber: string } // เช่น "หน้า {page} / {total}"
): string {
  const sheet = wb.sheets[page.sheetIndex];
  const geo = pageGeometry(settings);
  const scale = page.scale;
  const rows = [...page.titleRows, ...page.rows];
  const cols = [...page.titleCols, ...page.cols];
  const tableW = cols.reduce((s, c) => s + colWidthMm(sheet, c) * scale, 0);

  // เซลผสาน: หาเซลแรกของกลุ่มที่อยู่ในหน้านี้ และจำนวนแถว/คอลัมน์ของกลุ่มที่อยู่ในหน้า
  const anchors = new Map<string, { rowspan: number; colspan: number; isTrueAnchor: boolean }>();
  const covered = new Set<string>();
  for (const m of sheet.merges) {
    const mRows = rows.filter((r) => r >= m.r1 && r <= m.r2);
    const mCols = cols.filter((c) => c >= m.c1 && c <= m.c2);
    if (mRows.length === 0 || mCols.length === 0) continue;
    const first = `${mRows[0]},${mCols[0]}`;
    anchors.set(first, {
      rowspan: mRows.length,
      colspan: mCols.length,
      isTrueAnchor: mRows[0] === m.r1 && mCols[0] === m.c1,
    });
    for (const r of mRows) for (const c of mCols) if (`${r},${c}` !== first) covered.add(`${r},${c}`);
  }

  const out: string[] = [];
  out.push('<div class="page">');
  if (settings.header) {
    out.push(
      `<div class="page-header"><span>${escapeHtml(meta.fileName)}</span><span>${escapeHtml(meta.printedAt)}</span></div>`
    );
  }

  out.push(`<table style="width:${mm(tableW)};font-size:${mm(BASE_FONT_PX * PX_TO_MM * scale)}">`);
  out.push('<colgroup>');
  for (const c of cols) out.push(`<col style="width:${mm(colWidthMm(sheet, c) * scale)}">`);
  out.push('</colgroup><tbody>');

  for (const r of rows) {
    out.push(`<tr style="height:${mm(rowHeightMm(sheet, r) * scale)}">`);
    for (const c of cols) {
      const pos = `${r},${c}`;
      if (covered.has(pos)) continue;
      const merge = anchors.get(pos);
      const key = cellToA1(r, c);
      const cell = sheet.cells[key];
      const showValue = !merge || merge.isTrueAnchor;
      const value = showValue && cell?.raw ? evaluator.getCellValue(key, page.sheetIndex) : '';
      const display = value === '' ? '' : formatDisplayValue(value, cell?.fmt);
      const isError = isCellError(value);
      const css = cellCss(cell?.style, settings, scale, typeof value === 'number', isError);
      const span = merge
        ? `${merge.rowspan > 1 ? ` rowspan="${merge.rowspan}"` : ''}${merge.colspan > 1 ? ` colspan="${merge.colspan}"` : ''}`
        : '';
      out.push(`<td${span} style="${escapeHtml(css)}">${escapeHtml(display)}</td>`);
    }
    out.push('</tr>');
  }
  out.push('</tbody></table>');

  if (settings.pageNumbers) {
    const text = labels.pageNumber.replace('{page}', String(pageNumber)).replace('{total}', String(totalPages));
    out.push(`<div class="page-footer">${escapeHtml(text)}</div>`);
  }
  out.push('</div>');

  return out.join('');
}

/**
 * HTML ทั้งเอกสาร (ใช้ทั้งตัวอย่างก่อนพิมพ์และการพิมพ์/PDF จริง)
 * - CSP ห้ามโหลดสคริปต์และทรัพยากรภายนอก (ฟอนต์เป็น data: URL ที่ฝังมา)
 * - @page กำหนดขนาด/แนวกระดาษ ขอบเป็น 0 (ระยะขอบอยู่ใน .page เอง) ให้ตรงกับตัวอย่าง
 * - พื้นหลังขาวเสมอ ไม่ขึ้นกับธีมสี
 */
export function renderPrintDocument(pagesHtml: string[], settings: PrintSettings, fontCss: string): string {
  const g = pageGeometry(settings);
  const css = `
${fontCss}
@page { size: ${mm(g.pageW)} ${mm(g.pageH)}; margin: 0; }
html, body { margin: 0; padding: 0; background: #ffffff; color: #000000; }
body { font-family: 'Sarabun', 'Tahoma', sans-serif; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.page { position: relative; box-sizing: border-box; width: ${mm(g.pageW)}; height: ${mm(g.pageH)};
  padding: ${mm(g.margins.top + g.headerH)} ${mm(g.margins.right)} ${mm(g.margins.bottom + g.footerH)} ${mm(g.margins.left)};
  overflow: hidden; background: #ffffff; break-after: page; page-break-after: always; }
.page:last-child { break-after: auto; page-break-after: auto; }
.page-header { position: absolute; top: ${mm(g.margins.top)}; left: ${mm(g.margins.left)}; right: ${mm(g.margins.right)};
  height: ${mm(g.headerH)}; display: flex; justify-content: space-between; font-size: 3.2mm; color: #475569; }
.page-footer { position: absolute; bottom: ${mm(g.margins.bottom)}; left: 0; right: 0; height: ${mm(g.footerH)};
  display: flex; align-items: flex-end; justify-content: center; font-size: 3.2mm; color: #475569; }
table { table-layout: fixed; border-collapse: collapse; }
td { overflow: hidden; white-space: nowrap; text-overflow: clip; padding: 0 ${mm(CELL_PADDING_PX * PX_TO_MM)};
  vertical-align: bottom; box-sizing: border-box; ${settings.gridlines ? 'border: 0.2mm solid #c0c0c0;' : ''} }
`;
  return [
    '<!doctype html><html lang="th"><head><meta charset="utf-8">',
    `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; font-src data:; img-src data:">`,
    `<style>${css}</style></head><body>`,
    pagesHtml.join(''),
    '</body></html>',
  ].join('');
}
