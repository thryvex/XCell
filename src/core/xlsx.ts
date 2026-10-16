import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import {
  Workbook,
  Sheet,
  Cell,
  CellStyle,
  NumFmt,
  MergeRange,
  createEmptySheet,
  CellValue,
  PrintSettings,
  PaperSize,
  PrintMargins,
  DEFAULT_PRINT_SETTINGS,
  DEFAULT_COL_WIDTH_PX,
  DEFAULT_ROW_HEIGHT_PX,
  MIN_COL_WIDTH_PX,
} from './model';
import { MARGINS_MM } from './paginate';
import { Evaluator } from './evaluator';
import { cellToA1, parseA1 } from './address';
import { dateToSerial, THAI_TO_ERROR_MAP } from './numfmt';
import { isCellError, FUNCTIONS } from './functions';
import { tokenize } from './tokenizer';
import { normalizeSheetRefsInFormula } from './refshift';

const NUM_FMT_EXCEL: Record<string, string> = {
  number2: '#,##0.00',
  integer: '#,##0',
  thousands: '#,##0.00',
  baht: '"฿"#,##0.00',
  percent: '0.00%',
  date: 'dd/mm/yyyy',
};

/**
 * ความกว้างคอลัมน์: 1 หน่วยของ .xlsx = 7 px (ตรงกับที่ Google Sheets แปลง)
 * ความสูงแถว: pt = px × 0.75; ทั้งสองค่าเขียนเป็นทศนิยม 2 ตำแหน่ง เพื่อให้เปิดกลับมาได้ px เท่าเดิม
 */
const PX_PER_WIDTH_UNIT = 7;
const PT_PER_PX = 0.75;
const round2 = (n: number) => Math.round(n * 100) / 100;

export function pxToColWidth(px: number): number {
  return round2(Math.max(MIN_COL_WIDTH_PX, px) / PX_PER_WIDTH_UNIT);
}
export function colWidthToPx(width: number): number {
  return Math.round(width * PX_PER_WIDTH_UNIT);
}
export function pxToRowHeight(px: number): number {
  return round2(px * PT_PER_PX);
}
export function rowHeightToPx(pt: number): number {
  return Math.round(pt / PT_PER_PX);
}

/** ชื่อไฟล์ในซิปที่บอกว่าไฟล์มีกราฟ (ดูแค่ชื่อไฟล์ ไม่อ่าน XML) */
const CHART_PART_PREFIX = 'xl/charts/';

/** ตรวจว่าไฟล์ .xlsx มีกราฟหรือไม่ จากรายชื่อไฟล์ในซิป */
export async function xlsxHasCharts(data: Uint8Array): Promise<boolean> {
  try {
    const zip = await JSZip.loadAsync(data);
    return Object.keys(zip.files).some((name) => name.startsWith(CHART_PART_PREFIX) && !zip.files[name].dir);
  } catch {
    return false;
  }
}

function argbToHex(argb?: string): string | undefined {
  if (!argb) return undefined;
  // ExcelJS returns ARGB e.g. "FFFF0000" or 6-digit hex
  const clean = argb.replace(/^#/g, '');
  if (clean.length === 8) {
    return '#' + clean.slice(2).toLowerCase();
  }
  if (clean.length === 6) {
    return '#' + clean.toLowerCase();
  }
  return undefined;
}

function hexToArgb(hex?: string): string | undefined {
  if (!hex) return undefined;
  const clean = hex.replace(/^#/g, '');
  if (clean.length === 6) {
    return 'FF' + clean.toUpperCase();
  }
  return clean.toUpperCase();
}

/**
 * Converts a XCell Workbook to an ExcelJS Workbook
 */
export async function workbookToExcelJS(wbModel: Workbook): Promise<ExcelJS.Workbook> {
  const excelWorkbook = new ExcelJS.Workbook();
  excelWorkbook.creator = 'XCell';
  excelWorkbook.lastModifiedBy = 'XCell';
  excelWorkbook.created = new Date();
  excelWorkbook.modified = new Date();

  const evaluator = new Evaluator(wbModel);

  const activeSheet = Math.max(0, Math.min(wbModel.sheets.length - 1, wbModel.activeSheet || 0));
  // ชีตที่เปิดอยู่ (activeTab นับรวมชีตซ่อนตามลำดับแท็บ)
  excelWorkbook.views = [
    { x: 0, y: 0, width: 20000, height: 12000, firstSheet: 0, activeTab: activeSheet, visibility: 'visible' },
  ];

  for (const [sheetIndex, sheetModel] of wbModel.sheets.entries()) {
    const ws = excelWorkbook.addWorksheet(sheetModel.name || 'Sheet1', {
      state: sheetModel.hidden ?? 'visible',
      // ค่าเริ่มต้นของคอลัมน์/แถวที่ไม่ได้ตั้ง (sheetFormatPr) ให้ตรงกับขนาดเริ่มต้นในแอป
      properties: {
        defaultColWidth: pxToColWidth(DEFAULT_COL_WIDTH_PX),
        defaultRowHeight: pxToRowHeight(DEFAULT_ROW_HEIGHT_PX),
      },
    });
    const tabSelected = sheetIndex === activeSheet;

    // Apply views (freezing + ชีตที่เลือก; tabSelected มีใน ExcelJS runtime แต่ไม่มีใน type)
    const isFrozen = sheetModel.frozenRows > 0 || sheetModel.frozenCols > 0;
    const zoomScale = sheetModel.zoom ?? 100; // ระดับซูมของชีต (4C)
    const view: Record<string, unknown> = isFrozen
      ? { state: 'frozen', xSplit: sheetModel.frozenCols, ySplit: sheetModel.frozenRows, tabSelected, zoomScale }
      : { state: 'normal', tabSelected, zoomScale };
    ws.views = [view as Partial<ExcelJS.WorksheetView>];

    // Apply column widths
    for (const [colIdxStr, widthPx] of Object.entries(sheetModel.colWidths)) {
      const colIdx = parseInt(colIdxStr, 10);
      const col = ws.getColumn(colIdx + 1);
      col.width = pxToColWidth(widthPx);
    }

    // Apply row heights
    for (const [rowIdxStr, heightPx] of Object.entries(sheetModel.rowHeights)) {
      const rowIdx = parseInt(rowIdxStr, 10);
      const row = ws.getRow(rowIdx + 1);
      row.height = pxToRowHeight(heightPx);
    }

    // Apply cells
    for (const [key, cellData] of Object.entries(sheetModel.cells)) {
      const p = parseA1(key);
      if (!p) continue;

      const excelCell = ws.getCell(p.r + 1, p.c + 1);
      const raw = cellData.raw;

      if (raw.startsWith('=')) {
        // Formula: replace aliases like MEAN -> AVERAGE for Google Sheets compatibility
        // และเขียนชื่อชีตในสูตรตามกติกา .xlsx (ใส่ '...' ยกเว้นชื่อ A-Z/ตัวเลข/ขีดล่างล้วน)
        const formulaStr = normalizeSheetRefsInFormula(raw).slice(1).replace(/\bMEAN\(/gi, 'AVERAGE(');
        let resultVal: any = undefined;

        if (cellData.unsupportedFormula && cellData.cachedValue !== undefined) {
          resultVal = cellData.cachedValue;
        } else {
          const computed = evaluator.getCellValue(key, sheetIndex);
          resultVal = typeof computed === 'number' || typeof computed === 'string' ? computed : undefined;

          if (typeof computed === 'string' && (isCellError(computed) || computed in THAI_TO_ERROR_MAP)) {
            let errCode = isCellError(computed) ? computed : THAI_TO_ERROR_MAP[computed];
            if (errCode === '#CYCLE!') errCode = '#REF!';
            if (errCode === '#ERR') errCode = '#VALUE!';
            resultVal = { error: errCode };
          }
        }

        excelCell.value = {
          formula: formulaStr,
          result: resultVal,
        };
      } else if (isCellError(raw) || raw in THAI_TO_ERROR_MAP) {
        let errCode = isCellError(raw) ? raw : THAI_TO_ERROR_MAP[raw];
        if (errCode === '#CYCLE!') errCode = '#REF!';
        if (errCode === '#ERR') errCode = '#VALUE!';
        excelCell.value = { error: errCode as any };
      } else {
        const cleaned = raw.replace(/,/g, '');
        if (cleaned !== '' && !isNaN(Number(cleaned))) {
          excelCell.value = Number(cleaned);
        } else {
          excelCell.value = raw.startsWith("'") ? raw.slice(1) : raw;
        }
      }

      // Number formatting
      if (cellData.fmt && NUM_FMT_EXCEL[cellData.fmt]) {
        excelCell.numFmt = NUM_FMT_EXCEL[cellData.fmt];
      }

      // Styles
      if (cellData.style) {
        const st = cellData.style;
        const font: Partial<ExcelJS.Font> = {};
        if (st.bold !== undefined) font.bold = st.bold;
        if (st.italic !== undefined) font.italic = st.italic;
        if (st.underline !== undefined) font.underline = st.underline;
        if (st.color) font.color = { argb: hexToArgb(st.color) };
        if (st.fontSize !== undefined) font.size = st.fontSize;
        if (st.fontFamily) font.name = st.fontFamily;
        if (Object.keys(font).length > 0) excelCell.font = font;

        if (st.fill) {
          excelCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: hexToArgb(st.fill)! },
          };
        }

        if (st.align || st.wrap) {
          excelCell.alignment = {
            horizontal: st.align,
            wrapText: st.wrap,
          };
        }

        if (st.border && st.border !== 'none') {
          const borderStyle: ExcelJS.BorderStyle = 'thin';
          excelCell.border = {
            top: { style: borderStyle },
            left: { style: borderStyle },
            bottom: { style: borderStyle },
            right: { style: borderStyle },
          };
        }
      }
    }

    // ตัวเลือกพิมพ์ (5B): เก็บใน pageSetup ของแต่ละชีต (ไม่เก็บ "ช่วงที่เลือก")
    if (wbModel.printSettings) {
      writePageSetup(ws, wbModel.printSettings, sheetModel.frozenRows);
    }

    // ตัวกรอง: เขียนช่วง autoFilter และสถานะแถวที่ซ่อนตามจริง (เงื่อนไขกรองไม่ถูกเขียน)
    if (sheetModel.filter) {
      const f = sheetModel.filter;
      ws.autoFilter = {
        from: { row: f.r1 + 1, column: f.c1 + 1 },
        to: { row: f.r2 + 1, column: f.c2 + 1 },
      };
    }
    for (const r of sheetModel.hiddenRows ?? []) {
      ws.getRow(r + 1).hidden = true;
    }

    // Apply merges
    for (const m of sheetModel.merges) {
      ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    }
  }

  return excelWorkbook;
}

/**
 * Detects features in an ExcelJS Workbook that are unsupported by XCell (SPEC 6.4)
 */
export function detectUnsupportedFeatures(excelWorkbook: ExcelJS.Workbook): string[] {
  const features = new Set<string>();

  if ((excelWorkbook as any).pivotTables && (excelWorkbook as any).pivotTables.length > 0) {
    features.add('ตาราง Pivot');
  }

  excelWorkbook.eachSheet((ws) => {
    // ชีตซ่อน: รองรับแล้วใน 4B (เก็บไว้ตามเดิม) จึงไม่นับเป็นส่วนที่ไม่รองรับ

    // 2. Images
    if (
      (typeof ws.getImages === 'function' && ws.getImages().length > 0) ||
      ((ws as any).media && (ws as any).media.length > 0)
    ) {
      features.add('รูปภาพ');
    }

    // 3. Charts / Drawings
    if (
      ((ws as any).drawings && (ws as any).drawings.length > 0) ||
      ((ws as any)._drawings && (ws as any)._drawings.length > 0) ||
      Boolean((ws.model as any)?.drawing)
    ) {
      features.add('แผนภูมิ/กราฟ');
    }

    // 4. Pivot tables
    if (
      ((ws as any).pivotTables && (ws as any).pivotTables.length > 0) ||
      Boolean((ws.model as any)?.pivotTables && (ws.model as any).pivotTables.length > 0)
    ) {
      features.add('ตาราง Pivot');
    }

    // 5. Conditional formatting
    if (
      ((ws as any).conditionalFormattings && (ws as any).conditionalFormattings.length > 0) ||
      Boolean((ws.model as any)?.conditionalFormatting && (ws.model as any).conditionalFormatting.length > 0)
    ) {
      features.add('การจัดรูปแบบตามเงื่อนไข (Conditional Formatting)');
    }

    // 6. Data validation
    const dvModel = (ws as any).dataValidations?.model || (ws.model as any)?.dataValidations;
    if (dvModel && Object.keys(dvModel).length > 0) {
      features.add('การตรวจสอบความถูกต้องของข้อมูล (Data Validation)');
    }
  });

  return Array.from(features);
}

/**
 * Checks if a formula string contains any functions not supported by XCell
 */
export function hasUnsupportedFunction(formulaStr: string): boolean {
  if (!formulaStr.startsWith('=')) return false;
  try {
    const tokens = tokenize(formulaStr);
    for (const t of tokens) {
      if (t.type === 'FUNCTION') {
        const fnName = t.value.toUpperCase();
        if (!FUNCTIONS[fnName] && fnName !== 'MEAN') {
          return true;
        }
      }
    }
  } catch {
    return true;
  }
  return false;
}

/**
 * Converts an ExcelJS Workbook to a XCell Workbook model
 */
export async function excelJSToWorkbook(excelWorkbook: ExcelJS.Workbook): Promise<Workbook> {
  const sheets: Sheet[] = [];
  const unsupportedFeatures = detectUnsupportedFeatures(excelWorkbook);

  excelWorkbook.eachSheet((ws, sheetId) => {
    const sheet: Sheet = createEmptySheet(ws.name || `Sheet${sheetId}`);

    // Hidden sheets: เก็บข้อมูลและสถานะซ่อนไว้ตามเดิม
    if (ws.state === 'hidden' || ws.state === 'veryHidden') {
      sheet.hidden = ws.state;
    }

    // Frozen panes
    const view = ws.views?.find((v) => v.state === 'frozen') as
      | { xSplit?: number; ySplit?: number }
      | undefined;
    if (view) {
      sheet.frozenCols = view.xSplit || 0;
      sheet.frozenRows = view.ySplit || 0;
    }

    // ระดับซูมของชีต (zoomScale) ปัดเป็นขั้น 10% ในช่วง 50–200%
    const zoomScale = (ws.views?.[0] as { zoomScale?: number } | undefined)?.zoomScale;
    if (typeof zoomScale === 'number' && zoomScale !== 100) {
      const z = Math.max(50, Math.min(200, Math.round(zoomScale / 10) * 10));
      if (z !== 100) sheet.zoom = z;
    }

    // Column widths / row heights (กลับด้านกับตอนเขียน: width × 7 px, height pt ÷ 0.75)
    ws.columns?.forEach((col, idx) => {
      if (col && typeof col.width === 'number' && (col as any).isCustomWidth !== false) {
        sheet.colWidths[idx] = colWidthToPx(col.width);
      }
    });
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (typeof row.height === 'number' && row.height > 0) {
        sheet.rowHeights[rowNumber - 1] = rowHeightToPx(row.height);
      }
    });

    // Iterate rows and cells
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const a1 = cellToA1(rowNumber - 1, colNumber - 1);
        let raw = '';
        let fmt: NumFmt | undefined = undefined;
        const style: CellStyle = {};
        let isUnsupportedFormula = false;
        let cachedVal: CellValue | undefined = undefined;

        // Parse value & formula
        if (cell.type === ExcelJS.ValueType.Formula) {
          const f = cell.formula;
          raw = f ? `=${f}` : '';

          // Extract cached result
          const res = cell.result;
          if (res !== null && res !== undefined) {
            if (typeof res === 'number' || typeof res === 'string' || typeof res === 'boolean') {
              cachedVal = res;
            } else if (typeof res === 'object' && 'error' in res) {
              cachedVal = (res as any).error;
            }
          }

          if (hasUnsupportedFunction(raw)) {
            isUnsupportedFormula = true;
          }
        } else if (
          cell.type === ExcelJS.ValueType.Error ||
          (cell.value && typeof cell.value === 'object' && 'error' in cell.value)
        ) {
          const errVal = (cell.value as any)?.error;
          raw = typeof errVal === 'string' ? errVal : '#ERR';
        } else if (cell.type === ExcelJS.ValueType.Date) {
          const d = cell.value as Date;
          if (d instanceof Date && !isNaN(d.getTime())) {
            raw = String(dateToSerial(d));
            fmt = 'date';
          }
        } else if (cell.value !== null && cell.value !== undefined) {
          raw = String(cell.value);
        }

        // Parse number format
        const nf = cell.numFmt || '';
        if (/฿|THB/.test(nf)) fmt = 'baht';
        else if (nf.includes('%')) fmt = 'percent';
        else if (nf.includes('0.00')) fmt = 'number2';
        else if (/dd|mm|yyyy|yy/i.test(nf)) fmt = 'date';

        // Parse styles
        if (cell.font) {
          if (cell.font.bold) style.bold = true;
          if (cell.font.italic) style.italic = true;
          if (cell.font.underline) style.underline = true;
          if (cell.font.color?.argb) {
            style.color = argbToHex(cell.font.color.argb);
          }
          if (cell.font.size) style.fontSize = cell.font.size;
          if (cell.font.name) style.fontFamily = cell.font.name;
        }

        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor?.argb) {
          style.fill = argbToHex(cell.fill.fgColor.argb);
        }

        if (cell.alignment) {
          if (cell.alignment.horizontal) {
            style.align = cell.alignment.horizontal as any;
          }
          if (cell.alignment.wrapText) {
            style.wrap = true;
          }
        }

        if (cell.border && (cell.border.top || cell.border.bottom || cell.border.left || cell.border.right)) {
          style.border = 'all';
        }

        const cellRecord: Cell = { raw };
        if (isUnsupportedFormula) {
          cellRecord.unsupportedFormula = true;
          cellRecord.cachedValue = cachedVal;
        }
        if (fmt) cellRecord.fmt = fmt;
        if (Object.keys(style).length > 0) cellRecord.style = style;

        sheet.cells[a1] = cellRecord;
      });
    });

    // Parse merges
    // ws.model.merges returns array of range strings e.g. ["A1:B2"]
    const rawMerges = (ws.model as any)?.merges || [];
    for (const rangeStr of rawMerges) {
      const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(rangeStr);
      if (match) {
        const p1 = parseA1(`${match[1]}${match[2]}`);
        const p2 = parseA1(`${match[3]}${match[4]}`);
        if (p1 && p2) {
          sheet.merges.push({
            r1: Math.min(p1.r, p2.r),
            c1: Math.min(p1.c, p2.c),
            r2: Math.max(p1.r, p2.r),
            c2: Math.max(p1.c, p2.c),
          });
        }
      }
    }

    // ตัวกรอง: .xlsx (ผ่าน ExcelJS) เก็บได้แค่ช่วง autoFilter ไม่มีเงื่อนไขกรอง
    const af = ws.autoFilter as ExcelJS.AutoFilter | undefined;
    const afRange = parseAutoFilterRange(af);
    if (afRange) sheet.filter = { ...afRange, columns: {} };
    // แถวที่ซ่อน (จากตัวกรอง)
    const hiddenRows: number[] = [];
    ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      if (row.hidden) hiddenRows.push(rowNumber - 1);
    });
    if (hiddenRows.length > 0) sheet.hiddenRows = hiddenRows;

    sheets.push(sheet);
  });

  const activeTab = (excelWorkbook.views?.[0] as { activeTab?: number } | undefined)?.activeTab;
  const activeSheet =
    typeof activeTab === 'number' && activeTab >= 0 && activeTab < sheets.length ? activeTab : 0;

  return {
    sheets: sheets.length > 0 ? sheets : [createEmptySheet('Sheet1')],
    activeSheet,
    unsupportedFeatures: unsupportedFeatures.length > 0 ? unsupportedFeatures : undefined,
    printSettings: readPrintSettings(excelWorkbook.worksheets[activeSheet], sheets[activeSheet]),
  };
}

/**
 * Serializes Workbook model to xlsx Uint8Array buffer
 */
export async function exportToXlsxBuffer(wb: Workbook): Promise<Uint8Array> {
  const excelWb = await workbookToExcelJS(wb);
  const buffer = await excelWb.xlsx.writeBuffer();
  return new Uint8Array(buffer);
}

/**
 * Parses xlsx Uint8Array buffer into Workbook model
 */
export async function importFromXlsxBuffer(data: Uint8Array): Promise<Workbook> {
  const excelWb = new ExcelJS.Workbook();
  // ExcelJS รับ Uint8Array ได้จริง แต่ type ประกาศเป็น Buffer ของ Node
  await excelWb.xlsx.load(data as unknown as Parameters<typeof excelWb.xlsx.load>[0]);
  const wb = await excelJSToWorkbook(excelWb);
  // ExcelJS ทิ้งกราฟตั้งแต่ตอนอ่าน จึงตรวจจากชื่อไฟล์ในซิปแทน และเตือนเรื่องกราฟแยกจากรายการอื่น
  if (await xlsxHasCharts(data)) {
    wb.sourceHasCharts = true;
    const others = (wb.unsupportedFeatures ?? []).filter((f) => f !== 'แผนภูมิ/กราฟ');
    wb.unsupportedFeatures = others.length > 0 ? others : undefined;
  }
  return wb;
}

/** แปลงค่า autoFilter ของ ExcelJS ("A1:C10" หรือ {from, to}) เป็นช่วง 0-based */
function parseAutoFilterRange(
  af: ExcelJS.AutoFilter | undefined
): { r1: number; c1: number; r2: number; c2: number } | null {
  if (!af) return null;
  const toCoord = (v: string | { row: number; column: number }) => {
    if (typeof v === 'string') {
      const p = parseA1(v.replace(/\$/g, ''));
      return p ? { r: p.r, c: p.c } : null;
    }
    return { r: v.row - 1, c: v.column - 1 };
  };
  let from: { r: number; c: number } | null;
  let to: { r: number; c: number } | null;
  if (typeof af === 'string') {
    const [a, b] = af.split(':');
    from = toCoord(a);
    to = toCoord(b ?? a);
  } else {
    from = toCoord(af.from);
    to = toCoord(af.to);
  }
  if (!from || !to) return null;
  return {
    r1: Math.min(from.r, to.r),
    c1: Math.min(from.c, to.c),
    r2: Math.max(from.r, to.r),
    c2: Math.max(from.c, to.c),
  };
}

// ---------------------------------------------------------------------------
// ตัวเลือกพิมพ์ ↔ ExcelJS pageSetup (5B)
// ---------------------------------------------------------------------------

const PAPER_CODES: Record<PaperSize, number> = { Letter: 1, A4: 9, A5: 11 };
const MM_PER_INCH = 25.4;

function writePageSetup(ws: ExcelJS.Worksheet, ps: PrintSettings, frozenRows: number) {
  const m = MARGINS_MM[ps.margins];
  ws.pageSetup.orientation = ps.orientation;
  ws.pageSetup.paperSize = PAPER_CODES[ps.paper] as ExcelJS.PaperSize;
  if (ps.scaleMode === 'fitWidth') {
    ws.pageSetup.fitToPage = true;
    ws.pageSetup.fitToWidth = 1;
    ws.pageSetup.fitToHeight = 0;
  } else {
    ws.pageSetup.fitToPage = false;
    ws.pageSetup.scale = ps.scaleMode === 'custom' ? Math.round(ps.scalePercent) : 100;
  }
  ws.pageSetup.margins = {
    top: m.top / MM_PER_INCH,
    bottom: m.bottom / MM_PER_INCH,
    left: m.left / MM_PER_INCH,
    right: m.right / MM_PER_INCH,
    header: 0.3,
    footer: 0.3,
  };
  ws.pageSetup.showGridLines = ps.gridlines;
  if (ps.repeatFrozen && frozenRows > 0) {
    ws.pageSetup.printTitlesRow = `1:${frozenRows}`;
  }
}

function nearestMargins(left: number | undefined, top: number | undefined): PrintMargins {
  if (typeof left !== 'number' || typeof top !== 'number') return 'normal';
  const leftMm = left * MM_PER_INCH;
  const topMm = top * MM_PER_INCH;
  let best: PrintMargins = 'normal';
  let bestDiff = Infinity;
  for (const key of Object.keys(MARGINS_MM) as PrintMargins[]) {
    const diff = Math.abs(MARGINS_MM[key].left - leftMm) + Math.abs(MARGINS_MM[key].top - topMm);
    if (diff < bestDiff) {
      bestDiff = diff;
      best = key;
    }
  }
  return best;
}

/** อ่านตัวเลือกพิมพ์จาก pageSetup ของชีตที่เปิดอยู่ (เท่าที่ .xlsx เก็บได้) */
function readPrintSettings(ws: ExcelJS.Worksheet | undefined, sheet: Sheet | undefined): PrintSettings {
  const ps: PrintSettings = { ...DEFAULT_PRINT_SETTINGS };
  if (!ws) return ps;
  const setup = ws.pageSetup ?? {};
  const paper = (Object.keys(PAPER_CODES) as PaperSize[]).find((k) => PAPER_CODES[k] === setup.paperSize);
  if (paper) ps.paper = paper;
  if (setup.orientation === 'landscape') ps.orientation = 'landscape';
  if (setup.fitToPage && setup.fitToWidth === 1) {
    ps.scaleMode = 'fitWidth';
  } else if (typeof setup.scale === 'number' && setup.scale !== 100) {
    ps.scaleMode = 'custom';
    ps.scalePercent = Math.max(30, Math.min(200, setup.scale));
  } else {
    ps.scaleMode = '100';
  }
  ps.margins = nearestMargins(setup.margins?.left, setup.margins?.top);
  ps.gridlines = setup.showGridLines === true;
  // มีแถวตรึงแต่ไฟล์ไม่ได้ตั้งแถวพิมพ์ซ้ำ → ไม่พิมพ์ซ้ำ
  ps.repeatFrozen = Boolean(setup.printTitlesRow) || !((sheet?.frozenRows ?? 0) > 0);
  return ps;
}
