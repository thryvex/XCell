import ExcelJS from 'exceljs';
import { Workbook, Sheet, Cell, CellStyle, NumFmt, MergeRange, createEmptySheet } from './model';
import { Evaluator } from './evaluator';
import { cellToA1, parseA1 } from './address';
import { dateToSerial, THAI_TO_ERROR_MAP } from './numfmt';
import { isCellError } from './functions';

const NUM_FMT_EXCEL: Record<string, string> = {
  number2: '#,##0.00',
  integer: '#,##0',
  thousands: '#,##0.00',
  baht: '"฿"#,##0.00',
  percent: '0.00%',
  date: 'dd/mm/yyyy',
};

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
 * Converts a SheetEasy Workbook to an ExcelJS Workbook
 */
export async function workbookToExcelJS(wbModel: Workbook): Promise<ExcelJS.Workbook> {
  const excelWorkbook = new ExcelJS.Workbook();
  excelWorkbook.creator = 'SheetEasy';
  excelWorkbook.lastModifiedBy = 'SheetEasy';
  excelWorkbook.created = new Date();
  excelWorkbook.modified = new Date();

  for (const sheetModel of wbModel.sheets) {
    const ws = excelWorkbook.addWorksheet(sheetModel.name || 'Sheet1');
    const evaluator = new Evaluator(sheetModel);

    // Apply views (freezing)
    if (sheetModel.frozenRows > 0 || sheetModel.frozenCols > 0) {
      ws.views = [
        {
          state: 'frozen',
          xSplit: sheetModel.frozenCols,
          ySplit: sheetModel.frozenRows,
        },
      ];
    }

    // Apply column widths
    for (const [colIdxStr, widthPx] of Object.entries(sheetModel.colWidths)) {
      const colIdx = parseInt(colIdxStr, 10);
      const col = ws.getColumn(colIdx + 1);
      // Rough conversion: 1 Excel width unit ~ 7.5 - 8 px
      col.width = Math.max(8, Math.round(widthPx / 8));
    }

    // Apply row heights
    for (const [rowIdxStr, heightPx] of Object.entries(sheetModel.rowHeights)) {
      const rowIdx = parseInt(rowIdxStr, 10);
      const row = ws.getRow(rowIdx + 1);
      // 1 pt ~ 1.33 px
      row.height = Math.round(heightPx * 0.75);
    }

    // Apply cells
    for (const [key, cellData] of Object.entries(sheetModel.cells)) {
      const p = parseA1(key);
      if (!p) continue;

      const excelCell = ws.getCell(p.r + 1, p.c + 1);
      const raw = cellData.raw;

      if (raw.startsWith('=')) {
        // Formula: replace aliases like MEAN -> AVERAGE for Google Sheets compatibility
        const formulaStr = raw.slice(1).replace(/\bMEAN\(/gi, 'AVERAGE(');
        const computed = evaluator.getCellValue(key);
        let resultVal: any = typeof computed === 'number' || typeof computed === 'string' ? computed : undefined;

        if (typeof computed === 'string' && (isCellError(computed) || computed in THAI_TO_ERROR_MAP)) {
          let errCode = isCellError(computed) ? computed : THAI_TO_ERROR_MAP[computed];
          if (errCode === '#CYCLE!') errCode = '#REF!';
          if (errCode === '#ERR') errCode = '#VALUE!';
          resultVal = { error: errCode };
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

    // Apply merges
    for (const m of sheetModel.merges) {
      ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    }
  }

  return excelWorkbook;
}

/**
 * Converts an ExcelJS Workbook to a SheetEasy Workbook model
 */
export async function excelJSToWorkbook(excelWorkbook: ExcelJS.Workbook): Promise<Workbook> {
  const sheets: Sheet[] = [];

  excelWorkbook.eachSheet((ws, sheetId) => {
    const sheet: Sheet = createEmptySheet(ws.name || `Sheet${sheetId}`);

    // Frozen panes
    const view = ws.views?.find((v) => v.state === 'frozen');
    if (view) {
      sheet.frozenCols = view.xSplit || 0;
      sheet.frozenRows = view.ySplit || 0;
    }

    // Iterate rows and cells
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        const a1 = cellToA1(rowNumber - 1, colNumber - 1);
        let raw = '';
        let fmt: NumFmt | undefined = undefined;
        const style: CellStyle = {};

        // Parse value & formula
        if (cell.type === ExcelJS.ValueType.Formula) {
          const f = cell.formula;
          raw = f ? `=${f}` : '';
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

    sheets.push(sheet);
  });

  return {
    sheets: sheets.length > 0 ? sheets : [createEmptySheet('Sheet1')],
    activeSheet: 0,
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
  await excelWb.xlsx.load(data);
  return await excelJSToWorkbook(excelWb);
}
