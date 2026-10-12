import { Sheet, Cell } from './model';
import { cellToA1 } from './address';

export interface ClipboardData {
  tsv: string;
  html: string;
  cells: Record<string, Cell>;
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  isCut?: boolean;
}

/**
 * Formats a selected range into TSV text (tab-separated values)
 */
export function rangeToTsv(
  sheet: Sheet,
  range: { r1: number; c1: number; r2: number; c2: number },
  getDisplayVal: (r: number, c: number) => string
): string {
  const rows: string[] = [];

  for (let r = range.r1; r <= range.r2; r++) {
    const cols: string[] = [];
    for (let c = range.c1; c <= range.c2; c++) {
      cols.push(getDisplayVal(r, c));
    }
    rows.push(cols.join('\t'));
  }

  return rows.join('\n');
}

/**
 * Formats a selected range into HTML table for external paste
 */
export function rangeToHtml(
  sheet: Sheet,
  range: { r1: number; c1: number; r2: number; c2: number },
  getDisplayVal: (r: number, c: number) => string
): string {
  let html = '<table>\n';

  for (let r = range.r1; r <= range.r2; r++) {
    html += '  <tr>\n';
    for (let c = range.c1; c <= range.c2; c++) {
      const a1 = cellToA1(r, c);
      const cell = sheet.cells[a1];
      const val = getDisplayVal(r, c);
      const isBold = cell?.style?.bold ? ' style="font-weight:bold;"' : '';
      html += `    <td${isBold}>${val}</td>\n`;
    }
    html += '  </tr>\n';
  }

  html += '</table>';
  return html;
}

/**
 * Parses TSV string into 2D grid of values
 */
export function parseTsv(text: string): string[][] {
  if (!text) return [];
  const clean = text.replace(/\r/g, '').replace(/\n$/, '');
  if (!clean) return [];

  return clean.split('\n').map((row) => row.split('\t'));
}
