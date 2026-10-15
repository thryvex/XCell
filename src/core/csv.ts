import { Sheet, Cell, NumFmt, CellValue, createEmptySheet } from './model';
import { Evaluator } from './evaluator';
import { cellToA1, parseA1 } from './address';
import { formatDisplayValue, validDateSerial, toGregorianYear } from './numfmt';
import { isCellError } from './functions';
import { FORBIDDEN_SHEET_NAME_CHARS } from './sheetName';

// ---------------------------------------------------------------------------
// ส่งออก CSV
// ---------------------------------------------------------------------------

/** ครอบด้วย " เมื่อมี comma, " หรือขึ้นบรรทัดใหม่ และเขียน " เป็น "" */
export function quoteCsvField(value: string, delimiter = ','): string {
  if (value.includes(delimiter) || value.includes('"') || value.includes('\n') || value.includes('\r')) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

/** ตาราง → ข้อความ CSV (CRLF) */
export function writeCsv(rows: string[][], delimiter = ','): string {
  return rows.map((row) => row.map((v) => quoteCsvField(v, delimiter)).join(delimiter)).join('\r\n');
}

/** ข้อความ → ไบต์ UTF-8 พร้อม BOM */
export function encodeUtf8WithBom(text: string): Uint8Array {
  const body = new TextEncoder().encode(text);
  const out = new Uint8Array(body.length + 3);
  out.set([0xef, 0xbb, 0xbf], 0);
  out.set(body, 3);
  return out;
}

/** ค่าที่แสดงสำหรับ CSV: error ใช้รหัสมาตรฐาน (#DIV/0! ฯลฯ) ไม่ใช่ชื่อไทย */
function csvDisplayValue(value: CellValue, fmt: NumFmt | undefined): string {
  if (isCellError(value)) {
    if (value === '#CYCLE!') return '#REF!';
    if (value === '#ERR') return '#ERROR!';
    return value;
  }
  return formatDisplayValue(value, fmt);
}

/**
 * ชีต → ตารางข้อความสำหรับ CSV
 * ขอบเขต A1 ถึงเซลสุดท้ายที่มีข้อมูล, เซลที่ถูกผสานทับเป็นค่าว่าง, รวมแถวที่ซ่อนด้วยตัวกรอง
 */
export function sheetToCsvRows(sheet: Sheet, evaluator: Evaluator, sheetIndex: number): string[][] {
  let maxR = -1;
  let maxC = -1;
  for (const [key, cell] of Object.entries(sheet.cells)) {
    if (!cell.raw) continue;
    const p = parseA1(key);
    if (!p) continue;
    maxR = Math.max(maxR, p.r);
    maxC = Math.max(maxC, p.c);
  }
  if (maxR < 0) return [];

  const covered = new Set<string>();
  for (const m of sheet.merges) {
    for (let r = m.r1; r <= m.r2; r++) {
      for (let c = m.c1; c <= m.c2; c++) {
        if (r !== m.r1 || c !== m.c1) covered.add(cellToA1(r, c));
      }
    }
  }

  const rows: string[][] = [];
  for (let r = 0; r <= maxR; r++) {
    const row: string[] = [];
    for (let c = 0; c <= maxC; c++) {
      const key = cellToA1(r, c);
      const cell = sheet.cells[key];
      if (!cell?.raw || covered.has(key)) {
        row.push('');
        continue;
      }
      row.push(csvDisplayValue(evaluator.getCellValue(key, sheetIndex), cell.fmt));
    }
    rows.push(row);
  }
  return rows;
}

// ---------------------------------------------------------------------------
// อ่าน CSV
// ---------------------------------------------------------------------------

export type CsvEncoding = 'utf-8' | 'windows-874' | 'utf-16le' | 'utf-16be';

/** เดา encoding: BOM ก่อน, ไม่มี BOM ลอง UTF-8 แบบเข้มงวด ถ้าไม่ผ่านใช้ Windows-874 */
export function detectEncoding(bytes: Uint8Array): CsvEncoding {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return 'utf-8';
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return 'utf-16le';
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return 'utf-16be';
  try {
    new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    return 'utf-8';
  } catch {
    return 'windows-874';
  }
}

/** ถอดรหัสไบต์เป็นข้อความ (ตัด BOM ออก) — ไม่ throw: ไบต์ที่ผิดกลายเป็นอักขระแทน */
export function decodeBytes(bytes: Uint8Array, encoding: CsvEncoding): string {
  const text = new TextDecoder(encoding).decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/** มีไบต์ 0 ทั้งที่ไม่ใช่ UTF-16 → น่าจะไม่ใช่ไฟล์ข้อความ (เช่น .xlsx ที่เปลี่ยนนามสกุล) */
export function looksBinary(bytes: Uint8Array, encoding: CsvEncoding): boolean {
  if (encoding === 'utf-16le' || encoding === 'utf-16be') return false;
  const limit = Math.min(bytes.length, 4096);
  for (let i = 0; i < limit; i++) if (bytes[i] === 0) return true;
  return false;
}

/**
 * แยก CSV เป็นตาราง: รองรับ " ครอบเซล, "" ในเซล, ขึ้นบรรทัดใหม่ในเซล, CRLF / LF / CR
 * บรรทัดว่างท้ายไฟล์ไม่นับเป็นแถว
 */
export function parseCsv(text: string, delimiter: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let inQuotes = false;
  let i = 0;
  const n = text.length;

  while (i < n) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      field += ch;
      i++;
      continue;
    }

    if (ch === '"' && field === '') {
      inQuotes = true;
      i++;
      continue;
    }
    if (delimiter && text.startsWith(delimiter, i)) {
      row.push(field);
      field = '';
      i += delimiter.length;
      continue;
    }
    if (ch === '\r' || ch === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    field += ch;
    i++;
  }

  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

export const DELIMITER_CANDIDATES = [',', ';', '\t', '|'];

/** เดาตัวคั่นจากตัวอย่าง: ตัวที่แยกได้หลายคอลัมน์และจำนวนคอลัมน์สม่ำเสมอที่สุด */
export function detectDelimiter(text: string): string {
  const sample = text.slice(0, 20000);
  let best = ',';
  let bestScore = -1;
  for (const d of DELIMITER_CANDIDATES) {
    const rows = parseCsv(sample, d).slice(0, 20).filter((r) => r.length > 1 || r[0] !== '');
    if (rows.length === 0) continue;
    const counts = rows.map((r) => r.length);
    const first = counts[0];
    if (first <= 1) continue;
    const consistent = counts.filter((c) => c === first).length / counts.length;
    const score = consistent * 1000 + first;
    if (score > bestScore) {
      bestScore = score;
      best = d;
    }
  }
  return best;
}

// ---------------------------------------------------------------------------
// แปลงค่าเซล
// ---------------------------------------------------------------------------

const NUMBER_RE = /^-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$/;
const PERCENT_RE = /^(-?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)%$/;
const DMY_RE = /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/;
const YMD_RE = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;

function digitCount(text: string): number {
  return (text.match(/\d/g) ?? []).length;
}

/** เลขที่ต้องคงเป็นข้อความ: ขึ้นต้นด้วย 0 และยาวเกิน 1 หลัก (ไม่ใช่ 0.x) หรือเกิน 15 หลัก */
function mustStayText(text: string): boolean {
  const unsigned = text.replace(/^-/, '');
  if (/^0\d/.test(unsigned)) return true;
  return digitCount(text) > 15;
}

/** ข้อความที่ต้องเก็บเป็นข้อความล้วน: ใส่ ' นำหน้า (ไม่แสดง) กันการตีความเป็นสูตร/ตัวเลข */
function asText(text: string): Cell {
  return { raw: `'${text}` };
}

/** ข้อความล้วนที่หน้าตาไม่เหมือนตัวเลข/สูตร เก็บตรงๆ ได้ */
function plainText(text: string): Cell {
  if (text.startsWith('=') || text.startsWith("'") || NUMBER_RE.test(text.trim())) return asText(text);
  return { raw: text };
}

/**
 * แปลงข้อความหนึ่งเซลจาก CSV เป็น Cell (null = เซลว่าง)
 * ข้อความที่ขึ้นต้นด้วย "=" เป็นข้อความธรรมดาเสมอ ไม่ใช่สูตร
 */
export function convertCsvValue(text: string, autoConvert: boolean): Cell | null {
  if (text === '') return null;
  if (!autoConvert) return plainText(text);

  const t = text.trim();
  if (t.startsWith('=')) return asText(text);

  if (NUMBER_RE.test(t)) {
    if (mustStayText(t)) return asText(text);
    const value = Number(t.replace(/,/g, ''));
    let fmt: NumFmt | undefined;
    if (t.includes(',')) fmt = t.includes('.') ? 'number2' : 'integer';
    return fmt ? { raw: String(value), fmt } : { raw: String(value) };
  }

  const pct = PERCENT_RE.exec(t);
  if (pct) {
    if (mustStayText(pct[1])) return asText(text);
    const value = Number(pct[1].replace(/,/g, '')) / 100;
    return { raw: String(parseFloat(value.toPrecision(15))), fmt: 'percent' };
  }

  const dmy = DMY_RE.exec(t);
  if (dmy) {
    // ปี >= 2400 เป็น พ.ศ. (กติกาเดียวกับการพิมพ์เข้า)
    const serial = validDateSerial(toGregorianYear(Number(dmy[3])), Number(dmy[2]), Number(dmy[1]));
    if (serial !== null) return { raw: String(serial), fmt: 'date' };
  }
  const ymd = YMD_RE.exec(t);
  if (ymd) {
    const serial = validDateSerial(toGregorianYear(Number(ymd[1])), Number(ymd[2]), Number(ymd[3]));
    if (serial !== null) return { raw: String(serial), fmt: 'date' };
  }

  return plainText(text);
}

// ---------------------------------------------------------------------------
// สร้างชีตจาก CSV
// ---------------------------------------------------------------------------

export const CSV_MAX_ROWS = 10000;
export const CSV_MAX_COLS = 702; // A–ZZ

export function csvSize(rows: string[][]): { rows: number; cols: number } {
  return { rows: rows.length, cols: rows.reduce((m, r) => Math.max(m, r.length), 0) };
}

export function csvExceedsLimits(rows: string[][]): boolean {
  const size = csvSize(rows);
  return size.rows > CSV_MAX_ROWS || size.cols > CSV_MAX_COLS;
}

/** ตาราง CSV → ชีต (ตัดเฉพาะส่วนที่เกินขีดจำกัด เมื่อผู้ใช้ยืนยันแล้ว) */
export function buildSheetFromCsv(rows: string[][], name: string, autoConvert: boolean): Sheet {
  const sheet = createEmptySheet(name);
  const maxRows = Math.min(rows.length, CSV_MAX_ROWS);
  for (let r = 0; r < maxRows; r++) {
    const row = rows[r];
    const maxCols = Math.min(row.length, CSV_MAX_COLS);
    for (let c = 0; c < maxCols; c++) {
      const cell = convertCsvValue(row[c], autoConvert);
      if (cell) sheet.cells[cellToA1(r, c)] = cell;
    }
  }
  return sheet;
}

/** ชื่อชีตจากชื่อไฟล์ ตามกติกาชื่อชีต (ตัดอักขระต้องห้าม, ไม่เกิน 31, ไม่ว่าง, ไม่ซ้ำ) */
export function sheetNameFromFileName(fileName: string, existingNames: string[] = []): string {
  const base = fileName.split(/[\\/]/).pop()!.replace(/\.[^.]+$/, '');
  let name = base;
  for (const ch of FORBIDDEN_SHEET_NAME_CHARS) name = name.split(ch).join('');
  name = name.trim().slice(0, 31).trim();
  if (!name) name = 'ชีต1';

  const taken = new Set(existingNames.map((n) => n.toLowerCase()));
  if (!taken.has(name.toLowerCase())) return name;
  for (let i = 2; i < 1000; i++) {
    const suffix = ` (${i})`;
    const candidate = name.slice(0, 31 - suffix.length) + suffix;
    if (!taken.has(candidate.toLowerCase())) return candidate;
  }
  return name;
}
