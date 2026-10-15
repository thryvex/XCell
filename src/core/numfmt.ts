import { CellValue, NumFmt, CellError } from './model';
import { th } from '../i18n/th';

// Excel Epoch: Dec 30, 1899 UTC
const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30)).getTime();
const MS_PER_DAY = 24 * 60 * 60 * 1000;

export function dateToSerial(date: Date): number {
  const utcDate = Date.UTC(date.getFullYear(), date.getMonth(), date.getDate());
  return Math.round((utcDate - EXCEL_EPOCH) / MS_PER_DAY);
}

export function serialToDate(serial: number): Date {
  const utcTime = EXCEL_EPOCH + Math.round(serial) * MS_PER_DAY;
  return new Date(utcTime);
}

// ---------------------------------------------------------------------------
// ปี พ.ศ. (5A): ข้อมูลเก็บเป็น serial ค.ศ. เสมอ พ.ศ. เป็นเรื่องการแสดงผลและการพิมพ์เข้าเท่านั้น
// ---------------------------------------------------------------------------

export type YearMode = 'ce' | 'be';
export const BE_OFFSET = 543;
/** ปีตัวเลข 4 หลักที่ >= 2400 ถือเป็น พ.ศ. ไม่ว่าตั้งค่าโหมดไหน */
export const BE_YEAR_THRESHOLD = 2400;

let displayYearMode: YearMode = 'ce';

/** ตั้งโหมดแสดงปีของแอป (มีผลกับ formatDate / formatDisplayValue เท่านั้น ไม่กระทบสูตร) */
export function setDisplayYearMode(mode: YearMode) {
  displayYearMode = mode;
}

export function getDisplayYearMode(): YearMode {
  return displayYearMode;
}

/** ปีที่ผู้ใช้พิมพ์ → ปี ค.ศ. (>= 2400 = พ.ศ. ลบ 543) */
export function toGregorianYear(year: number): number {
  return year >= BE_YEAR_THRESHOLD ? year - BE_OFFSET : year;
}

/** serial ของวันนี้ (ไม่มีเวลา) */
export function todaySerial(): number {
  return dateToSerial(new Date());
}

/** serial จาก ปี/เดือน/วัน (ค.ศ.) ถ้าเป็นวันที่ที่มีจริง ไม่เช่นนั้น null (เช่น 29/2 ของปีที่ไม่ใช่ปีอธิกสุรทิน) */
export function validDateSerial(year: number, month: number, day: number): number | null {
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }
  return dateToSerial(date);
}

/**
 * Parses user input date string into Excel serial number.
 * Interpreted as Day/Month/Year (D/M, D/M/YYYY, DD-MM-YYYY) or YYYY-MM-DD.
 * If year is omitted (e.g. "1/10"), uses current year; 2-digit year = 20xx;
 * 4-digit year >= 2400 = พ.ศ. (เช่น 15/10/2569 = 15 ต.ค. 2026)
 */
export function parseDateString(str: string): number | null {
  if (!str) return null;
  const trimmed = str.trim();

  // Match: D/M or DD/MM or D/M/YYYY or DD-MM-YYYY
  const m = /^(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?$/.exec(trimmed);
  if (m) {
    const day = parseInt(m[1], 10);
    const month = parseInt(m[2], 10);
    let year = m[3] ? parseInt(m[3], 10) : new Date().getFullYear();
    if (m[3] && m[3].length <= 2) {
      // Handle 2-digit year (e.g. 26 -> 2026)
      year += 2000;
    } else if (m[3] && m[3].length === 4) {
      year = toGregorianYear(year);
    } else if (m[3]) {
      return null; // ปี 3 หลัก
    }
    return validDateSerial(year, month, day);
  }

  // Match: YYYY-MM-DD or YYYY/MM/DD
  const isoMatch = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(trimmed);
  if (isoMatch) {
    return validDateSerial(
      toGregorianYear(parseInt(isoMatch[1], 10)),
      parseInt(isoMatch[2], 10),
      parseInt(isoMatch[3], 10)
    );
  }

  return null;
}

/** แสดงวันที่ dd/mm/yyyy ตามโหมดปีของแอป (พ.ศ. = ค.ศ. + 543) */
export function formatDate(serial: number, mode: YearMode = displayYearMode): string {
  const d = serialToDate(serial);
  const day = String(d.getUTCDate()).padStart(2, '0');
  const month = String(d.getUTCMonth() + 1).padStart(2, '0');
  const year = d.getUTCFullYear() + (mode === 'be' ? BE_OFFSET : 0);
  return `${day}/${month}/${year}`;
}

export const ERROR_DISPLAY_MAP: Record<CellError, string> = {
  '#DIV/0!': th.errors.div0,
  '#VALUE!': th.errors.value,
  '#REF!': th.errors.ref,
  '#NAME?': th.errors.name,
  '#CYCLE!': th.errors.cycle,
  '#ERR': th.errors.syntax,
};

export const ERROR_TO_THAI_MAP = ERROR_DISPLAY_MAP;

export const THAI_TO_ERROR_MAP: Record<string, CellError> = {
  '#หารศูนย์': '#DIV/0!',
  '#ค่าผิด': '#VALUE!',
  '#อ้างอิงหาย': '#REF!',
  '#ไม่รู้จักชื่อ': '#NAME?',
  '#วนซ้ำ': '#CYCLE!',
  '#สูตรผิด': '#ERR',
};

export const ERROR_TOOLTIP_MAP: Record<string, string> = {
  '#DIV/0!': th.errors.div0Tooltip,
  '#VALUE!': th.errors.valueTooltip,
  '#REF!': th.errors.refTooltip,
  '#NAME?': th.errors.nameTooltip,
  '#CYCLE!': th.errors.cycleTooltip,
  '#ERR': th.errors.syntaxTooltip,
  '#หารศูนย์': th.errors.div0Tooltip,
  '#ค่าผิด': th.errors.valueTooltip,
  '#อ้างอิงหาย': th.errors.refTooltip,
  '#ไม่รู้จักชื่อ': th.errors.nameTooltip,
  '#วนซ้ำ': th.errors.cycleTooltip,
  '#สูตรผิด': th.errors.syntaxTooltip,
};

export function formatNumberWithCommas(val: number, decimals: number): string {
  const dec = Math.max(0, Math.min(10, decimals));
  return val.toLocaleString('en-US', {
    minimumFractionDigits: dec,
    maximumFractionDigits: dec,
  });
}

/**
 * Formats cell value for UI display
 */
export function formatDisplayValue(
  value: CellValue,
  fmt: NumFmt = 'general',
  decimalOffset = 0
): string {
  if (value === '' || value === null || value === undefined) {
    return '';
  }

  // Error codes -> Thai display text
  if (typeof value === 'string' && value in ERROR_DISPLAY_MAP) {
    return ERROR_DISPLAY_MAP[value as CellError];
  }

  // Booleans
  if (typeof value === 'boolean') {
    return value ? 'TRUE' : 'FALSE';
  }

  // Plain strings
  if (typeof value !== 'number') {
    return String(value);
  }

  const num = value;

  switch (fmt) {
    case 'number2':
      return formatNumberWithCommas(num, 2 + decimalOffset);

    case 'integer':
      return formatNumberWithCommas(num, Math.max(0, 0 + decimalOffset));

    case 'thousands':
      return formatNumberWithCommas(num, 2 + decimalOffset);

    case 'baht':
      return `฿${formatNumberWithCommas(num, 2 + decimalOffset)}`;

    case 'percent': {
      const p = num * 100;
      const dec = Math.max(0, Math.min(10, 2 + decimalOffset));
      return `${formatNumberWithCommas(p, dec)}%`;
    }

    case 'date':
      return formatDate(num);

    case 'general':
    default: {
      if (decimalOffset !== 0) {
        return formatNumberWithCommas(num, Math.max(0, 2 + decimalOffset));
      }
      // General format: avoid float artifacts
      return String(Math.round(num * 1e10) / 1e10);
    }
  }
}

/**
 * ข้อความที่ผู้ใช้พิมพ์/วางซึ่งเป็นรูปแบบวันที่ → วันที่จริง (serial ค.ศ. + รูปแบบวันที่) ตาม SPEC 5.9 / 5.15
 * คืน null ถ้าไม่ใช่วันที่ที่ถูกต้อง (คงเป็นข้อความตามเดิม)
 */
export function convertTypedDate(text: string): { raw: string; fmt: 'date' } | null {
  const t = text.trim();
  if (!t || t.startsWith('=') || t.startsWith("'")) return null;
  const serial = parseDateString(t);
  return serial === null ? null : { raw: String(serial), fmt: 'date' };
}

/** ข้อความสำหรับแก้ไขเซล: เซลวันที่ (serial) แสดงเป็น dd/mm/yyyy ตามโหมดปี นอกนั้นใช้ค่าที่พิมพ์ไว้ */
export function editTextForCell(raw: string, fmt: NumFmt | undefined): string {
  if (fmt === 'date' && raw !== '' && !raw.startsWith('=') && !isNaN(Number(raw))) {
    return formatDate(Number(raw));
  }
  return raw;
}
