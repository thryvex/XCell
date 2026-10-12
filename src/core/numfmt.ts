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

/**
 * Parses user input date string into Excel serial number.
 * Interpreted as Day/Month/Year (DD/MM/YYYY or DD-MM-YYYY).
 * If year is omitted (e.g. "1/10"), uses current year.
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

    // Handle 2-digit year (e.g. 26 -> 2026)
    if (year < 100) {
      year += 2000;
    }

    // Validate range
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      const date = new Date(year, month - 1, day);
      if (date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day) {
        return dateToSerial(date);
      }
    }
  }

  // Match: YYYY-MM-DD or YYYY/MM/DD
  const isoMatch = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(trimmed);
  if (isoMatch) {
    const y = parseInt(isoMatch[1], 10);
    const m = parseInt(isoMatch[2], 10);
    const d = parseInt(isoMatch[3], 10);
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31) {
      const date = new Date(y, m - 1, d);
      if (date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d) {
        return dateToSerial(date);
      }
    }
  }

  return null;
}

export function formatDate(serial: number): string {
  const d = serialToDate(serial);
  const day = String(d.getUTCDate()).padStart(2, '0');
  const month = String(d.getUTCMonth() + 1).padStart(2, '0');
  const year = d.getUTCFullYear();
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
