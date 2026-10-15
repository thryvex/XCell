import { CellValue, CellError } from './model';
import {
  dateToSerial,
  serialToDate,
  parseDateString,
  formatNumberWithCommas,
} from './numfmt';

export type FunctionImpl = (...args: any[]) => CellValue;

export function normalizeFloat(n: number): number {
  if (!isFinite(n) || isNaN(n)) return n;
  return parseFloat(n.toPrecision(15));
}

// Helper to flatten nested arrays (from ranges or multiple arguments)
export function flattenArgs(args: any[]): any[] {
  const result: any[] = [];
  for (const item of args) {
    if (Array.isArray(item)) {
      result.push(...flattenArgs(item));
    } else {
      result.push(item);
    }
  }
  return result;
}

// Extract numbers from range arguments (skips strings, empty cells, booleans)
export function extractNumbers(args: any[]): number[] {
  const flat = flattenArgs(args);
  const numbers: number[] = [];
  for (const v of flat) {
    if (typeof v === 'number' && !isNaN(v) && isFinite(v)) {
      numbers.push(v);
    }
  }
  return numbers;
}

export function isCellError(val: any): val is CellError {
  return (
    typeof val === 'string' &&
    (val === '#DIV/0!' ||
      val === '#VALUE!' ||
      val === '#REF!' ||
      val === '#NAME?' ||
      val === '#CYCLE!' ||
      val === '#ERR')
  );
}

function escapeRegex(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export function wildcardToRegExp(pattern: string): RegExp {
  let res = '^';
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === '~' && i + 1 < pattern.length) {
      const next = pattern[i + 1];
      if (next === '*' || next === '?' || next === '~') {
        res += escapeRegex(next);
        i++;
        continue;
      }
    }
    if (c === '*') {
      res += '.*';
    } else if (c === '?') {
      res += '.';
    } else {
      res += escapeRegex(c);
    }
  }
  res += '$';
  return new RegExp(res, 'i');
}

export function createCriteriaMatcher(criteria: any): (val: any) => boolean {
  if (isCellError(criteria)) {
    return () => true;
  }

  if (criteria === null || criteria === undefined) {
    return (val: any) => val === '' || val === null || val === undefined;
  }

  if (typeof criteria === 'number' || typeof criteria === 'boolean') {
    return (val: any) => val === criteria;
  }

  const str = String(criteria);

  let op = '=';
  let rawTarget = str;

  if (str.startsWith('<=') || str.startsWith('>=') || str.startsWith('<>')) {
    op = str.slice(0, 2);
    rawTarget = str.slice(2);
  } else if (str.startsWith('<') || str.startsWith('>') || str.startsWith('=')) {
    op = str.slice(0, 1);
    rawTarget = str.slice(1);
  }

  // Handle empty target
  if (rawTarget === '') {
    if (op === '=') {
      return (val: any) => val === '' || val === null || val === undefined;
    }
    if (op === '<>') {
      return (val: any) => val !== '' && val !== null && val !== undefined;
    }
  }

  // Check if rawTarget is numeric
  const trimmed = rawTarget.trim();
  const numTarget = trimmed !== '' && !isNaN(Number(trimmed)) ? Number(trimmed) : null;

  if (numTarget !== null) {
    return (val: any) => {
      let cellNum: number | null = null;
      if (typeof val === 'number') {
        cellNum = val;
      } else if (typeof val === 'string' && val.trim() !== '') {
        const cleaned = val.replace(/,/g, '');
        if (!isNaN(Number(cleaned))) {
          cellNum = Number(cleaned);
        } else {
          const parsed = parseDateString(val);
          if (parsed !== null) cellNum = parsed;
        }
      }

      if (cellNum === null) {
        if (op === '<>') return true;
        return false;
      }

      switch (op) {
        case '=':
          return cellNum === numTarget;
        case '<>':
          return cellNum !== numTarget;
        case '<':
          return cellNum < numTarget;
        case '<=':
          return cellNum <= numTarget;
        case '>':
          return cellNum > numTarget;
        case '>=':
          return cellNum >= numTarget;
        default:
          return false;
      }
    };
  }

  // Text / wildcard matching
  const regex = wildcardToRegExp(rawTarget);

  return (val: any) => {
    if (val === null || val === undefined) {
      val = '';
    }
    const valStr = String(val);

    if (op === '=') {
      return regex.test(valStr);
    }
    if (op === '<>') {
      return !regex.test(valStr);
    }

    const cmp = valStr.localeCompare(rawTarget, undefined, { sensitivity: 'base' });
    switch (op) {
      case '<':
        return cmp < 0;
      case '<=':
        return cmp <= 0;
      case '>':
        return cmp > 0;
      case '>=':
        return cmp >= 0;
      default:
        return false;
    }
  };
}

export function formatWithTextFunction(value: any, fmtStr: any): CellValue {
  if (isCellError(value)) return value;
  if (isCellError(fmtStr)) return fmtStr;

  if (typeof fmtStr !== 'string') return '#VALUE!';
  const f = fmtStr.trim();

  const ALLOWED_FORMATS = new Set([
    '0',
    '0.00',
    '#,##0',
    '#,##0.00',
    '0%',
    '0.00%',
    'dd/mm/yyyy',
    'mm/yyyy',
    'yyyy-mm-dd',
    'd/m/yyyy',
  ]);

  if (!ALLOWED_FORMATS.has(f)) {
    return '#VALUE!';
  }

  if (['dd/mm/yyyy', 'mm/yyyy', 'yyyy-mm-dd', 'd/m/yyyy'].includes(f)) {
    let serial: number | null = null;
    if (typeof value === 'number') {
      serial = value;
    } else if (typeof value === 'string') {
      serial = parseDateString(value);
      if (serial === null && !isNaN(Number(value))) {
        serial = Number(value);
      }
    }
    if (serial === null || isNaN(serial) || !isFinite(serial)) {
      return '#VALUE!';
    }

    const d = serialToDate(serial);
    const dayPadded = String(d.getUTCDate()).padStart(2, '0');
    const dayRaw = String(d.getUTCDate());
    const monthPadded = String(d.getUTCMonth() + 1).padStart(2, '0');
    const monthRaw = String(d.getUTCMonth() + 1);
    const year = d.getUTCFullYear();

    switch (f) {
      case 'dd/mm/yyyy':
        return `${dayPadded}/${monthPadded}/${year}`;
      case 'd/m/yyyy':
        return `${dayRaw}/${monthRaw}/${year}`;
      case 'mm/yyyy':
        return `${monthPadded}/${year}`;
      case 'yyyy-mm-dd':
        return `${year}-${monthPadded}-${dayPadded}`;
      default:
        return '#VALUE!';
    }
  }

  const num = typeof value === 'number' ? value : Number(value);
  if (isNaN(num) || !isFinite(num)) {
    return '#VALUE!';
  }

  switch (f) {
    case '0':
      return Math.round(num).toString();
    case '0.00':
      return num.toFixed(2);
    case '#,##0':
      return formatNumberWithCommas(Math.round(num), 0);
    case '#,##0.00':
      return formatNumberWithCommas(num, 2);
    case '0%':
      return Math.round(num * 100).toString() + '%';
    case '0.00%':
      return (num * 100).toFixed(2) + '%';
    default:
      return '#VALUE!';
  }
}

export const FUNCTIONS: Record<string, FunctionImpl> = {
  // Math & Statistics
  SUM: (...args: any[]): CellValue => {
    for (const v of flattenArgs(args)) {
      if (isCellError(v)) return v;
    }
    const nums = extractNumbers(args);
    return nums.reduce((acc, curr) => acc + curr, 0);
  },

  AVERAGE: (...args: any[]): CellValue => {
    for (const v of flattenArgs(args)) {
      if (isCellError(v)) return v;
    }
    const nums = extractNumbers(args);
    if (nums.length === 0) return '#DIV/0!';
    const sum = nums.reduce((acc, curr) => acc + curr, 0);
    return sum / nums.length;
  },

  MEDIAN: (...args: any[]): CellValue => {
    for (const v of flattenArgs(args)) {
      if (isCellError(v)) return v;
    }
    const nums = extractNumbers(args);
    if (nums.length === 0) return '#DIV/0!';
    nums.sort((a, b) => a - b);
    const mid = Math.floor(nums.length / 2);
    if (nums.length % 2 === 1) {
      return nums[mid];
    }
    return (nums[mid - 1] + nums[mid]) / 2;
  },

  MIN: (...args: any[]): CellValue => {
    for (const v of flattenArgs(args)) {
      if (isCellError(v)) return v;
    }
    const nums = extractNumbers(args);
    if (nums.length === 0) return 0;
    return Math.min(...nums);
  },

  MAX: (...args: any[]): CellValue => {
    for (const v of flattenArgs(args)) {
      if (isCellError(v)) return v;
    }
    const nums = extractNumbers(args);
    if (nums.length === 0) return 0;
    return Math.max(...nums);
  },

  COUNT: (...args: any[]): CellValue => {
    return extractNumbers(args).length;
  },

  COUNTA: (...args: any[]): CellValue => {
    const flat = flattenArgs(args);
    return flat.filter((v) => v !== '' && v !== null && v !== undefined).length;
  },

  PRODUCT: (...args: any[]): CellValue => {
    for (const v of flattenArgs(args)) {
      if (isCellError(v)) return v;
    }
    const nums = extractNumbers(args);
    if (nums.length === 0) return 0;
    return nums.reduce((acc, curr) => acc * curr, 1);
  },

  ROUND: (val: any, digits: any = 0): CellValue => {
    if (isCellError(val)) return val;
    const num = Number(val);
    const d = Number(digits);
    if (isNaN(num) || isNaN(d)) return '#VALUE!';
    const factor = Math.pow(10, d);
    return Math.round(num * factor) / factor;
  },

  ROUNDUP: (val: any, digits: any = 0): CellValue => {
    if (isCellError(val)) return val;
    const num = Number(val);
    const d = Number(digits);
    if (isNaN(num) || isNaN(d)) return '#VALUE!';
    const factor = Math.pow(10, d);
    const sign = num >= 0 ? 1 : -1;
    return (sign * Math.ceil(Math.abs(num) * factor)) / factor;
  },

  ROUNDDOWN: (val: any, digits: any = 0): CellValue => {
    if (isCellError(val)) return val;
    const num = Number(val);
    const d = Number(digits);
    if (isNaN(num) || isNaN(d)) return '#VALUE!';
    const factor = Math.pow(10, d);
    const sign = num >= 0 ? 1 : -1;
    return (sign * Math.floor(Math.abs(num) * factor)) / factor;
  },

  ABS: (val: any): CellValue => {
    if (isCellError(val)) return val;
    const num = Number(val);
    if (isNaN(num)) return '#VALUE!';
    return Math.abs(num);
  },

  INT: (val: any): CellValue => {
    if (isCellError(val)) return val;
    const num = Number(val);
    if (isNaN(num)) return '#VALUE!';
    return Math.floor(num);
  },

  SQRT: (val: any): CellValue => {
    if (isCellError(val)) return val;
    const num = Number(val);
    if (isNaN(num) || num < 0) return '#VALUE!';
    return Math.sqrt(num);
  },

  POWER: (base: any, exp: any): CellValue => {
    if (isCellError(base)) return base;
    if (isCellError(exp)) return exp;
    const b = Number(base);
    const e = Number(exp);
    if (isNaN(b) || isNaN(e)) return '#VALUE!';
    return Math.pow(b, e);
  },

  MOD: (num: any, divisor: any): CellValue => {
    if (isCellError(num)) return num;
    if (isCellError(divisor)) return divisor;
    const n = Number(num);
    const d = Number(divisor);
    if (isNaN(n) || isNaN(d)) return '#VALUE!';
    if (d === 0) return '#DIV/0!';
    return n - d * Math.floor(n / d);
  },

  // Logical
  IF: (condition: any, valIfTrue: any, valIfFalse: any = false): CellValue => {
    if (isCellError(condition)) return condition;
    const isTrue =
      condition === true ||
      condition === 'TRUE' ||
      (typeof condition === 'number' && condition !== 0) ||
      (typeof condition === 'string' && condition.length > 0 && condition !== 'FALSE');
    return isTrue ? valIfTrue : valIfFalse;
  },

  IFERROR: (val: any, valIfError: any = ''): CellValue => {
    return isCellError(val) ? valIfError : val;
  },

  AND: (...args: any[]): CellValue => {
    const flat = flattenArgs(args);
    if (flat.length === 0) return '#VALUE!';
    for (const v of flat) {
      if (isCellError(v)) return v;
      if (v === false || v === 0 || v === 'FALSE' || v === '') return false;
    }
    return true;
  },

  OR: (...args: any[]): CellValue => {
    const flat = flattenArgs(args);
    if (flat.length === 0) return '#VALUE!';
    for (const v of flat) {
      if (isCellError(v)) return v;
      if (v === true || (typeof v === 'number' && v !== 0) || v === 'TRUE') return true;
    }
    return false;
  },

  NOT: (val: any): CellValue => {
    if (isCellError(val)) return val;
    if (val === true || (typeof val === 'number' && val !== 0) || val === 'TRUE') {
      return false;
    }
    return true;
  },

  // Phase 4A: Conditional Aggregations
  COUNTIF: (range: any, criteria: any): CellValue => {
    if (isCellError(criteria)) return criteria;
    const flat = flattenArgs([range]);
    const matcher = createCriteriaMatcher(criteria);
    let count = 0;
    for (const v of flat) {
      if (isCellError(v)) return v;
      if (matcher(v)) {
        count++;
      }
    }
    return count;
  },

  SUMIF: (range: any, criteria: any, sumRange?: any): CellValue => {
    if (isCellError(criteria)) return criteria;
    const flatRange = flattenArgs([range]);
    const flatSum = sumRange !== undefined ? flattenArgs([sumRange]) : flatRange;
    const matcher = createCriteriaMatcher(criteria);
    let sum = 0;

    for (let i = 0; i < flatRange.length; i++) {
      const condVal = flatRange[i];
      if (isCellError(condVal)) return condVal;
      if (matcher(condVal)) {
        const sumVal = i < flatSum.length ? flatSum[i] : 0;
        if (isCellError(sumVal)) return sumVal;
        if (typeof sumVal === 'number' && isFinite(sumVal)) {
          sum += sumVal;
        } else if (typeof sumVal === 'string' && sumVal.trim() !== '') {
          const num = Number(sumVal.replace(/,/g, ''));
          if (!isNaN(num) && isFinite(num)) {
            sum += num;
          }
        }
      }
    }
    return normalizeFloat(sum);
  },

  AVERAGEIF: (range: any, criteria: any, avgRange?: any): CellValue => {
    if (isCellError(criteria)) return criteria;
    const flatRange = flattenArgs([range]);
    const flatAvg = avgRange !== undefined ? flattenArgs([avgRange]) : flatRange;
    const matcher = createCriteriaMatcher(criteria);
    let sum = 0;
    let count = 0;

    for (let i = 0; i < flatRange.length; i++) {
      const condVal = flatRange[i];
      if (isCellError(condVal)) return condVal;
      if (matcher(condVal)) {
        const avgVal = i < flatAvg.length ? flatAvg[i] : '';
        if (isCellError(avgVal)) return avgVal;
        if (typeof avgVal === 'number' && isFinite(avgVal)) {
          sum += avgVal;
          count++;
        } else if (typeof avgVal === 'string' && avgVal.trim() !== '') {
          const num = Number(avgVal.replace(/,/g, ''));
          if (!isNaN(num) && isFinite(num)) {
            sum += num;
            count++;
          }
        }
      }
    }
    if (count === 0) return '#DIV/0!';
    return normalizeFloat(sum / count);
  },

  SUMIFS: (sumRange: any, ...criteriaPairs: any[]): CellValue => {
    if (criteriaPairs.length < 2 || criteriaPairs.length % 2 !== 0) {
      return '#VALUE!';
    }
    const flatSum = flattenArgs([sumRange]);
    const numPairs = criteriaPairs.length / 2;
    const critRanges: any[][] = [];
    const matchers: ((val: any) => boolean)[] = [];

    for (let p = 0; p < numPairs; p++) {
      const cRange = criteriaPairs[p * 2];
      const crit = criteriaPairs[p * 2 + 1];
      if (isCellError(crit)) return crit;
      const flatCRange = flattenArgs([cRange]);
      if (flatCRange.length !== flatSum.length) {
        return '#VALUE!';
      }
      critRanges.push(flatCRange);
      matchers.push(createCriteriaMatcher(crit));
    }

    let sum = 0;
    for (let i = 0; i < flatSum.length; i++) {
      let allMatch = true;
      for (let p = 0; p < numPairs; p++) {
        const condVal = critRanges[p][i];
        if (isCellError(condVal)) return condVal;
        if (!matchers[p](condVal)) {
          allMatch = false;
          break;
        }
      }

      if (allMatch) {
        const sumVal = flatSum[i];
        if (isCellError(sumVal)) return sumVal;
        if (typeof sumVal === 'number' && isFinite(sumVal)) {
          sum += sumVal;
        } else if (typeof sumVal === 'string' && sumVal.trim() !== '') {
          const num = Number(sumVal.replace(/,/g, ''));
          if (!isNaN(num) && isFinite(num)) {
            sum += num;
          }
        }
      }
    }

    return normalizeFloat(sum);
  },

  COUNTIFS: (...criteriaPairs: any[]): CellValue => {
    if (criteriaPairs.length < 2 || criteriaPairs.length % 2 !== 0) {
      return '#VALUE!';
    }
    const numPairs = criteriaPairs.length / 2;
    const critRanges: any[][] = [];
    const matchers: ((val: any) => boolean)[] = [];

    const firstRange = flattenArgs([criteriaPairs[0]]);
    const expectedLen = firstRange.length;
    critRanges.push(firstRange);
    if (isCellError(criteriaPairs[1])) return criteriaPairs[1];
    matchers.push(createCriteriaMatcher(criteriaPairs[1]));

    for (let p = 1; p < numPairs; p++) {
      const cRange = criteriaPairs[p * 2];
      const crit = criteriaPairs[p * 2 + 1];
      if (isCellError(crit)) return crit;
      const flatCRange = flattenArgs([cRange]);
      if (flatCRange.length !== expectedLen) {
        return '#VALUE!';
      }
      critRanges.push(flatCRange);
      matchers.push(createCriteriaMatcher(crit));
    }

    let count = 0;
    for (let i = 0; i < expectedLen; i++) {
      let allMatch = true;
      for (let p = 0; p < numPairs; p++) {
        const condVal = critRanges[p][i];
        if (isCellError(condVal)) return condVal;
        if (!matchers[p](condVal)) {
          allMatch = false;
          break;
        }
      }
      if (allMatch) count++;
    }

    return count;
  },

  // Phase 4A: Date & Time functions
  DATE: (y: any, m: any, d: any): CellValue => {
    if (isCellError(y)) return y;
    if (isCellError(m)) return m;
    if (isCellError(d)) return d;

    const yearNum = Number(y);
    const monthNum = Number(m);
    const dayNum = Number(d);

    if (isNaN(yearNum) || isNaN(monthNum) || isNaN(dayNum)) return '#VALUE!';

    let year = Math.floor(yearNum);
    if (year < 0 || year >= 10000) return '#VALUE!';
    if (year < 100) year += 1900;

    const month = Math.floor(monthNum);
    const day = Math.floor(dayNum);

    const dateObj = new Date(year, month - 1, day);
    return dateToSerial(dateObj);
  },

  YEAR: (dateVal: any): CellValue => {
    if (isCellError(dateVal)) return dateVal;
    let serial: number | null = null;
    if (typeof dateVal === 'number') {
      serial = dateVal;
    } else if (typeof dateVal === 'string') {
      serial = parseDateString(dateVal);
      if (serial === null && !isNaN(Number(dateVal))) {
        serial = Number(dateVal);
      }
    }
    if (serial === null || isNaN(serial) || !isFinite(serial)) return '#VALUE!';
    return serialToDate(serial).getUTCFullYear();
  },

  MONTH: (dateVal: any): CellValue => {
    if (isCellError(dateVal)) return dateVal;
    let serial: number | null = null;
    if (typeof dateVal === 'number') {
      serial = dateVal;
    } else if (typeof dateVal === 'string') {
      serial = parseDateString(dateVal);
      if (serial === null && !isNaN(Number(dateVal))) {
        serial = Number(dateVal);
      }
    }
    if (serial === null || isNaN(serial) || !isFinite(serial)) return '#VALUE!';
    return serialToDate(serial).getUTCMonth() + 1;
  },

  DAY: (dateVal: any): CellValue => {
    if (isCellError(dateVal)) return dateVal;
    let serial: number | null = null;
    if (typeof dateVal === 'number') {
      serial = dateVal;
    } else if (typeof dateVal === 'string') {
      serial = parseDateString(dateVal);
      if (serial === null && !isNaN(Number(dateVal))) {
        serial = Number(dateVal);
      }
    }
    if (serial === null || isNaN(serial) || !isFinite(serial)) return '#VALUE!';
    return serialToDate(serial).getUTCDate();
  },

  TODAY: (): CellValue => {
    const now = new Date();
    return dateToSerial(now);
  },

  NOW: (): CellValue => {
    const now = new Date();
    const baseSerial = dateToSerial(now);
    const frac =
      (now.getHours() * 3600 +
        now.getMinutes() * 60 +
        now.getSeconds() +
        now.getMilliseconds() / 1000) /
      86400;
    return normalizeFloat(baseSerial + frac);
  },

  // Phase 4A: Text functions
  TEXT: (val: any, formatText: any): CellValue => {
    return formatWithTextFunction(val, formatText);
  },

  CONCAT: (...args: any[]): CellValue => {
    const flat = flattenArgs(args);
    for (const v of flat) {
      if (isCellError(v)) return v;
    }
    return flat.map((v) => (v === null || v === undefined ? '' : String(v))).join('');
  },

  LEN: (val: any): CellValue => {
    if (isCellError(val)) return val;
    if (val === null || val === undefined || val === '') return 0;
    return String(val).length;
  },
};

// Aliases
FUNCTIONS.MEAN = FUNCTIONS.AVERAGE;
FUNCTIONS.CONCATENATE = FUNCTIONS.CONCAT;
