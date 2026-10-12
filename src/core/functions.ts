import { CellValue, CellError } from './model';

export type FunctionImpl = (...args: any[]) => CellValue;

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

export const FUNCTIONS: Record<string, FunctionImpl> = {
  // Math & Statistics
  SUM: (...args: any[]): CellValue => {
    // Check for propagation of errors
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
    // Excel MOD returns result with same sign as divisor
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
};

// Alias MEAN -> AVERAGE
FUNCTIONS.MEAN = FUNCTIONS.AVERAGE;
