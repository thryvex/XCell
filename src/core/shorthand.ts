import { tokenize } from './tokenizer';
import { FormulaParser } from './evaluator';
import { parseA1, parseRange } from './address';
import { FUNCTIONS } from './functions';

const FUNCTION_ALIASES: Record<string, string> = {
  รวม: 'SUM',
  sum: 'SUM',
  เฉลี่ย: 'AVERAGE',
  avg: 'AVERAGE',
  mean: 'AVERAGE',
  average: 'AVERAGE',
  มัธยฐาน: 'MEDIAN',
  median: 'MEDIAN',
  ต่ำสุด: 'MIN',
  min: 'MIN',
  สูงสุด: 'MAX',
  max: 'MAX',
  นับ: 'COUNT',
  count: 'COUNT',
  รวมถ้า: 'SUMIF',
  sumif: 'SUMIF',
  นับถ้า: 'COUNTIF',
  countif: 'COUNTIF',
  เฉลี่ยถ้า: 'AVERAGEIF',
  averageif: 'AVERAGEIF',
  sumifs: 'SUMIFS',
  countifs: 'COUNTIFS',
};

/**
 * Checks if input can be converted from shorthand into a formula.
 * Returns formula string starting with '=' if convertible, or null if plain text.
 */
export function convertShorthand(input: string): string | null {
  if (!input) return null;
  const trimmed = input.trim();
  if (!trimmed) return null;

  // Already a formula or explicitly escaped text
  if (trimmed.startsWith('=') || trimmed.startsWith("'")) {
    return null;
  }

  // 1. Function call without '=' e.g. "sumif(a1:a10, \">10\")", "today()"
  const callMatch = /^([A-Za-zก-๙_]+)\s*\((.*)\)$/.exec(trimmed);
  if (callMatch) {
    const fnWord = callMatch[1].toLowerCase();
    const fnName =
      FUNCTION_ALIASES[fnWord] ||
      FUNCTION_ALIASES[callMatch[1]] ||
      (FUNCTIONS[callMatch[1].toUpperCase()] ? callMatch[1].toUpperCase() : null);
    if (fnName) {
      return `=${fnName}(${callMatch[2]})`;
    }
  }

  // 2. Function shorthand without parens: e.g. "sum a1:a10", "รวม a1:a10", "avg a1,b1,c1", "รวมถ้า a1:a10, b1"
  const fnMatch =
    /^(รวมถ้า|นับถ้า|เฉลี่ยถ้า|sumif|countif|averageif|sumifs|countifs|รวม|sum|เฉลี่ย|avg|mean|average|มัธยฐาน|median|ต่ำสุด|min|สูงสุด|max|นับ|count)\s+([A-Za-z0-9_$,:\s]+)$/i.exec(
      trimmed
    );
  if (fnMatch) {
    const fnWord = fnMatch[1].toLowerCase();
    const fnName = FUNCTION_ALIASES[fnWord] || FUNCTION_ALIASES[fnMatch[1]];
    if (fnName) {
      const argsPart = fnMatch[2].trim();
      // Split args by comma or whitespace
      const rawArgs = argsPart.split(/[\s,]+/).filter(Boolean);

      // Verify every argument is a valid cell ref, range, or number
      const validArgs = rawArgs.every((arg) => {
        return parseA1(arg) !== null || parseRange(arg) !== null || !isNaN(Number(arg));
      });

      if (validArgs && rawArgs.length > 0) {
        const formattedArgs = rawArgs.map((arg) => arg.toUpperCase()).join(',');
        return `=${fnName}(${formattedArgs})`;
      }
    }
  }

  // 2. Expression shorthand: e.g. "a1+c1", "b1*c1", "a1-b1", "a1/b1", "a1+5"
  // Must contain at least one operator (+ - * / ^ % &) and at least one cell reference
  const hasOperator = /[+\-*/^%&]/.test(trimmed);
  if (!hasOperator) return null;

  try {
    const tokens = tokenize('=' + trimmed, { includeWhitespace: false });
    let hasCellRef = false;
    let onlyValidTypes = true;

    for (const t of tokens) {
      if ((t.type === 'CELL_REF' || t.type === 'RANGE') && t.sheet !== undefined) {
        // ไม่แปลงข้อความที่มีการอ้างข้ามชีตให้เป็นสูตรอัตโนมัติ (ต้องพิมพ์ = เอง)
        onlyValidTypes = false;
        break;
      } else if (t.type === 'CELL_REF' || t.type === 'RANGE') {
        hasCellRef = true;
      } else if (
        t.type !== 'NUMBER' &&
        t.type !== 'OPERATOR' &&
        t.type !== 'LPAREN' &&
        t.type !== 'RPAREN'
      ) {
        onlyValidTypes = false;
        break;
      }
    }

    if (!hasCellRef || !onlyValidTypes) {
      return null;
    }

    // Verify it parses as a valid formula AST without errors
    const parser = new FormulaParser(tokens);
    parser.parse();

    // Reconstruct with uppercase references
    const formulaTokens = tokenize('=' + trimmed, { includeWhitespace: true });
    return (
      '=' +
      formulaTokens
        .map((t) => (t.type === 'CELL_REF' || t.type === 'RANGE' ? t.value.toUpperCase() : t.raw))
        .join('')
    );
  } catch {
    return null;
  }
}
