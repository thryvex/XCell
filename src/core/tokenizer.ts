export type TokenType =
  | 'NUMBER'
  | 'STRING'
  | 'BOOLEAN'
  | 'ERROR'
  | 'CELL_REF'
  | 'RANGE'
  | 'FUNCTION'
  | 'OPERATOR'
  | 'LPAREN'
  | 'RPAREN'
  | 'COMMA'
  | 'WHITESPACE'
  | 'UNKNOWN';

export interface Token {
  type: TokenType;
  value: string; // สำหรับ CELL_REF/RANGE คือเฉพาะส่วนที่อยู่ (เช่น "A1:B2") ตัวพิมพ์ใหญ่
  raw: string;
  start: number;
  end: number;
  sheet?: string; // ชื่อชีต (ถอดเครื่องหมายคำพูดแล้ว) ถ้า reference ระบุชีต เช่น 'ชีต 2'!A1
}

const CELL_REF_REGEX = /^\$?[A-Za-z]+\$?\d+$/;
const RANGE_REGEX = /^\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d+$/;

export function tokenize(formulaText: string, options: { includeWhitespace?: boolean } = {}): Token[] {
  const tokens: Token[] = [];
  if (!formulaText) return tokens;

  let text = formulaText;
  let offset = 0;
  if (text.startsWith('=')) {
    text = text.slice(1);
    offset = 1;
  }

  let i = 0;
  const len = text.length;

  while (i < len) {
    const ch = text[i];

    // 1. Whitespace
    if (/\s/.test(ch)) {
      const start = i;
      while (i < len && /\s/.test(text[i])) {
        i++;
      }
      if (options.includeWhitespace) {
        tokens.push({
          type: 'WHITESPACE',
          value: text.slice(start, i),
          raw: text.slice(start, i),
          start: start + offset,
          end: i + offset,
        });
      }
      continue;
    }

    // 2. String literal "..."
    if (ch === '"') {
      const start = i;
      i++; // skip opening quote
      let strVal = '';
      while (i < len) {
        if (text[i] === '"') {
          if (i + 1 < len && text[i + 1] === '"') {
            strVal += '"';
            i += 2;
          } else {
            i++; // skip closing quote
            break;
          }
        } else {
          strVal += text[i];
          i++;
        }
      }
      tokens.push({
        type: 'STRING',
        value: strVal,
        raw: text.slice(start, i),
        start: start + offset,
        end: i + offset,
      });
      continue;
    }

    // 3. Numbers
    if (/\d/.test(ch) || (ch === '.' && i + 1 < len && /\d/.test(text[i + 1]))) {
      const start = i;
      let hasDot = ch === '.';
      i++;
      while (i < len) {
        if (text[i] === '.') {
          if (hasDot) break;
          hasDot = true;
          i++;
        } else if (/\d/.test(text[i])) {
          i++;
        } else {
          break;
        }
      }
      tokens.push({
        type: 'NUMBER',
        value: text.slice(start, i),
        raw: text.slice(start, i),
        start: start + offset,
        end: i + offset,
      });
      continue;
    }

    // 4. Parentheses and Comma
    if (ch === '(') {
      tokens.push({ type: 'LPAREN', value: '(', raw: '(', start: i + offset, end: i + 1 + offset });
      i++;
      continue;
    }
    if (ch === ')') {
      tokens.push({ type: 'RPAREN', value: ')', raw: ')', start: i + offset, end: i + 1 + offset });
      i++;
      continue;
    }
    if (ch === ',') {
      tokens.push({ type: 'COMMA', value: ',', raw: ',', start: i + offset, end: i + 1 + offset });
      i++;
      continue;
    }

    // 5. Comparison operators
    if (ch === '<') {
      if (i + 1 < len && (text[i + 1] === '=' || text[i + 1] === '>')) {
        const val = text.slice(i, i + 2);
        tokens.push({ type: 'OPERATOR', value: val, raw: val, start: i + offset, end: i + 2 + offset });
        i += 2;
      } else {
        tokens.push({ type: 'OPERATOR', value: '<', raw: '<', start: i + offset, end: i + 1 + offset });
        i++;
      }
      continue;
    }
    if (ch === '>') {
      if (i + 1 < len && text[i + 1] === '=') {
        const val = text.slice(i, i + 2);
        tokens.push({ type: 'OPERATOR', value: val, raw: val, start: i + offset, end: i + 2 + offset });
        i += 2;
      } else {
        tokens.push({ type: 'OPERATOR', value: '>', raw: '>', start: i + offset, end: i + 1 + offset });
        i++;
      }
      continue;
    }
    if (ch === '=') {
      tokens.push({ type: 'OPERATOR', value: '=', raw: '=', start: i + offset, end: i + 1 + offset });
      i++;
      continue;
    }

    // 6. Arithmetic operators
    if ('+-*/^%&'.includes(ch)) {
      tokens.push({ type: 'OPERATOR', value: ch, raw: ch, start: i + offset, end: i + 1 + offset });
      i++;
      continue;
    }

    // 7. Words
    if (ch === '$' || /[A-Za-z_\u0E00-\u0E7F]/.test(ch)) {
      const start = i;
      while (i < len && (/[A-Za-z0-9_$\u0E00-\u0E7F]/.test(text[i]) || text[i] === ':')) {
        i++;
      }
      const rawWord = text.slice(start, i);

      if (RANGE_REGEX.test(rawWord)) {
        tokens.push({
          type: 'RANGE',
          value: rawWord.toUpperCase(),
          raw: rawWord,
          start: start + offset,
          end: i + offset,
        });
        continue;
      }

      let nextNonWs = i;
      while (nextNonWs < len && /\s/.test(text[nextNonWs])) nextNonWs++;
      if (nextNonWs < len && text[nextNonWs] === '(' && !rawWord.includes(':')) {
        tokens.push({
          type: 'FUNCTION',
          value: rawWord.toUpperCase(),
          raw: rawWord,
          start: start + offset,
          end: i + offset,
        });
        continue;
      }

      const upper = rawWord.toUpperCase();
      if (upper === 'TRUE' || upper === 'FALSE') {
        tokens.push({
          type: 'BOOLEAN',
          value: upper,
          raw: rawWord,
          start: start + offset,
          end: i + offset,
        });
        continue;
      }

      if (CELL_REF_REGEX.test(rawWord)) {
        tokens.push({
          type: 'CELL_REF',
          value: upper,
          raw: rawWord,
          start: start + offset,
          end: i + offset,
        });
        continue;
      }

      tokens.push({
        type: 'FUNCTION',
        value: upper,
        raw: rawWord,
        start: start + offset,
        end: i + offset,
      });
      continue;
    }

    // 8. Unknown fallback
    tokens.push({
      type: 'UNKNOWN',
      value: ch,
      raw: ch,
      start: i + offset,
      end: i + 1 + offset,
    });
    i++;
  }

  return tokens;
}
