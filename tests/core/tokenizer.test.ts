import { describe, it, expect } from 'vitest';
import { tokenize } from '../../src/core/tokenizer';

describe('src/core/tokenizer', () => {
  it('tokenizes basic arithmetic and numbers', () => {
    const tokens = tokenize('=1 + 2.5 * 3');
    expect(tokens.map((t) => t.type)).toEqual(['NUMBER', 'OPERATOR', 'NUMBER', 'OPERATOR', 'NUMBER']);
    expect(tokens.map((t) => t.value)).toEqual(['1', '+', '2.5', '*', '3']);
  });

  it('tokenizes cell references and ranges', () => {
    const tokens = tokenize('=SUM(A1:B10, $C$1, D$2)');
    expect(tokens.map((t) => t.type)).toEqual([
      'FUNCTION',
      'LPAREN',
      'RANGE',
      'COMMA',
      'CELL_REF',
      'COMMA',
      'CELL_REF',
      'RPAREN',
    ]);
    expect(tokens.map((t) => t.value)).toEqual(['SUM', '(', 'A1:B10', ',', '$C$1', ',', 'D$2', ')']);
  });

  it('tokenizes comparison operators', () => {
    const tokens = tokenize('=A1 <> B1 AND C1 <= 10 AND D1 >= 5 AND E1 = 1');
    const ops = tokens.filter((t) => t.type === 'OPERATOR').map((t) => t.value);
    expect(ops).toContain('<>');
    expect(ops).toContain('<=');
    expect(ops).toContain('>=');
    expect(ops).toContain('=');
  });

  it('tokenizes percentage and exponentiation', () => {
    const tokens = tokenize('=10% + 2^3');
    expect(tokens.map((t) => t.value)).toEqual(['10', '%', '+', '2', '^', '3']);
  });

  it('tokenizes strings with escaped quotes', () => {
    const tokens = tokenize('="Hello ""World""" & " Test"');
    expect(tokens[0]).toEqual({
      type: 'STRING',
      value: 'Hello "World"',
      raw: '"Hello ""World"""',
      start: 1,
      end: 18,
    });
    expect(tokens[1]).toEqual({
      type: 'OPERATOR',
      value: '&',
      raw: '&',
      start: 19,
      end: 20,
    });
  });

  it('tokenizes boolean literals', () => {
    const tokens = tokenize('=IF(TRUE, 1, FALSE)');
    expect(tokens[2].type).toBe('BOOLEAN');
    expect(tokens[2].value).toBe('TRUE');
    expect(tokens[6].type).toBe('BOOLEAN');
    expect(tokens[6].value).toBe('FALSE');
  });

  it('preserves whitespace when requested', () => {
    const tokens = tokenize('=A1 + B1', { includeWhitespace: true });
    expect(tokens.map((t) => t.type)).toEqual(['CELL_REF', 'WHITESPACE', 'OPERATOR', 'WHITESPACE', 'CELL_REF']);
  });
});
