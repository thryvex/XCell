import { CellValue, CellError, Sheet } from './model';
import { tokenize, Token } from './tokenizer';
import { parseA1, parseRange, cellToA1 } from './address';
import { FUNCTIONS, isCellError } from './functions';

// AST Nodes
export type ASTNode =
  | { type: 'Number'; value: number }
  | { type: 'String'; value: string }
  | { type: 'Boolean'; value: boolean }
  | { type: 'CellRef'; a1: string }
  | { type: 'Range'; rangeStr: string }
  | { type: 'UnaryOp'; op: '+' | '-'; expr: ASTNode }
  | { type: 'PostfixOp'; op: '%'; expr: ASTNode }
  | { type: 'BinaryOp'; op: string; left: ASTNode; right: ASTNode }
  | { type: 'FunctionCall'; name: string; args: ASTNode[] };

export function normalizeFloat(n: number): number {
  if (!isFinite(n) || isNaN(n)) return n;
  return parseFloat(n.toPrecision(15));
}

// Parser Class
export class FormulaParser {
  private tokens: Token[];
  private pos = 0;

  constructor(tokens: Token[]) {
    // Filter out whitespace for parsing
    this.tokens = tokens.filter((t) => t.type !== 'WHITESPACE');
  }

  private peek(): Token | undefined {
    return this.tokens[this.pos];
  }

  private consume(): Token {
    const t = this.tokens[this.pos];
    this.pos++;
    return t;
  }

  private match(type: string, value?: string): boolean {
    const t = this.peek();
    if (!t) return false;
    if (t.type !== type) return false;
    if (value !== undefined && t.value.toUpperCase() !== value.toUpperCase()) return false;
    return true;
  }

  public parse(): ASTNode {
    const node = this.parseExpression();
    if (this.pos < this.tokens.length) {
      throw new Error('#ERR');
    }
    return node;
  }

  // Precedence levels:
  // 1. Comparison: =, <>, <, >, <=, >=
  // 2. Concat: &
  // 3. Add/Sub: +, -
  // 4. Mul/Div: *, /
  // 5. Power: ^
  // 6. Unary: +, -
  // 7. Postfix: %
  // 8. Primary

  private parseExpression(): ASTNode {
    return this.parseComparison();
  }

  private parseComparison(): ASTNode {
    let left = this.parseConcat();
    while (
      this.match('OPERATOR', '=') ||
      this.match('OPERATOR', '<>') ||
      this.match('OPERATOR', '<') ||
      this.match('OPERATOR', '>') ||
      this.match('OPERATOR', '<=') ||
      this.match('OPERATOR', '>=')
    ) {
      const op = this.consume().value;
      const right = this.parseConcat();
      left = { type: 'BinaryOp', op, left, right };
    }
    return left;
  }

  private parseConcat(): ASTNode {
    let left = this.parseAddSub();
    while (this.match('OPERATOR', '&')) {
      const op = this.consume().value;
      const right = this.parseAddSub();
      left = { type: 'BinaryOp', op, left, right };
    }
    return left;
  }

  private parseAddSub(): ASTNode {
    let left = this.parseMulDiv();
    while (this.match('OPERATOR', '+') || this.match('OPERATOR', '-')) {
      const op = this.consume().value;
      const right = this.parseMulDiv();
      left = { type: 'BinaryOp', op, left, right };
    }
    return left;
  }

  private parseMulDiv(): ASTNode {
    let left = this.parsePower();
    while (this.match('OPERATOR', '*') || this.match('OPERATOR', '/')) {
      const op = this.consume().value;
      const right = this.parsePower();
      left = { type: 'BinaryOp', op, left, right };
    }
    return left;
  }

  private parsePower(): ASTNode {
    let left = this.parseUnary();
    if (this.match('OPERATOR', '^')) {
      const op = this.consume().value;
      const right = this.parsePower(); // right-associative
      left = { type: 'BinaryOp', op, left, right };
    }
    return left;
  }

  private parseUnary(): ASTNode {
    if (this.match('OPERATOR', '+') || this.match('OPERATOR', '-')) {
      const op = this.consume().value as '+' | '-';
      const expr = this.parseUnary();
      return { type: 'UnaryOp', op, expr };
    }
    return this.parsePostfix();
  }

  private parsePostfix(): ASTNode {
    let expr = this.parsePrimary();
    while (this.match('OPERATOR', '%')) {
      this.consume();
      expr = { type: 'PostfixOp', op: '%', expr };
    }
    return expr;
  }

  private parsePrimary(): ASTNode {
    const t = this.peek();
    if (!t) throw new Error('#ERR');

    if (t.type === 'NUMBER') {
      this.consume();
      return { type: 'Number', value: parseFloat(t.value) };
    }

    if (t.type === 'STRING') {
      this.consume();
      return { type: 'String', value: t.value };
    }

    if (t.type === 'BOOLEAN') {
      this.consume();
      return { type: 'Boolean', value: t.value.toUpperCase() === 'TRUE' };
    }

    if (t.type === 'CELL_REF') {
      this.consume();
      return { type: 'CellRef', a1: t.value.toUpperCase() };
    }

    if (t.type === 'RANGE') {
      this.consume();
      return { type: 'Range', rangeStr: t.value.toUpperCase() };
    }

    if (t.type === 'FUNCTION') {
      const name = this.consume().value.toUpperCase();
      if (!this.match('LPAREN')) {
        // Unknown standalone word
        throw new Error('#NAME?');
      }
      this.consume(); // '('
      const args: ASTNode[] = [];
      if (!this.match('RPAREN')) {
        while (true) {
          args.push(this.parseExpression());
          if (this.match('COMMA')) {
            this.consume();
          } else {
            break;
          }
        }
      }
      if (!this.match('RPAREN')) {
        throw new Error('#ERR');
      }
      this.consume(); // ')'
      return { type: 'FunctionCall', name, args };
    }

    if (t.type === 'LPAREN') {
      this.consume();
      const expr = this.parseExpression();
      if (!this.match('RPAREN')) {
        throw new Error('#ERR');
      }
      this.consume();
      return expr;
    }

    throw new Error('#ERR');
  }
}

// Extract dependent cell addresses from formula
export function extractDependencies(rawFormula: string): string[] {
  if (!rawFormula || !rawFormula.startsWith('=')) return [];
  const deps: Set<string> = new Set();
  const tokens = tokenize(rawFormula);

  for (const t of tokens) {
    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (p) deps.add(cellToA1(p.r, p.c));
    } else if (t.type === 'RANGE') {
      const rng = parseRange(t.value);
      if (rng) {
        for (let r = rng.r1; r <= rng.r2; r++) {
          for (let c = rng.c1; c <= rng.c2; c++) {
            deps.add(cellToA1(r, c));
          }
        }
      }
    }
  }

  return Array.from(deps);
}

// Evaluator Class
export class Evaluator {
  private sheet: Sheet;
  private cache: Record<string, CellValue> = {};
  private calculating: Set<string> = new Set();
  private depGraph: Map<string, Set<string>> = new Map(); // cell -> cells it depends on
  private revDepGraph: Map<string, Set<string>> = new Map(); // cell -> cells that depend on it

  constructor(sheet: Sheet) {
    this.sheet = sheet;
    this.buildDepGraph();
  }

  public setSheet(sheet: Sheet) {
    this.sheet = sheet;
    this.cache = {};
    this.calculating.clear();
    this.buildDepGraph();
  }

  private buildDepGraph() {
    this.depGraph.clear();
    this.revDepGraph.clear();

    for (const [cellKey, cell] of Object.entries(this.sheet.cells)) {
      if (cell.raw && cell.raw.startsWith('=')) {
        const deps = extractDependencies(cell.raw);
        this.depGraph.set(cellKey, new Set(deps));
        for (const dep of deps) {
          if (!this.revDepGraph.has(dep)) {
            this.revDepGraph.set(dep, new Set());
          }
          this.revDepGraph.get(dep)!.add(cellKey);
        }
      }
    }
  }

  public getCellValue(a1Key: string): CellValue {
    const key = a1Key.toUpperCase();
    if (key in this.cache) {
      return this.cache[key];
    }

    if (this.calculating.has(key)) {
      return '#CYCLE!';
    }

    const cell = this.sheet.cells[key];
    if (!cell || cell.raw === '' || cell.raw === undefined || cell.raw === null) {
      this.cache[key] = '';
      return '';
    }

    const raw = cell.raw.trim();

    // Plain text escaping with leading single quote: 'hello
    if (raw.startsWith("'")) {
      const val = raw.slice(1);
      this.cache[key] = val;
      return val;
    }

    // Formulas start with '='
    if (raw.startsWith('=')) {
      this.calculating.add(key);
      try {
        const val = this.evaluateFormula(raw);
        this.cache[key] = val;
        return val;
      } catch (err: any) {
        const errCode = isCellError(err.message) ? err.message : '#ERR';
        this.cache[key] = errCode;
        return errCode;
      } finally {
        this.calculating.delete(key);
      }
    }

    // Number or plain string
    const cleaned = raw.replace(/,/g, '');
    if (cleaned !== '' && !isNaN(Number(cleaned))) {
      const num = Number(cleaned);
      this.cache[key] = num;
      return num;
    }

    this.cache[key] = raw;
    return raw;
  }

  public evaluateFormula(formulaText: string): CellValue {
    try {
      const tokens = tokenize(formulaText);
      const parser = new FormulaParser(tokens);
      const ast = parser.parse();
      return this.evalAST(ast);
    } catch (err: any) {
      if (isCellError(err.message)) return err.message;
      return '#ERR';
    }
  }

  private evalAST(node: ASTNode): any {
    switch (node.type) {
      case 'Number':
        return node.value;

      case 'String':
        return node.value;

      case 'Boolean':
        return node.value;

      case 'CellRef': {
        const val = this.getCellValue(node.a1);
        return val === '' ? 0 : val;
      }

      case 'Range': {
        const rng = parseRange(node.rangeStr);
        if (!rng) return '#REF!';
        const values: CellValue[] = [];
        for (let r = rng.r1; r <= rng.r2; r++) {
          for (let c = rng.c1; c <= rng.c2; c++) {
            values.push(this.getCellValue(cellToA1(r, c)));
          }
        }
        return values;
      }

      case 'UnaryOp': {
        const val = this.evalAST(node.expr);
        if (isCellError(val)) return val;
        const num = Number(val);
        if (isNaN(num)) return '#VALUE!';
        return node.op === '-' ? -num : num;
      }

      case 'PostfixOp': {
        if (node.op === '%') {
          const val = this.evalAST(node.expr);
          if (isCellError(val)) return val;
          const num = Number(val);
          if (isNaN(num)) return '#VALUE!';
          return normalizeFloat(num / 100);
        }
        return '#ERR';
      }

      case 'BinaryOp': {
        const left = this.evalAST(node.left);
        const right = this.evalAST(node.right);
        if (isCellError(left)) return left;
        if (isCellError(right)) return right;

        if (node.op === '&') {
          return `${left ?? ''}${right ?? ''}`;
        }

        // Comparisons
        if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
          return this.evalComparison(node.op, left, right);
        }

        // Arithmetic
        const lNum = Number(left);
        const rNum = Number(right);
        if (isNaN(lNum) || isNaN(rNum)) return '#VALUE!';

        switch (node.op) {
          case '+':
            return normalizeFloat(lNum + rNum);
          case '-':
            return normalizeFloat(lNum - rNum);
          case '*':
            return normalizeFloat(lNum * rNum);
          case '/':
            if (rNum === 0) return '#DIV/0!';
            return normalizeFloat(lNum / rNum);
          case '^':
            return normalizeFloat(Math.pow(lNum, rNum));
          default:
            return '#ERR';
        }
      }

      case 'FunctionCall': {
        const fn = FUNCTIONS[node.name];
        if (!fn) return '#NAME?';

        // Evaluate arguments
        const evaluatedArgs: any[] = [];
        for (const argNode of node.args) {
          evaluatedArgs.push(this.evalAST(argNode));
        }

        return fn(...evaluatedArgs);
      }
    }
  }

  private evalComparison(op: string, left: any, right: any): boolean {
    const l = typeof left === 'string' ? left.toLowerCase() : left;
    const r = typeof right === 'string' ? right.toLowerCase() : right;

    switch (op) {
      case '=':
        return l === r;
      case '<>':
        return l !== r;
      case '<':
        return l < r;
      case '>':
        return l > r;
      case '<=':
        return l <= r;
      case '>=':
        return l >= r;
      default:
        return false;
    }
  }

  // Invalidate cache for a cell and all dependent cells
  public invalidate(cellKey: string) {
    const key = cellKey.toUpperCase();
    const toInvalidate: Set<string> = new Set([key]);
    const queue = [key];

    while (queue.length > 0) {
      const current = queue.shift()!;
      const dependents = this.revDepGraph.get(current);
      if (dependents) {
        for (const dep of dependents) {
          if (!toInvalidate.has(dep)) {
            toInvalidate.add(dep);
            queue.push(dep);
          }
        }
      }
    }

    for (const cell of toInvalidate) {
      delete this.cache[cell];
    }

    // Rebuild graph for changed cell if needed
    const raw = this.sheet.cells[key]?.raw;
    if (raw && raw.startsWith('=')) {
      const deps = extractDependencies(raw);
      this.depGraph.set(key, new Set(deps));
      for (const dep of deps) {
        if (!this.revDepGraph.has(dep)) {
          this.revDepGraph.set(dep, new Set());
        }
        this.revDepGraph.get(dep)!.add(key);
      }
    } else {
      this.depGraph.delete(key);
    }
  }
}
