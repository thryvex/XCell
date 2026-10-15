import { CellValue, Sheet, Workbook } from './model';
import { tokenize, Token } from './tokenizer';
import { parseA1, parseRange, cellToA1 } from './address';
import { FUNCTIONS, isCellError } from './functions';

// AST Nodes
export type ASTNode =
  | { type: 'Number'; value: number }
  | { type: 'String'; value: string }
  | { type: 'Boolean'; value: boolean }
  | { type: 'Error'; value: string }
  | { type: 'CellRef'; a1: string; sheet?: string }
  | { type: 'Range'; rangeStr: string; sheet?: string }
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

    if (t.type === 'ERROR') {
      this.consume();
      return { type: 'Error', value: t.value };
    }

    if (t.type === 'CELL_REF') {
      this.consume();
      return { type: 'CellRef', a1: t.value.toUpperCase(), sheet: t.sheet };
    }

    if (t.type === 'RANGE') {
      this.consume();
      return { type: 'Range', rangeStr: t.value.toUpperCase(), sheet: t.sheet };
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

/** คืน index ของชีตจากชื่อ (ไม่สนตัวพิมพ์เล็ก/ใหญ่) หรือ -1 ถ้าไม่มี */
export type SheetResolver = (name: string) => number;

/** key ของเซลในกราฟ dependency ระดับสมุดงาน เช่น "0!A1" */
export function qualifiedKey(sheetIndex: number, a1: string): string {
  return `${sheetIndex}!${a1.toUpperCase()}`;
}

// Extract dependent cell addresses (sheet-qualified) from formula
export function extractDependencies(
  rawFormula: string,
  sheetIndex = 0,
  resolveSheet?: SheetResolver
): string[] {
  if (!rawFormula || !rawFormula.startsWith('=')) return [];
  const deps: Set<string> = new Set();
  const tokens = tokenize(rawFormula);

  for (const t of tokens) {
    if (t.type !== 'CELL_REF' && t.type !== 'RANGE') continue;

    let targetIndex = sheetIndex;
    if (t.sheet !== undefined) {
      targetIndex = resolveSheet ? resolveSheet(t.sheet) : -1;
      if (targetIndex < 0) continue;
    }

    if (t.type === 'CELL_REF') {
      const p = parseA1(t.value);
      if (p) deps.add(qualifiedKey(targetIndex, cellToA1(p.r, p.c)));
    } else {
      const rng = parseRange(t.value);
      if (rng) {
        for (let r = rng.r1; r <= rng.r2; r++) {
          for (let c = rng.c1; c <= rng.c2; c++) {
            deps.add(qualifiedKey(targetIndex, cellToA1(r, c)));
          }
        }
      }
    }
  }

  return Array.from(deps);
}

function isWorkbook(source: Sheet | Workbook): source is Workbook {
  return Array.isArray((source as Workbook).sheets);
}

// Evaluator Class: คำนวณทั้งสมุดงาน (ทุกชีต รวมชีตซ่อน) ด้วย dependency graph ข้ามชีต
export class Evaluator {
  private wb: Workbook;
  private cache: Map<string, CellValue> = new Map();
  private calculating: Set<string> = new Set();
  private revDepGraph: Map<string, Set<string>> = new Map(); // cell -> cells that depend on it
  private sheetIndexByName: Map<string, number> = new Map();

  constructor(source: Sheet | Workbook) {
    this.wb = isWorkbook(source) ? source : { sheets: [source], activeSheet: 0 };
    this.rebuild();
  }

  /** ใช้เมื่อโครงสร้างสมุดงานเปลี่ยน (เพิ่ม/ลบ/เปลี่ยนชื่อ/ย้ายชีต, Undo/Redo, แทรก/ลบแถว) */
  public setWorkbook(wb: Workbook) {
    this.wb = wb;
    this.rebuild();
  }

  /** รูปแบบเดิมสำหรับสมุดงานชีตเดียว */
  public setSheet(sheet: Sheet) {
    this.setWorkbook({ sheets: [sheet], activeSheet: 0 });
  }

  public resolveSheet = (name: string): number => {
    const idx = this.sheetIndexByName.get(name.toLowerCase());
    return idx === undefined ? -1 : idx;
  };

  private rebuild() {
    this.cache.clear();
    this.calculating.clear();
    this.revDepGraph.clear();
    this.sheetIndexByName.clear();

    this.wb.sheets.forEach((sheet, idx) => {
      this.sheetIndexByName.set(sheet.name.toLowerCase(), idx);
    });

    this.wb.sheets.forEach((sheet, idx) => {
      for (const [cellKey, cell] of Object.entries(sheet.cells)) {
        if (cell.raw && cell.raw.startsWith('=')) {
          this.addDependencies(idx, cellKey, cell.raw);
        }
      }
    });
  }

  private addDependencies(sheetIndex: number, a1: string, raw: string) {
    const key = qualifiedKey(sheetIndex, a1);
    for (const dep of extractDependencies(raw, sheetIndex, this.resolveSheet)) {
      if (!this.revDepGraph.has(dep)) {
        this.revDepGraph.set(dep, new Set());
      }
      this.revDepGraph.get(dep)!.add(key);
    }
  }

  public getCellValue(a1Key: string, sheetIndex = 0): CellValue {
    const a1 = a1Key.toUpperCase();
    const key = qualifiedKey(sheetIndex, a1);
    const cached = this.cache.get(key);
    if (cached !== undefined) {
      return cached;
    }

    if (this.calculating.has(key)) {
      return '#CYCLE!';
    }

    const cell = this.wb.sheets[sheetIndex]?.cells[a1];
    if (!cell || cell.raw === '' || cell.raw === undefined || cell.raw === null) {
      this.cache.set(key, '');
      return '';
    }

    // Preserve cached value for unsupported formulas
    if (cell.unsupportedFormula && cell.cachedValue !== undefined) {
      this.cache.set(key, cell.cachedValue);
      return cell.cachedValue;
    }

    // Plain text escaping with leading single quote: 'hello
    if (cell.raw.startsWith("'")) {
      const val = cell.raw.slice(1);
      this.cache.set(key, val);
      return val;
    }

    const trimmed = cell.raw.trim();

    // Formulas start with '='
    if (trimmed.startsWith('=')) {
      this.calculating.add(key);
      try {
        const val = this.evaluateFormula(trimmed, sheetIndex);
        this.cache.set(key, val);
        return val;
      } catch (err: any) {
        const errCode = isCellError(err.message) ? err.message : '#ERR';
        this.cache.set(key, errCode);
        return errCode;
      } finally {
        this.calculating.delete(key);
      }
    }

    // Number or plain string
    const cleaned = trimmed.replace(/,/g, '');
    if (cleaned !== '' && !isNaN(Number(cleaned))) {
      const num = Number(cleaned);
      this.cache.set(key, num);
      return num;
    }

    this.cache.set(key, cell.raw);
    return cell.raw;
  }

  public evaluateFormula(formulaText: string, sheetIndex = 0): CellValue {
    try {
      const tokens = tokenize(formulaText);
      const parser = new FormulaParser(tokens);
      const ast = parser.parse();
      return this.evalAST(ast, sheetIndex);
    } catch (err: any) {
      if (isCellError(err.message)) return err.message;
      return '#ERR';
    }
  }

  private targetSheet(refSheet: string | undefined, sheetIndex: number): number {
    if (refSheet === undefined) return sheetIndex;
    return this.resolveSheet(refSheet);
  }

  private evalAST(node: ASTNode, sheetIndex: number): any {
    switch (node.type) {
      case 'Number':
        return node.value;

      case 'String':
        return node.value;

      case 'Boolean':
        return node.value;

      case 'Error':
        return node.value;

      case 'CellRef': {
        const target = this.targetSheet(node.sheet, sheetIndex);
        if (target < 0) return '#REF!';
        return this.getCellValue(node.a1, target);
      }

      case 'Range': {
        const target = this.targetSheet(node.sheet, sheetIndex);
        if (target < 0) return '#REF!';
        const rng = parseRange(node.rangeStr);
        if (!rng) return '#REF!';
        const values: CellValue[] = [];
        for (let r = rng.r1; r <= rng.r2; r++) {
          for (let c = rng.c1; c <= rng.c2; c++) {
            values.push(this.getCellValue(cellToA1(r, c), target));
          }
        }
        return values;
      }

      case 'UnaryOp': {
        const val = this.evalAST(node.expr, sheetIndex);
        if (isCellError(val)) return val;
        const num = Number(val);
        if (isNaN(num)) return '#VALUE!';
        return node.op === '-' ? -num : num;
      }

      case 'PostfixOp': {
        if (node.op === '%') {
          const val = this.evalAST(node.expr, sheetIndex);
          if (isCellError(val)) return val;
          const num = Number(val);
          if (isNaN(num)) return '#VALUE!';
          return normalizeFloat(num / 100);
        }
        return '#ERR';
      }

      case 'BinaryOp': {
        const left = this.evalAST(node.left, sheetIndex);
        const right = this.evalAST(node.right, sheetIndex);
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
          evaluatedArgs.push(this.evalAST(argNode, sheetIndex));
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

  // Invalidate cache for a cell and all dependent cells (รวมเซลในชีตอื่นที่อ้างถึง)
  public invalidate(cellKey: string, sheetIndex = 0) {
    const a1 = cellKey.toUpperCase();
    const key = qualifiedKey(sheetIndex, a1);
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
      this.cache.delete(cell);
    }

    // Register dependencies of the changed cell
    const raw = this.wb.sheets[sheetIndex]?.cells[a1]?.raw;
    if (raw && raw.startsWith('=')) {
      this.addDependencies(sheetIndex, a1, raw);
    }
  }
}
