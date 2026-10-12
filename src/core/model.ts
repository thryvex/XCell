export type CellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // #RRGGBB ตัวอักษร
  fill?: string; // #RRGGBB พื้นเซล
  align?: 'left' | 'center' | 'right';
  wrap?: boolean;
  border?: 'none' | 'all' | 'outer';
};

export type NumFmt =
  | 'general'
  | 'number2'
  | 'integer'
  | 'thousands'
  | 'baht'
  | 'percent'
  | 'date';

export type Cell = {
  raw: string; // สิ่งที่ผู้ใช้พิมพ์ ถ้าขึ้นต้นด้วย "=" คือสูตร
  fmt?: NumFmt;
  style?: CellStyle;
};

export type MergeRange = {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}; // 0-based: r1, c1 (top-left) to r2, c2 (bottom-right)

export type Sheet = {
  name: string;
  cells: Record<string, Cell>; // key = "B3"
  merges: MergeRange[];
  colWidths: Record<number, number>; // 0-based col index -> width in px
  rowHeights: Record<number, number>; // 0-based row index -> height in px
  frozenRows: number;
  frozenCols: number;
};

export type Workbook = {
  sheets: Sheet[];
  activeSheet: number;
};

export type CellError = '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#CYCLE!' | '#ERR';

export type CellValue = string | number | boolean | CellError;

export function createEmptySheet(name: string = 'Sheet1', initialRows = 100, initialCols = 26): Sheet {
  return {
    name,
    cells: {},
    merges: [],
    colWidths: {},
    rowHeights: {},
    frozenRows: 0,
    frozenCols: 0,
  };
}

export function createEmptyWorkbook(): Workbook {
  return {
    sheets: [createEmptySheet('Sheet1')],
    activeSheet: 0,
  };
}
