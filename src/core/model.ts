export type CellStyle = {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // #RRGGBB ตัวอักษร
  fill?: string; // #RRGGBB พื้นเซล
  align?: 'left' | 'center' | 'right';
  wrap?: boolean;
  border?: 'none' | 'all' | 'outer';
  fontSize?: number; // สำหรับเก็บรักษาเมื่อเปิดจากไฟล์ภายนอก
  fontFamily?: string; // สำหรับเก็บรักษาเมื่อเปิดจากไฟล์ภายนอก
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
  unsupportedFormula?: boolean;
  cachedValue?: CellValue;
};

export type MergeRange = {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}; // 0-based: r1, c1 (top-left) to r2, c2 (bottom-right)

export type Sheet = {
  id?: string; // รหัสภายในแอป (ไม่บันทึกลง .xlsx) ใช้ผูกสถานะการมองเห็นต่อชีต
  name: string;
  hidden?: 'hidden' | 'veryHidden'; // ชีตซ่อนจากไฟล์ .xlsx: ไม่แสดงแท็บ แต่เก็บข้อมูลไว้ตามเดิม
  cells: Record<string, Cell>; // key = "B3"
  merges: MergeRange[];
  colWidths: Record<number, number>; // 0-based col index -> width in px
  rowHeights: Record<number, number>; // 0-based row index -> height in px
  frozenRows: number;
  frozenCols: number;
  filter?: SheetFilter; // ตัวกรอง (4C)
  hiddenRows?: number[]; // แถวที่ซ่อนจากตัวกรอง (0-based, เรียงจากน้อยไปมาก)
  zoom?: number; // ระดับซูมเป็นเปอร์เซ็นต์ (ไม่มี = 100)
};

/** เงื่อนไขกรองแบบ "กรองตามเงื่อนไข" */
export type FilterCondition =
  | { type: 'empty' }
  | { type: 'notEmpty' }
  | { type: 'textContains'; value: string }
  | { type: 'textNotContains'; value: string }
  | { type: 'startsWith'; value: string }
  | { type: 'equals'; value: string }
  | { type: 'greaterThan'; value: string }
  | { type: 'lessThan'; value: string }
  | { type: 'between'; min: string; max: string };

/** ตัวกรองของคอลัมน์หนึ่ง: กรองตามค่า (ค่าที่ซ่อน) หรือกรองตามเงื่อนไข อย่างใดอย่างหนึ่ง */
export type ColumnFilter =
  | { kind: 'values'; hiddenValues: string[] }
  | { kind: 'condition'; condition: FilterCondition };

/**
 * ตัวกรองของชีต: r1 คือแถวหัวตาราง, ข้อมูลคือแถว r1+1..r2
 * columns: key = index คอลัมน์จริง (0-based)
 */
export type SheetFilter = {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
  columns: Record<number, ColumnFilter>;
};

export type Workbook = {
  sheets: Sheet[];
  activeSheet: number;
  unsupportedFeatures?: string[];
};

export type CellError = '#DIV/0!' | '#VALUE!' | '#REF!' | '#NAME?' | '#CYCLE!' | '#ERR';

export type CellValue = string | number | boolean | CellError;

let sheetIdCounter = 0;

export function newSheetId(): string {
  sheetIdCounter++;
  return `s${Date.now().toString(36)}-${sheetIdCounter}`;
}

/**
 * ใส่ id ให้ทุกชีตที่ยังไม่มี (ไฟล์ที่เปิดมา/ไฟล์กู้คืนรุ่นเก่า)
 */
export function ensureSheetIds(wb: Workbook): Workbook {
  for (const sheet of wb.sheets) {
    if (!sheet.id) sheet.id = newSheetId();
  }
  return wb;
}

export function createEmptySheet(name: string = 'Sheet1', initialRows = 100, initialCols = 26): Sheet {
  return {
    id: newSheetId(),
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
    sheets: [createEmptySheet('ชีต1')],
    activeSheet: 0,
  };
}
