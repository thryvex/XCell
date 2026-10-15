import { Workbook, Sheet, createEmptySheet, CellStyle } from '../core/model';
import { dateToSerial } from '../core/numfmt';
import { buildTemplateWorkbook } from './workbook';

export const LEDGER_SHEET_NAME = 'บันทึก';
export const SUMMARY_SHEET_NAME = 'สรุป';

/** หมวดตัวอย่างในตาราง "รวมตามหมวด" (ข้อมูลตัวอย่างในชีตบันทึกใช้หมวดเหล่านี้) */
export const SUMMARY_CATEGORIES = ['อาหาร', 'เดินทาง', 'ที่พัก', 'อื่นๆ'];
const EXTRA_CATEGORY_ROWS = 6;

export const THAI_MONTHS = [
  'มกราคม',
  'กุมภาพันธ์',
  'มีนาคม',
  'เมษายน',
  'พฤษภาคม',
  'มิถุนายน',
  'กรกฎาคม',
  'สิงหาคม',
  'กันยายน',
  'ตุลาคม',
  'พฤศจิกายน',
  'ธันวาคม',
];

const headerStyle: CellStyle = {
  bold: true,
  fill: '#e2e8f0',
  align: 'center',
  border: 'all',
};

const totalStyle: CellStyle = { bold: true, fill: '#dbeafe', border: 'all' };

/**
 * ชีต "บันทึก" (ตาม 3D) ข้อมูลตัวอย่างอยู่ในเดือนปัจจุบันของปีนี้ และหมวดตรงกับตารางสรุป
 */
function createLedgerSheet(today: Date): Sheet {
  const sheet: Sheet = createEmptySheet(LEDGER_SHEET_NAME, 100, 26);
  sheet.frozenRows = 1;

  // Column widths
  sheet.colWidths = {
    0: 120, // วันที่
    1: 220, // รายการ (กว้างกว่าคอลัมน์อื่น)
    2: 130, // หมวด
    3: 130, // รายรับ
    4: 130, // รายจ่าย
    5: 130, // คงเหลือ
    6: 40, // ช่องว่างคั่น
    7: 120, // หัวข้อสรุป
    8: 140, // ค่ายอดรวม
  };

  // Header row A1:F1
  sheet.cells['A1'] = { raw: 'วันที่', style: headerStyle };
  sheet.cells['B1'] = { raw: 'รายการ', style: headerStyle };
  sheet.cells['C1'] = { raw: 'หมวด', style: headerStyle };
  sheet.cells['D1'] = { raw: 'รายรับ', style: headerStyle };
  sheet.cells['E1'] = { raw: 'รายจ่าย', style: headerStyle };
  sheet.cells['F1'] = { raw: 'คงเหลือ', style: headerStyle };

  // Set number formats for rows 2 to 1000
  for (let r = 2; r <= 1000; r++) {
    sheet.cells[`A${r}`] = { raw: '', fmt: 'date' };
    sheet.cells[`D${r}`] = { raw: '', fmt: 'baht' };
    sheet.cells[`E${r}`] = { raw: '', fmt: 'baht' };
    sheet.cells[`F${r}`] = { raw: '', fmt: 'baht' };
  }

  // วันที่ตัวอย่าง: วันที่ 1–3 ของเดือนปัจจุบัน ปีนี้ (เก็บเป็นวันที่จริง)
  const sampleDate = (day: number) =>
    String(dateToSerial(new Date(today.getFullYear(), today.getMonth(), day)));

  // 3 Sample Rows (Rows 2-4)
  sheet.cells['A2'] = { raw: sampleDate(1), fmt: 'date' };
  sheet.cells['B2'] = { raw: 'เงินเดือน' };
  sheet.cells['C2'] = { raw: 'อื่นๆ' };
  sheet.cells['D2'] = { raw: '25000', fmt: 'baht' };
  sheet.cells['E2'] = { raw: '0', fmt: 'baht' };
  sheet.cells['F2'] = { raw: '=D2-E2', fmt: 'baht' };

  sheet.cells['A3'] = { raw: sampleDate(2), fmt: 'date' };
  sheet.cells['B3'] = { raw: 'ค่าอาหาร' };
  sheet.cells['C3'] = { raw: 'อาหาร' };
  sheet.cells['D3'] = { raw: '0', fmt: 'baht' };
  sheet.cells['E3'] = { raw: '150', fmt: 'baht' };
  sheet.cells['F3'] = { raw: '=F2+D3-E3', fmt: 'baht' };

  sheet.cells['A4'] = { raw: sampleDate(3), fmt: 'date' };
  sheet.cells['B4'] = { raw: 'ค่าเดินทาง' };
  sheet.cells['C4'] = { raw: 'เดินทาง' };
  sheet.cells['D4'] = { raw: '0', fmt: 'baht' };
  sheet.cells['E4'] = { raw: '80', fmt: 'baht' };
  sheet.cells['F4'] = { raw: '=F3+D4-E4', fmt: 'baht' };

  // Summary section on the right (H1:I3)
  sheet.cells['H1'] = { raw: 'รวมรายรับ', style: { bold: true, align: 'right' } };
  sheet.cells['I1'] = {
    raw: '=SUM(D2:D1000)',
    fmt: 'baht',
    style: { bold: true, border: 'all' },
  };

  sheet.cells['H2'] = { raw: 'รวมรายจ่าย', style: { bold: true, align: 'right' } };
  sheet.cells['I2'] = {
    raw: '=SUM(E2:E1000)',
    fmt: 'baht',
    style: { bold: true, border: 'all' },
  };

  sheet.cells['H3'] = { raw: 'คงเหลือ', style: { bold: true, align: 'right' } };
  sheet.cells['I3'] = {
    raw: '=I1-I2',
    fmt: 'baht',
    style: { bold: true, fill: '#dbeafe', border: 'all' },
  };

  return sheet;
}

// ช่วงข้อมูลในชีตบันทึก (ใส่ชื่อชีตในเครื่องหมายคำพูดตามกติกา)
const LEDGER = `'${LEDGER_SHEET_NAME}'`;
const DATES = `${LEDGER}!$A$2:$A$1000`;
const CATEGORIES = `${LEDGER}!$C$2:$C$1000`;
const INCOME = `${LEDGER}!$D$2:$D$1000`;
const EXPENSE = `${LEDGER}!$E$2:$E$1000`;

// ตำแหน่งตารางในชีตสรุป (เลขแถวตามที่เห็นบนจอ)
export const SUMMARY_LAYOUT = {
  yearCell: 'B1',
  categoryHeaderRow: 3,
  categoryFirstRow: 4,
  categoryLastRow: 4 + SUMMARY_CATEGORIES.length + EXTRA_CATEGORY_ROWS - 1, // 13
  categoryTotalRow: 4 + SUMMARY_CATEGORIES.length + EXTRA_CATEGORY_ROWS, // 14
  monthHeaderRow: 17,
  monthFirstRow: 18,
  monthTotalRow: 30,
};

/**
 * ชีต "สรุป": ปี, ตารางรวมตามหมวด, ตารางรายเดือน (ใช้ SUMIFS อ้างชีตบันทึก)
 */
function createSummarySheet(): Sheet {
  const sheet: Sheet = createEmptySheet(SUMMARY_SHEET_NAME, 100, 26);
  const L = SUMMARY_LAYOUT;
  sheet.frozenRows = 1;
  sheet.colWidths = { 0: 120, 1: 140, 2: 130, 3: 130, 4: 130 };

  // a) ปี (ผู้ใช้พิมพ์ทับเป็นปีอื่นได้)
  sheet.cells['A1'] = { raw: 'ปี', style: { bold: true } };
  sheet.cells['B1'] = { raw: '=YEAR(TODAY())', style: { bold: true, fill: '#fef9c3', border: 'all' } };

  // b) รวมตามหมวด
  sheet.cells['A2'] = { raw: 'รวมตามหมวด', style: { bold: true } };
  sheet.cells[`A${L.categoryHeaderRow}`] = { raw: 'หมวด', style: headerStyle };
  sheet.cells[`B${L.categoryHeaderRow}`] = { raw: 'รายรับ', style: headerStyle };
  sheet.cells[`C${L.categoryHeaderRow}`] = { raw: 'รายจ่าย', style: headerStyle };

  for (let r = L.categoryFirstRow; r <= L.categoryLastRow; r++) {
    const category = SUMMARY_CATEGORIES[r - L.categoryFirstRow];
    sheet.cells[`A${r}`] = { raw: category ?? '', style: { border: 'all' } };
    // หมวดว่าง → ผลลัพธ์ว่าง (กัน SUMIFS รวมแถวที่หมวดว่างทั้งหมด)
    sheet.cells[`B${r}`] = {
      raw: `=IF(A${r}="","",SUMIFS(${INCOME},${CATEGORIES},A${r}))`,
      fmt: 'baht',
      style: { border: 'all' },
    };
    sheet.cells[`C${r}`] = {
      raw: `=IF(A${r}="","",SUMIFS(${EXPENSE},${CATEGORIES},A${r}))`,
      fmt: 'baht',
      style: { border: 'all' },
    };
  }

  const ct = L.categoryTotalRow;
  sheet.cells[`A${ct}`] = { raw: 'รวม', style: totalStyle };
  sheet.cells[`B${ct}`] = {
    raw: `=SUM(B${L.categoryFirstRow}:B${L.categoryLastRow})`,
    fmt: 'baht',
    style: totalStyle,
  };
  sheet.cells[`C${ct}`] = {
    raw: `=SUM(C${L.categoryFirstRow}:C${L.categoryLastRow})`,
    fmt: 'baht',
    style: totalStyle,
  };

  // c) รายเดือน (คอลัมน์ A = เลขเดือนที่ใช้ในสูตร, B = ชื่อเดือน)
  sheet.cells[`A${L.monthHeaderRow - 1}`] = { raw: 'รายเดือน', style: { bold: true } };
  const monthHeaders = ['เดือนที่', 'เดือน', 'รายรับ', 'รายจ่าย', 'คงเหลือ'];
  monthHeaders.forEach((h, i) => {
    sheet.cells[`${'ABCDE'[i]}${L.monthHeaderRow}`] = { raw: h, style: headerStyle };
  });

  THAI_MONTHS.forEach((month, i) => {
    const r = L.monthFirstRow + i;
    const from = `">="&DATE($B$1,A${r},1)`;
    const to = `"<"&DATE($B$1,A${r}+1,1)`;
    sheet.cells[`A${r}`] = {
      raw: String(i + 1),
      style: { align: 'center', color: '#64748b', border: 'all' },
    };
    sheet.cells[`B${r}`] = { raw: month, style: { border: 'all' } };
    sheet.cells[`C${r}`] = {
      raw: `=SUMIFS(${INCOME},${DATES},${from},${DATES},${to})`,
      fmt: 'baht',
      style: { border: 'all' },
    };
    sheet.cells[`D${r}`] = {
      raw: `=SUMIFS(${EXPENSE},${DATES},${from},${DATES},${to})`,
      fmt: 'baht',
      style: { border: 'all' },
    };
    sheet.cells[`E${r}`] = { raw: `=C${r}-D${r}`, fmt: 'baht', style: { border: 'all' } };
  });

  const mt = L.monthTotalRow;
  const lastMonthRow = L.monthFirstRow + 11;
  sheet.cells[`A${mt}`] = { raw: '', style: totalStyle };
  sheet.cells[`B${mt}`] = { raw: 'รวมทั้งปี', style: totalStyle };
  for (const col of ['C', 'D', 'E']) {
    sheet.cells[`${col}${mt}`] = {
      raw: `=SUM(${col}${L.monthFirstRow}:${col}${lastMonthRow})`,
      fmt: 'baht',
      style: totalStyle,
    };
  }

  return sheet;
}

/**
 * Creates the "รายรับรายจ่าย" (Income & Expense Ledger) template workbook (SPEC 7, Phase 3D + 4D)
 * ชีต "บันทึก" + ชีต "สรุป"; เปิดมาที่ชีตบันทึก
 */
export function createLedgerWorkbook(today: Date = new Date()): Workbook {
  return buildTemplateWorkbook([createLedgerSheet(today), createSummarySheet()], 0);
}
