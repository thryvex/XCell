import { Workbook, Sheet, createEmptySheet, CellStyle } from '../core/model';

/**
 * Creates the "รายรับรายจ่าย" (Income & Expense Ledger) template workbook (SPEC 7 & Phase 3D)
 */
export function createLedgerWorkbook(): Workbook {
  const sheet: Sheet = createEmptySheet('บันทึก', 100, 26);
  sheet.frozenRows = 1;

  // Column widths
  sheet.colWidths = {
    0: 120, // วันที่
    1: 220, // รายการ (กว้างกว่าคอลัมน์อื่น)
    2: 130, // หมวด
    3: 130, // รายรับ
    4: 130, // รายจ่าย
    5: 130, // คงเหลือ
    6: 40,  // ช่องว่างคั่น
    7: 120, // หัวข้อสรุป
    8: 140, // ค่ายอดรวม
  };

  const headerStyle: CellStyle = {
    bold: true,
    fill: '#e2e8f0',
    align: 'center',
    border: 'all',
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

  // 3 Sample Rows (Rows 2-4)
  sheet.cells['A2'] = { raw: '1/10/2026', fmt: 'date' };
  sheet.cells['B2'] = { raw: 'เงินเดือน' };
  sheet.cells['C2'] = { raw: 'รายรับ' };
  sheet.cells['D2'] = { raw: '25000', fmt: 'baht' };
  sheet.cells['E2'] = { raw: '0', fmt: 'baht' };
  sheet.cells['F2'] = { raw: '=D2-E2', fmt: 'baht' };

  sheet.cells['A3'] = { raw: '2/10/2026', fmt: 'date' };
  sheet.cells['B3'] = { raw: 'ค่าอาหาร' };
  sheet.cells['C3'] = { raw: 'อาหาร' };
  sheet.cells['D3'] = { raw: '0', fmt: 'baht' };
  sheet.cells['E3'] = { raw: '150', fmt: 'baht' };
  sheet.cells['F3'] = { raw: '=F2+D3-E3', fmt: 'baht' };

  sheet.cells['A4'] = { raw: '3/10/2026', fmt: 'date' };
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

  return {
    sheets: [sheet],
    activeSheet: 0,
  };
}
