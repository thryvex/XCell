import { Workbook, Sheet, createEmptySheet, CellStyle } from '../core/model';

/**
 * Creates the "รายการสินค้า" (Product Inventory) template workbook (SPEC 7 & Phase 3D)
 */
export function createInventoryWorkbook(): Workbook {
  const sheet: Sheet = createEmptySheet('รายการสินค้า', 100, 26);
  sheet.frozenRows = 1;

  // Column widths
  sheet.colWidths = {
    0: 80,  // ลำดับ
    1: 220, // สินค้า (กว้างกว่าคอลัมน์อื่น)
    2: 100, // จำนวน
    3: 140, // ราคาต่อหน่วย
    4: 140, // รวม
    5: 40,  // ช่องว่างคั่น
    6: 120, // หัวข้อสรุป
    7: 150, // ค่ายอดรวม
  };

  const headerStyle: CellStyle = {
    bold: true,
    fill: '#e2e8f0',
    align: 'center',
    border: 'all',
  };

  // Header row A1:E1
  sheet.cells['A1'] = { raw: 'ลำดับ', style: headerStyle };
  sheet.cells['B1'] = { raw: 'สินค้า', style: headerStyle };
  sheet.cells['C1'] = { raw: 'จำนวน', style: headerStyle };
  sheet.cells['D1'] = { raw: 'ราคาต่อหน่วย', style: headerStyle };
  sheet.cells['E1'] = { raw: 'รวม', style: headerStyle };

  // Set number formats for rows 2 to 1000
  for (let r = 2; r <= 1000; r++) {
    sheet.cells[`D${r}`] = { raw: '', fmt: 'baht' };
    sheet.cells[`E${r}`] = { raw: '', fmt: 'baht' };
  }

  // 3 Sample Rows (Rows 2-4)
  sheet.cells['A2'] = { raw: '1' };
  sheet.cells['B2'] = { raw: 'สมุดบันทึก' };
  sheet.cells['C2'] = { raw: '5' };
  sheet.cells['D2'] = { raw: '35', fmt: 'baht' };
  sheet.cells['E2'] = { raw: '=C2*D2', fmt: 'baht' };

  sheet.cells['A3'] = { raw: '2' };
  sheet.cells['B3'] = { raw: 'ปากกาลูกลื่น' };
  sheet.cells['C3'] = { raw: '10' };
  sheet.cells['D3'] = { raw: '15', fmt: 'baht' };
  sheet.cells['E3'] = { raw: '=C3*D3', fmt: 'baht' };

  sheet.cells['A4'] = { raw: '3' };
  sheet.cells['B4'] = { raw: 'แฟ้มเอกสาร' };
  sheet.cells['C4'] = { raw: '3' };
  sheet.cells['D4'] = { raw: '45', fmt: 'baht' };
  sheet.cells['E4'] = { raw: '=C4*D4', fmt: 'baht' };

  // Summary section on the right (G1:H1)
  sheet.cells['G1'] = { raw: 'รวมทั้งหมด', style: { bold: true, align: 'right' } };
  sheet.cells['H1'] = {
    raw: '=SUM(E2:E1000)',
    fmt: 'baht',
    style: { bold: true, fill: '#dbeafe', border: 'all' },
  };

  return {
    sheets: [sheet],
    activeSheet: 0,
  };
}
