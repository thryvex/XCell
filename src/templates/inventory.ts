import { Workbook, Sheet, createEmptySheet, CellStyle } from '../core/model';
import { buildTemplateWorkbook } from './workbook';

/** แถวสุดท้ายที่สูตรรวมทั้งหมดครอบ (แอปรองรับสูงสุด 10,000 แถว; ยังไม่รองรับ G:G ทั้งคอลัมน์) */
export const INVENTORY_TOTAL_LAST_ROW = 10000;

const HEADERS = ['ลำดับ', 'ชื่อ', 'น้ำหนัก', 'เปอร์เซ็น', 'น้ำหนักสุทธิ', 'ราคา', 'รวม'];

/** ตัวอย่าง: [ลำดับ, ชื่อ, น้ำหนัก, เปอร์เซ็น, ราคา] */
const SAMPLES: [string, string, string, string, string][] = [
  ['1', 'ตัวอย่าง 1', '100', '38', '40'],
  ['2', 'ตัวอย่าง 2', '250', '21', '40'],
  ['3', 'ตัวอย่าง 3', '80', '66', '41'],
];

/**
 * Creates the "รายการสินค้า" (Product Inventory) template workbook (SPEC 7 & Phase 3D)
 * คอลัมน์: ลำดับ | ชื่อ | น้ำหนัก | เปอร์เซ็น | น้ำหนักสุทธิ (=C*D%) | ราคา | รวม (=E*F); I1:J1 รวมทั้งหมด
 */
export function createInventoryWorkbook(): Workbook {
  const sheet: Sheet = createEmptySheet('รายการสินค้า', 100, 26);

  // ความกว้าง (7 px ต่อหน่วย .xlsx): A ≈ 6.6, B ≈ 17.9, H ≈ 4.4 หน่วย; คอลัมน์อื่นใช้ค่าเริ่มต้น
  sheet.colWidths = {
    0: 46, // ลำดับ
    1: 125, // ชื่อ
    7: 31, // ช่องว่างคั่น
  };

  // หัวตาราง: สีพื้นเขียว ตัวอักษรไม่หนา รูปแบบทั่วไป
  const headerStyle: CellStyle = { fill: '#93c47d' };
  HEADERS.forEach((text, i) => {
    sheet.cells[`${'ABCDEFG'[i]}1`] = { raw: text, style: { ...headerStyle } };
  });

  SAMPLES.forEach(([no, name, weight, percent, price], i) => {
    const r = i + 2;
    sheet.cells[`A${r}`] = { raw: no };
    sheet.cells[`B${r}`] = { raw: name };
    sheet.cells[`C${r}`] = { raw: weight };
    sheet.cells[`D${r}`] = { raw: percent };
    sheet.cells[`E${r}`] = { raw: `=C${r}*D${r}%` };
    sheet.cells[`F${r}`] = { raw: price };
    sheet.cells[`G${r}`] = { raw: `=E${r}*F${r}` };
  });

  // สรุปด้านขวา (H1 เว้นว่าง)
  sheet.cells['I1'] = { raw: 'รวมทั้งหมด', style: { fill: '#6d9eeb' } };
  sheet.cells['J1'] = { raw: `=SUM(G2:G${INVENTORY_TOTAL_LAST_ROW})`, style: { fill: '#ffffff' } };

  return buildTemplateWorkbook([sheet]);
}
