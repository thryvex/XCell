/**
 * แปลงความกว้างคอลัมน์/ความสูงแถว px ของแอป <-> หน่วยใน .xlsx
 * (ความกว้าง ~8 px ต่อหน่วย ปัดเป็นจำนวนเต็ม ขั้นต่ำ 8 หน่วย, ความสูงเป็น pt ปัดเป็นจำนวนเต็ม)
 */
export function pxToXlsxColWidth(px: number): number {
  return Math.max(8, Math.round(px / 8));
}

export function xlsxColWidthToPx(width: number): number {
  return Math.round(width * 8);
}

export function pxToXlsxRowHeight(px: number): number {
  return Math.round(px * 0.75);
}

export function xlsxRowHeightToPx(pt: number): number {
  return Math.round(pt / 0.75);
}

/** ขนาด px ที่จะได้กลับมาเมื่อเปิดไฟล์ที่บันทึก (ใช้คำนวณตำแหน่งกราฟในไฟล์ให้ตรงกับคอลัมน์/แถวในไฟล์) */
export function storedColWidthPx(px: number): number {
  return xlsxColWidthToPx(pxToXlsxColWidth(px));
}

export function storedRowHeightPx(px: number): number {
  return xlsxRowHeightToPx(pxToXlsxRowHeight(px));
}
