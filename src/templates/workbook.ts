import { Workbook, Sheet } from '../core/model';

/**
 * สร้างสมุดงานเทมเพลตจากรายการชีต (รองรับหลายชีต เช่น ชีต "สรุป" ใน 4D)
 * - ชื่อชีตต้องไม่ซ้ำกัน (ไม่สนตัวพิมพ์เล็ก/ใหญ่)
 * - activeSheet คือชีตที่เปิดแสดงเมื่อโหลดเทมเพลต
 */
export function buildTemplateWorkbook(sheets: Sheet[], activeSheet = 0): Workbook {
  if (sheets.length === 0) {
    throw new Error('template must define at least one sheet');
  }
  const seen = new Set<string>();
  for (const sheet of sheets) {
    const key = sheet.name.toLowerCase();
    if (seen.has(key)) {
      throw new Error(`duplicate template sheet name: ${sheet.name}`);
    }
    seen.add(key);
  }
  return {
    sheets,
    activeSheet: Math.max(0, Math.min(sheets.length - 1, activeSheet)),
  };
}
