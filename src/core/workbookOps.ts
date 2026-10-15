import { Workbook, Sheet, createEmptySheet, newSheetId } from './model';
import {
  validateSheetName,
  generateNextSheetName,
  generateDuplicateSheetName,
} from './sheetName';
import { th } from '../i18n/th';

export interface SheetOpResult {
  ok: boolean;
  error?: string;
  /** index ของชีตที่ควรแสดงหลังทำรายการ */
  index?: number;
}

export function isSheetVisible(sheet: Sheet): boolean {
  return !sheet.hidden;
}

/** index ของชีตที่แสดงแท็บได้ (ไม่รวมชีตซ่อน) เรียงตามลำดับในสมุดงาน */
export function visibleSheetIndices(wb: Workbook): number[] {
  const result: number[] = [];
  wb.sheets.forEach((s, i) => {
    if (isSheetVisible(s)) result.push(i);
  });
  return result;
}

/** ชีตที่มองเห็นถัดไป/ก่อนหน้า (ไม่วนรอบ) หรือ null ถ้าไม่มี */
export function adjacentVisibleSheet(wb: Workbook, from: number, direction: 1 | -1): number | null {
  for (let i = from + direction; i >= 0 && i < wb.sheets.length; i += direction) {
    if (isSheetVisible(wb.sheets[i])) return i;
  }
  return null;
}

/** ชีตแรกที่มองเห็นได้ ใกล้ index ที่ให้มาที่สุด (ใช้ตอนเปิดไฟล์ที่ชีต active ถูกซ่อน) */
export function nearestVisibleSheet(wb: Workbook, index: number): number {
  const clamped = Math.max(0, Math.min(wb.sheets.length - 1, index));
  if (wb.sheets[clamped] && isSheetVisible(wb.sheets[clamped])) return clamped;
  return (
    adjacentVisibleSheet(wb, clamped, 1) ??
    adjacentVisibleSheet(wb, clamped, -1) ??
    clamped
  );
}

function sheetNames(wb: Workbook): string[] {
  return wb.sheets.map((s) => s.name);
}

/** เพิ่มชีตใหม่ต่อท้าย ชื่อ "ชีตN" ที่ไม่ซ้ำ */
export function addSheet(wb: Workbook): SheetOpResult {
  const name = generateNextSheetName(sheetNames(wb));
  wb.sheets.push(createEmptySheet(name));
  return { ok: true, index: wb.sheets.length - 1 };
}

/** ทำสำเนาชีตไว้ถัดจากต้นฉบับ ชื่อ "สำเนาของ X" ที่ไม่ซ้ำ; สูตรคัดลอกตามเดิมทุกตัวอักษร */
export function duplicateSheet(wb: Workbook, index: number): SheetOpResult {
  const source = wb.sheets[index];
  if (!source) return { ok: false };

  const copy: Sheet = JSON.parse(JSON.stringify(source));
  copy.id = newSheetId();
  copy.name = generateDuplicateSheetName(source.name, sheetNames(wb));
  delete copy.hidden;

  wb.sheets.splice(index + 1, 0, copy);
  return { ok: true, index: index + 1 };
}

export function renameSheet(wb: Workbook, index: number, newName: string): SheetOpResult {
  const sheet = wb.sheets[index];
  if (!sheet) return { ok: false };

  const trimmed = newName.trim();
  const validation = validateSheetName(trimmed, sheetNames(wb), sheet.name);
  if (!validation.valid) {
    return { ok: false, error: validation.error };
  }

  sheet.name = trimmed;
  return { ok: true, index };
}

/** ลบชีต: ห้ามลบชีตสุดท้ายที่มองเห็นได้ */
export function deleteSheet(wb: Workbook, index: number): SheetOpResult {
  const sheet = wb.sheets[index];
  if (!sheet) return { ok: false };

  if (isSheetVisible(sheet) && visibleSheetIndices(wb).length <= 1) {
    return { ok: false, error: th.sheetTabs.cannotDeleteLastSheet };
  }

  wb.sheets.splice(index, 1);
  const next = nearestVisibleSheet(wb, Math.min(index, wb.sheets.length - 1));
  return { ok: true, index: next };
}

/** ย้ายชีตจากตำแหน่ง from ไปตำแหน่ง to (index หลังย้าย) */
export function moveSheet(wb: Workbook, from: number, to: number): SheetOpResult {
  if (from === to) return { ok: false };
  if (from < 0 || from >= wb.sheets.length || to < 0 || to >= wb.sheets.length) {
    return { ok: false };
  }
  const [sheet] = wb.sheets.splice(from, 1);
  wb.sheets.splice(to, 0, sheet);
  return { ok: true, index: to };
}
