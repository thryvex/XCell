import { th } from '../i18n/th';
import { tokenize } from './tokenizer';

export const FORBIDDEN_SHEET_NAME_CHARS = ['[', ']', ':', '*', '?', '/', '\\'];

export interface SheetNameValidationResult {
  valid: boolean;
  error?: string;
}

/**
 * Validates a sheet name according to SPEC 5.13 and Agreement 6:
 * - Length: 1 to 31 characters
 * - Cannot be empty or only spaces
 * - Cannot contain forbidden chars: [ ] : * ? / \
 * - Cannot duplicate existing sheet names (case-insensitive)
 */
export function validateSheetName(
  name: string,
  existingNames: string[],
  currentName?: string
): SheetNameValidationResult {
  const trimmed = name.trim();

  if (!trimmed) {
    return { valid: false, error: th.sheetTabs.nameEmpty };
  }

  if (trimmed.length > 31) {
    return { valid: false, error: th.sheetTabs.nameTooLong };
  }

  for (const char of FORBIDDEN_SHEET_NAME_CHARS) {
    if (trimmed.includes(char)) {
      return { valid: false, error: th.sheetTabs.nameInvalidChars };
    }
  }

  const lowerTrimmed = trimmed.toLowerCase();
  const lowerCurrent = currentName?.trim().toLowerCase();

  for (const existing of existingNames) {
    const lowerExisting = existing.trim().toLowerCase();
    if (lowerExisting === lowerTrimmed && lowerExisting !== lowerCurrent) {
      return {
        valid: false,
        error: th.sheetTabs.nameDuplicate.replace('{name}', trimmed),
      };
    }
  }

  return { valid: true };
}

/**
 * Generates the next default sheet name ("ชีต1", "ชีต2", ...) ensuring no collision.
 */
export function generateNextSheetName(existingNames: string[]): string {
  const prefix = th.sheetTabs.defaultSheetPrefix || 'ชีต';
  const lowerExisting = new Set(existingNames.map((n) => n.trim().toLowerCase()));

  let index = existingNames.length + 1;
  while (index < 1000) {
    const candidate = `${prefix}${index}`;
    if (!lowerExisting.has(candidate.toLowerCase())) {
      return candidate;
    }
    index++;
  }

  // Fallback
  return `${prefix}${Date.now()}`.slice(0, 31);
}

/**
 * Generates duplicate sheet name ("สำเนาของ X", "สำเนาของ X (2)", ...) ensuring no collision and <= 31 chars.
 */
export function generateDuplicateSheetName(sourceName: string, existingNames: string[]): string {
  const lowerExisting = new Set(existingNames.map((n) => n.trim().toLowerCase()));
  const prefixTemplate = th.sheetTabs.copyOf || 'สำเนาของ {name}';

  let candidate = prefixTemplate.replace('{name}', sourceName);
  if (candidate.length > 31) {
    // Truncate to 31 characters
    candidate = candidate.slice(0, 31);
  }

  if (!lowerExisting.has(candidate.toLowerCase())) {
    return candidate;
  }

  let counter = 2;
  while (counter < 1000) {
    const suffix = ` (${counter})`;
    const maxSourceLen = Math.max(1, 31 - prefixTemplate.replace('{name}', '').length - suffix.length);
    const truncatedSource = sourceName.slice(0, maxSourceLen);
    const numberedCandidate = prefixTemplate.replace('{name}', truncatedSource) + suffix;

    if (!lowerExisting.has(numberedCandidate.toLowerCase())) {
      return numberedCandidate;
    }
    counter++;
  }

  return candidate;
}

/** เทียบชื่อชีตแบบไม่สนตัวพิมพ์เล็ก/ใหญ่ */
export function sameSheetName(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

/**
 * เขียนชื่อชีตสำหรับใช้ในสูตร (ทั้งในแอปและใน .xlsx):
 * - ไม่ใส่เครื่องหมายคำพูดเฉพาะชื่อที่เป็น A-Z / ตัวเลข / ขีดล่างล้วน ที่ไม่ขึ้นต้นด้วยตัวเลข
 *   และไม่หน้าตาเหมือนที่อยู่เซล (เช่น "A1") หรือ TRUE/FALSE
 * - นอกนั้นใส่ '...' และเขียน ' ในชื่อเป็น ''
 */
export function formatSheetReference(sheetName: string): string {
  const isPlain =
    /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) &&
    !/^[A-Za-z]{1,3}\d+$/.test(sheetName) &&
    !/^(TRUE|FALSE)$/i.test(sheetName);
  if (isPlain) {
    return sheetName;
  }
  const escaped = sheetName.replace(/'/g, "''");
  return `'${escaped}'`;
}

/**
 * Checks whether a raw formula string references the specified sheet.
 */
export function formulaReferencesSheet(rawFormula: string, targetSheetName: string): boolean {
  if (!rawFormula || !rawFormula.startsWith('=')) return false;
  return tokenize(rawFormula).some(
    (t) =>
      (t.type === 'CELL_REF' || t.type === 'RANGE') &&
      t.sheet !== undefined &&
      sameSheetName(t.sheet, targetSheetName)
  );
}

/**
 * Counts how many formula cells across other sheets in the workbook reference targetSheetName.
 */
export function countFormulasReferencingSheet(
  sheets: Array<{ name: string; cells: Record<string, { raw?: string }> }>,
  targetSheetName: string
): number {
  let count = 0;

  for (const sheet of sheets) {
    if (sameSheetName(sheet.name, targetSheetName)) continue;

    for (const cell of Object.values(sheet.cells)) {
      if (cell.raw && formulaReferencesSheet(cell.raw, targetSheetName)) {
        count++;
      }
    }
  }

  return count;
}
