import { th } from '../i18n/th';

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

/**
 * Formats a sheet name for formula representation according to Excel/Google Sheets quoting rules:
 * - Simple ASCII letters/numbers/underscore don't require quotes
 * - Otherwise wrap in single quotes '...' and escape existing single quotes as ''
 */
export function formatSheetReference(sheetName: string): string {
  // If simple ASCII alphanumeric or underscore, doesn't require quotes
  if (/^[A-Za-z0-9_]+$/.test(sheetName)) {
    return sheetName;
  }
  // Otherwise wrap in single quotes and escape ' with ''
  const escaped = sheetName.replace(/'/g, "''");
  return `'${escaped}'`;
}

/**
 * Checks whether a raw formula string references the specified sheet.
 */
export function formulaReferencesSheet(rawFormula: string, targetSheetName: string): boolean {
  if (!rawFormula || !rawFormula.startsWith('=')) return false;

  // 1. Quoted reference: 'SheetName'! (with '' escaped)
  const escapedQuotes = targetSheetName.replace(/'/g, "''");
  const regexQuotedTarget = escapedQuotes.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
  const quotedRegex = new RegExp(`'${regexQuotedTarget}'\\s*!`, 'i');
  if (quotedRegex.test(rawFormula)) return true;

  // 2. Unquoted reference (only valid if sheetName has no special characters or spaces)
  if (!/[\s'!"#$%&()*+,/:;<=>?@[\\\]^`{|}~-]/.test(targetSheetName)) {
    const regexTarget = targetSheetName.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
    const unquotedRegex = new RegExp(`(?<![A-Za-z0-9_ก-๙])${regexTarget}\\s*!`, 'i');
    if (unquotedRegex.test(rawFormula)) return true;
  }

  return false;
}

/**
 * Counts how many formula cells across other sheets in the workbook reference targetSheetName.
 */
export function countFormulasReferencingSheet(
  sheets: Array<{ name: string; cells: Record<string, { raw?: string }> }>,
  targetSheetName: string
): number {
  let count = 0;
  const targetLower = targetSheetName.toLowerCase();

  for (const sheet of sheets) {
    if (sheet.name.toLowerCase() === targetLower) continue;

    for (const cell of Object.values(sheet.cells)) {
      if (cell.raw && formulaReferencesSheet(cell.raw, targetSheetName)) {
        count++;
      }
    }
  }

  return count;
}
