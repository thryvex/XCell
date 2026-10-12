import { Workbook } from './model';

export interface RecoverySessionData {
  id: string;
  originalPath: string | null;
  fileName: string;
  savedAt: number;
  workbook: Workbook;
}

export function serializeRecoveryData(data: RecoverySessionData): string {
  return JSON.stringify(data);
}

export function deserializeRecoveryData(raw: string): RecoverySessionData | null {
  try {
    const parsed = JSON.parse(raw);
    if (parsed && parsed.workbook && Array.isArray(parsed.workbook.sheets)) {
      return parsed;
    }
  } catch {}
  return null;
}
