import type { YearMode } from './numfmt';

export interface AppPreferences {
  showWelcomeGuide: boolean;
  yearMode?: YearMode; // รูปแบบปีที่แสดง (5A): ค่าของแอป ไม่ใช่ค่าของไฟล์
  hideBuddhistSaveNotice?: boolean; // ไม่ต้องแสดงข้อความ "ไฟล์เก็บปีเป็น ค.ศ." อีก
}

export const DEFAULT_PREFERENCES: AppPreferences = {
  showWelcomeGuide: true,
  yearMode: 'ce',
  hideBuddhistSaveNotice: false,
};

export function parsePreferences(jsonStr: string | null | undefined): AppPreferences {
  if (!jsonStr) return { ...DEFAULT_PREFERENCES };
  try {
    const data = JSON.parse(jsonStr);
    return {
      showWelcomeGuide: typeof data.showWelcomeGuide === 'boolean' ? data.showWelcomeGuide : true,
      yearMode: data.yearMode === 'be' ? 'be' : 'ce',
      hideBuddhistSaveNotice: data.hideBuddhistSaveNotice === true,
    };
  } catch {
    return { ...DEFAULT_PREFERENCES };
  }
}

export function serializePreferences(prefs: AppPreferences): string {
  return JSON.stringify(prefs, null, 2);
}

/** รวมค่าที่เปลี่ยนเข้ากับค่าเดิม (บันทึกทีละค่าโดยไม่ทับค่าอื่น) */
export function mergePreferences(current: AppPreferences, patch: Partial<AppPreferences>): AppPreferences {
  return parsePreferences(JSON.stringify({ ...current, ...patch }));
}
