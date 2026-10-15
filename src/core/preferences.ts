import type { YearMode } from './numfmt';
import type { ThemePref } from './theme';

export interface AppPreferences {
  showWelcomeGuide: boolean;
  yearMode?: YearMode; // รูปแบบปีที่แสดง (5A): ค่าของแอป ไม่ใช่ค่าของไฟล์
  hideBuddhistSaveNotice?: boolean; // ไม่ต้องแสดงข้อความ "ไฟล์เก็บปีเป็น ค.ศ." อีก
  theme?: ThemePref; // ธีมสี (5C): ค่าของแอป ไม่เขียนลง .xlsx หรือไฟล์กู้คืน
  gridWhite?: boolean; // พื้นตารางขาวเสมอ แม้อยู่ในโหมดมืด
}

export const DEFAULT_PREFERENCES: AppPreferences = {
  showWelcomeGuide: true,
  yearMode: 'ce',
  hideBuddhistSaveNotice: false,
  theme: 'system',
  gridWhite: false,
};

export function parsePreferences(jsonStr: string | null | undefined): AppPreferences {
  if (!jsonStr) return { ...DEFAULT_PREFERENCES };
  try {
    const data = JSON.parse(jsonStr);
    return {
      showWelcomeGuide: typeof data.showWelcomeGuide === 'boolean' ? data.showWelcomeGuide : true,
      yearMode: data.yearMode === 'be' ? 'be' : 'ce',
      hideBuddhistSaveNotice: data.hideBuddhistSaveNotice === true,
      theme: data.theme === 'light' || data.theme === 'dark' ? data.theme : 'system',
      gridWhite: data.gridWhite === true,
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
