import type { Workbook } from './model';

export type ThemePref = 'system' | 'light' | 'dark';
export type ResolvedTheme = 'light' | 'dark';

/** ตัดสินธีมที่ใช้จริง: ตามระบบใช้โหมดสีของระบบปฏิบัติการ */
export function resolveTheme(pref: ThemePref, systemPrefersDark: boolean): ResolvedTheme {
  if (pref === 'light' || pref === 'dark') return pref;
  return systemPrefersDark ? 'dark' : 'light';
}

// ---------------------------------------------------------------------------
// Contrast (WCAG 2.x)
// ---------------------------------------------------------------------------

/** #RGB หรือ #RRGGBB → [r, g, b] (0–255) หรือ null */
export function parseHexColor(hex: string): [number, number, number] | null {
  const m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(hex.trim());
  if (!m) return null;
  let h = m[1];
  if (h.length === 3) h = h.split('').map((ch) => ch + ch).join('');
  return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
}

/** relative luminance ตาม WCAG */
export function relativeLuminance(rgb: [number, number, number]): number {
  const [r, g, b] = rgb.map((v) => {
    const c = v / 255;
    return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/** contrast ratio ระหว่างสองสี (1–21) หรือ NaN ถ้าสีไม่ถูกรูปแบบ */
export function contrastRatio(a: string, b: string): number {
  const ca = parseHexColor(a);
  const cb = parseHexColor(b);
  if (!ca || !cb) return NaN;
  const la = relativeLuminance(ca);
  const lb = relativeLuminance(cb);
  const [hi, lo] = la > lb ? [la, lb] : [lb, la];
  return (hi + 0.05) / (lo + 0.05);
}

// ---------------------------------------------------------------------------
// เซลที่สีตัวอักษรอาจมองยาก (เปิดไฟล์ในโหมดมืด)
// ---------------------------------------------------------------------------

export const LOW_CONTRAST_THRESHOLD = 3;

export interface LowContrastCell {
  sheetIndex: number;
  key: string;
  ratio: number;
}

/**
 * หาเซลที่ตั้งสีตัวอักษรไว้ชัดเจนและ contrast ต่ำกว่า 3:1 กับพื้นที่แสดงจริง
 * (สีพื้นเซลถ้ามี มิฉะนั้นพื้นตาราง) — ไม่นับเซลว่างและเซลที่ไม่ได้ตั้งสีตัวอักษร
 */
export function findLowContrastCells(
  wb: Workbook,
  gridBackground: string,
  threshold = LOW_CONTRAST_THRESHOLD
): LowContrastCell[] {
  const result: LowContrastCell[] = [];
  wb.sheets.forEach((sheet, sheetIndex) => {
    for (const [key, cell] of Object.entries(sheet.cells)) {
      const color = cell.style?.color;
      if (!color || cell.raw === '') continue;
      const bg = cell.style?.fill || gridBackground;
      const ratio = contrastRatio(color, bg);
      if (!isNaN(ratio) && ratio < threshold) result.push({ sheetIndex, key, ratio });
    }
  });
  return result;
}

// ---------------------------------------------------------------------------
// อ่าน design token จาก theme.css (ใช้ในเทสต์และตอนต้องการค่าสีจริง)
// ---------------------------------------------------------------------------

/** ค่าตัวแปร CSS ในบล็อกแรกที่ selector ตรง (เช่น '[data-theme="dark"]') */
export function parseTokenBlock(css: string, selector: string): Record<string, string> {
  const start = css.indexOf(selector);
  if (start < 0) return {};
  const open = css.indexOf('{', start);
  const close = css.indexOf('}', open);
  const body = css.slice(open + 1, close);
  const tokens: Record<string, string> = {};
  for (const m of body.matchAll(/(--[a-z0-9-]+)\s*:\s*([^;]+);/gi)) {
    tokens[m[1]] = m[2].trim();
  }
  return tokens;
}

/** แทนค่า var(--x) ซ้อนกันจนได้ค่าจริง */
export function resolveToken(tokens: Record<string, string>, name: string, depth = 0): string | undefined {
  const value = tokens[name];
  if (value === undefined || depth > 10) return value;
  const m = /^var\((--[a-z0-9-]+)\)$/i.exec(value);
  return m ? resolveToken(tokens, m[1], depth + 1) : value;
}

/**
 * สีพื้นหน้าต่าง Electron ตอนสร้าง (ก่อนหน้าเว็บโหลด) — ต้องตรงกับ --bg-app ใน theme.css
 * (มีเทสต์ตรวจ) เพื่อไม่ให้หน้าต่างขาววาบตอนเปิดในโหมดมืด
 */
export const WINDOW_BACKGROUND: Record<ResolvedTheme, string> = {
  light: '#ffffff',
  dark: '#121417',
};

/** อาร์กิวเมนต์ที่ main ส่งธีมเริ่มต้นให้หน้าเว็บ (อ่านใน preload ก่อน React เริ่ม) */
export const INITIAL_THEME_ARG = '--xcell-theme=';

/** ต้องเตือน "สีตัวอักษรอาจมองยากในโหมดมืด" หรือไม่ (เฉพาะโหมดมืดที่พื้นตารางไม่ได้ตั้งเป็นขาว) */
export function needsLowContrastWarning(
  wb: Workbook,
  theme: ResolvedTheme,
  gridWhite: boolean,
  gridBackground: string
): boolean {
  if (theme !== 'dark' || gridWhite) return false;
  return findLowContrastCells(wb, gridBackground).length > 0;
}

/** สีตัวอักษรหลักของแต่ละธีม (ต้องตรงกับ --grid-fg ใน theme.css มีเทสต์ตรวจ) */
export const THEME_TEXT: Record<ResolvedTheme, string> = {
  light: '#1f2328',
  dark: '#e2e8f0',
};

/**
 * สีตัวอักษรที่ใช้แสดงเมื่อเซลตั้งสีพื้นไว้แต่ไม่ได้ตั้งสีตัวอักษร:
 * เลือกสีตัวอักษรของธีมที่อ่านบนสีพื้นนั้นได้ชัดกว่า (เช่น หัวตารางพื้นเทาอ่อนในโหมดมืดยังเป็นตัวเข้ม)
 * ใช้แสดงผลเท่านั้น ไม่เขียนลงไฟล์ — คืน undefined ถ้าอ่านสีพื้นไม่ได้ (ให้ใช้สีของธีม)
 */
export function defaultTextColorOnFill(fill: string): string | undefined {
  const dark = contrastRatio(THEME_TEXT.light, fill);
  const light = contrastRatio(THEME_TEXT.dark, fill);
  if (isNaN(dark) || isNaN(light)) return undefined;
  return dark >= light ? THEME_TEXT.light : THEME_TEXT.dark;
}
