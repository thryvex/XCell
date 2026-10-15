import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { contrastRatio, parseTokenBlock, resolveToken, parseHexColor } from '../../src/core/theme';

const css = fs.readFileSync(path.resolve(__dirname, '../../src/ui/theme.css'), 'utf8');
const light = parseTokenBlock(css, ':root[data-theme="light"]');
const darkOverrides = parseTokenBlock(css, ':root[data-theme="dark"]');
const whiteGrid = parseTokenBlock(css, '[data-grid-white="true"]');

/** ธีมมืดสืบทอดค่าที่ไม่ได้ประกาศจากชุดสว่าง (เหมือน CSS cascade) */
const THEMES: Record<string, Record<string, string>> = {
  light,
  dark: { ...light, ...darkOverrides },
  'dark + พื้นตารางขาว': { ...light, ...darkOverrides, ...whiteGrid },
};

// คู่ "ตัวอักษร / พื้นหลัง" ที่ใช้จริง ต้อง >= 4.5:1
const TEXT_PAIRS: Array<[string, string]> = [
  ['--fg-main', '--bg-app'],
  ['--fg-main', '--bg-surface'],
  ['--fg-main', '--bg-header'],
  ['--fg-main', '--bg-hover'],
  ['--fg-main', '--dialog-bg'],
  ['--fg-main', '--menu-bg'],
  ['--fg-main', '--menu-hover'],
  ['--fg-main', '--bg-selection'],
  ['--fg-muted', '--bg-app'],
  ['--fg-muted', '--bg-surface'],
  ['--fg-muted', '--bg-header'],
  ['--fg-muted', '--dialog-bg'],
  ['--fg-muted', '--menu-bg'],
  ['--color-primary', '--bg-app'],
  ['--color-primary', '--bg-surface'],
  ['--color-primary', '--dialog-bg'],
  ['--fg-on-primary', '--color-primary'],
  ['--fg-on-primary', '--color-primary-hover'],
  ['--color-danger', '--bg-app'],
  ['--color-danger', '--dialog-bg'],
  ['--color-danger', '--menu-bg'],
  ['--color-danger', '--danger-bg'],
  ['--color-warning', '--bg-app'],
  ['--color-warning', '--dialog-bg'],
  ['--color-success', '--bg-surface'],
  ['--banner-warn-fg', '--banner-warn-bg'],
  ['--banner-info-fg', '--banner-info-bg'],
  ['--tab-fg', '--tab-bar-bg'],
  ['--tab-active-fg', '--tab-active-bg'],
  ['--status-fg', '--status-bg'],
  ['--color-danger', '--status-bg'],
  ['--color-success', '--status-bg'],
  ['--color-primary', '--status-bg'],
  ['--tooltip-fg', '--tooltip-bg'],
  ['--toast-fg', '--toast-bg'],
  ['--grid-fg', '--grid-bg'],
  ['--grid-fg', '--grid-selection'],
  ['--grid-fg', '--grid-active-bg'],
  ['--grid-header-fg', '--grid-header-bg'],
  ['--grid-header-fg', '--grid-selection'],
  ['--grid-error', '--grid-bg'],
  ['--grid-error', '--grid-active-bg'],
  ['--search-fg', '--search-match-bg'],
  ['--search-fg', '--search-current-bg'],
  ['--banner-warn-fg', '--bg-app'], // ปุ่มในแถบเตือน
  ['--banner-info-fg', '--bg-surface'],
  ['--fg-main', '--tab-active-bg'],
  ['--fg-main', '--status-bg'],
  ['--grid-fg', '--grid-header-bg'], // ปุ่มเพิ่มแถวท้ายตาราง
  ['--fg-on-primary', '--toast-bg'],
];

// เส้นขอบ/ส่วนประกอบที่ต้องมองเห็น ต้อง >= 3:1
const UI_PAIRS: Array<[string, string]> = [
  ['--control-border', '--bg-app'],
  ['--control-border', '--bg-surface'],
  ['--control-border', '--dialog-bg'],
  ['--color-primary', '--bg-header'], // ปุ่ม/แท็บที่ถูกเลือก, เส้นโฟกัส
  ['--grid-active-border', '--grid-bg'], // กรอบเซลที่ใช้งาน
  ['--grid-active-border', '--grid-selection'],
  ['--fill-handle', '--grid-bg'],
  ['--grid-unsupported', '--grid-bg'],
  ['--ref-1', '--grid-bg'],
  ['--ref-2', '--grid-bg'],
  ['--ref-3', '--grid-bg'],
  ['--ref-4', '--grid-bg'],
  ['--ref-5', '--grid-bg'],
  ['--ref-6', '--grid-bg'],
  ['--banner-warn-border', '--banner-warn-bg'],
  ['--banner-info-border', '--banner-info-bg'],
  ['--scrollbar-thumb-hover', '--scrollbar-track'],
  ['--control-border', '--grid-bg'], // ปุ่มตัวกรองในหัวตาราง
  ['--control-border', '--menu-bg'],
  ['--grid-active-border', '--grid-header-bg'], // เส้นตรึงแถว/คอลัมน์
  ['--tab-active-fg', '--tab-bar-bg'], // เส้นใต้แท็บที่เลือก
  ['--color-primary', '--tab-bar-bg'], // เส้นบอกตำแหน่งตอนลากแท็บ
];

function check(pairs: Array<[string, string]>, min: number) {
  const failures: string[] = [];
  for (const [name, tokens] of Object.entries(THEMES)) {
    for (const [fg, bg] of pairs) {
      const a = resolveToken(tokens, fg);
      const b = resolveToken(tokens, bg);
      const ratio = contrastRatio(a ?? '', b ?? '');
      if (!(ratio >= min)) {
        failures.push(`[${name}] ${fg} (${a}) บน ${bg} (${b}) = ${isNaN(ratio) ? 'อ่านค่าไม่ได้' : ratio.toFixed(2)}:1`);
      }
    }
  }
  return failures;
}

describe('ความอ่านง่ายของธีม (theme.css)', () => {
  it('token blocks are present', () => {
    expect(Object.keys(light).length).toBeGreaterThan(50);
    expect(Object.keys(darkOverrides).length).toBeGreaterThan(50);
    expect(Object.keys(whiteGrid).length).toBeGreaterThan(10);
  });

  it('text pairs meet 4.5:1 in every theme', () => {
    expect(check(TEXT_PAIRS, 4.5)).toEqual([]);
  });

  it('borders and components that must be visible meet 3:1 in every theme', () => {
    expect(check(UI_PAIRS, 3)).toEqual([]);
  });

  // ระยะห่าง RGB > 75: สีเข้มที่อ่านได้บนพื้นขาวมีช่วงสีจำกัด ค่านี้ยังแยกสีออกชัด (เช่น เขียว #15803d กับฟ้าอมเขียว #0e7490)
  it('the six formula reference colors are clearly distinct from each other', () => {
    for (const [name, tokens] of Object.entries(THEMES)) {
      const colors = [1, 2, 3, 4, 5, 6].map((i) => parseHexColor(resolveToken(tokens, `--ref-${i}`)!)!);
      for (let i = 0; i < colors.length; i++) {
        for (let j = i + 1; j < colors.length; j++) {
          const d = Math.hypot(...colors[i].map((v, k) => v - colors[j][k]));
          expect(d, `[${name}] --ref-${i + 1} กับ --ref-${j + 1} ใกล้กันเกินไป (${d.toFixed(0)})`).toBeGreaterThan(75);
        }
      }
    }
  });

  it('print preview paper is white in both themes', () => {
    expect(resolveToken(THEMES.light, '--paper-bg')).toBe('#ffffff');
    expect(resolveToken(THEMES.dark, '--paper-bg')).toBe('#ffffff');
  });
});
