import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import {
  contrastRatio,
  parseHexColor,
  resolveTheme,
  findLowContrastCells,
  needsLowContrastWarning,
  parseTokenBlock,
  resolveToken,
  WINDOW_BACKGROUND,
  INITIAL_THEME_ARG,
  THEME_TEXT,
  defaultTextColorOnFill,
} from '../../src/core/theme';
import { createEmptyWorkbook } from '../../src/core/model';
import { parsePreferences, mergePreferences, DEFAULT_PREFERENCES } from '../../src/core/preferences';

const DARK_GRID = '#121417';

describe('contrastRatio (WCAG)', () => {
  it('black on white is 21:1 and identical colors are 1:1', () => {
    expect(contrastRatio('#000000', '#ffffff')).toBeCloseTo(21, 5);
    expect(contrastRatio('#ffffff', '#000000')).toBeCloseTo(21, 5);
    expect(contrastRatio('#777777', '#777777')).toBe(1);
    expect(contrastRatio('#2563eb', '#2563eb')).toBe(1);
  });

  it('matches known reference values', () => {
    // ค่าอ้างอิงจากเครื่องมือ WebAIM contrast checker
    expect(contrastRatio('#767676', '#ffffff')).toBeCloseTo(4.54, 2);
    expect(contrastRatio('#0000ff', '#ffffff')).toBeCloseTo(8.59, 2);
  });

  it('accepts #rgb shorthand and rejects invalid input', () => {
    expect(contrastRatio('#000', '#fff')).toBeCloseTo(21, 5);
    expect(parseHexColor('#abc')).toEqual([170, 187, 204]);
    expect(contrastRatio('red', '#fff')).toBeNaN();
    expect(contrastRatio('#12345', '#fff')).toBeNaN();
  });
});

describe('resolveTheme', () => {
  it('explicit light/dark ignore the OS mode', () => {
    expect(resolveTheme('light', true)).toBe('light');
    expect(resolveTheme('dark', false)).toBe('dark');
  });

  it('system follows the OS mode, including when it changes', () => {
    let osDark = false;
    expect(resolveTheme('system', osDark)).toBe('light');
    osDark = true; // ผู้ใช้เปลี่ยนโหมดของเครื่องระหว่างใช้งาน
    expect(resolveTheme('system', osDark)).toBe('dark');
  });
});

describe('findLowContrastCells', () => {
  it('flags explicit text colors that are hard to read on the dark grid', () => {
    const wb = createEmptyWorkbook();
    const cells = wb.sheets[0].cells;
    cells['A1'] = { raw: 'ดำ', style: { color: '#000000' } }; // ดำบนพื้นมืด ~1.1:1
    cells['A2'] = { raw: 'ขาว', style: { color: '#ffffff' } }; // อ่านง่าย
    cells['A3'] = { raw: 'ไม่ตั้งสี' }; // ใช้สีธีม ไม่นับ
    cells['A4'] = { raw: '', style: { color: '#000000' } }; // เซลว่าง ไม่นับ
    const found = findLowContrastCells(wb, DARK_GRID);
    expect(found.map((f) => f.key)).toEqual(['A1']);
    expect(found[0].ratio).toBeLessThan(3);
  });

  it('compares against the cell fill when the cell has one', () => {
    const wb = createEmptyWorkbook();
    const cells = wb.sheets[0].cells;
    cells['A1'] = { raw: 'ดำบนเหลือง', style: { color: '#000000', fill: '#fef08a' } }; // อ่านง่าย
    cells['A2'] = { raw: 'เหลืองบนขาว', style: { color: '#fde047', fill: '#ffffff' } }; // อ่านยาก
    const found = findLowContrastCells(wb, DARK_GRID);
    expect(found.map((f) => f.key)).toEqual(['A2']);
  });

  it('checks every sheet', () => {
    const wb = createEmptyWorkbook();
    wb.sheets.push({ ...createEmptyWorkbook().sheets[0], name: 'ชีต2', cells: { B2: { raw: 'x', style: { color: '#1f2328' } } } });
    const found = findLowContrastCells(wb, DARK_GRID);
    expect(found).toEqual([expect.objectContaining({ sheetIndex: 1, key: 'B2' })]);
  });

  it('warns only in dark mode without the white grid', () => {
    const wb = createEmptyWorkbook();
    wb.sheets[0].cells['A1'] = { raw: 'ดำ', style: { color: '#000000' } };
    expect(needsLowContrastWarning(wb, 'dark', false, DARK_GRID)).toBe(true);
    expect(needsLowContrastWarning(wb, 'dark', true, DARK_GRID)).toBe(false);
    expect(needsLowContrastWarning(wb, 'light', false, '#ffffff')).toBe(false);
    expect(needsLowContrastWarning(createEmptyWorkbook(), 'dark', false, DARK_GRID)).toBe(false);
  });
});

describe('token parser and window background', () => {
  const css = fs.readFileSync(path.resolve(__dirname, '../../src/ui/theme.css'), 'utf8');

  it('parses a token block and resolves var() chains', () => {
    const tokens = parseTokenBlock(':root { --a: #111111; --b: var(--a); --c: var(--b); }', ':root');
    expect(resolveToken(tokens, '--c')).toBe('#111111');
    expect(resolveToken(tokens, '--missing')).toBeUndefined();
  });

  it('window background at creation matches --bg-app of each theme (no white flash)', () => {
    const light = parseTokenBlock(css, ':root[data-theme="light"]');
    const dark = parseTokenBlock(css, ':root[data-theme="dark"]');
    expect(light['--bg-app']).toBe(WINDOW_BACKGROUND.light);
    expect(dark['--bg-app']).toBe(WINDOW_BACKGROUND.dark);
  });

  it('preload reads the same initial-theme argument that main sends', () => {
    const preload = fs.readFileSync(path.resolve(__dirname, '../../electron/preload.ts'), 'utf8');
    expect(preload).toContain(`'${INITIAL_THEME_ARG}'`);
  });
});

describe('theme preferences', () => {
  it('defaults to system theme with the grid following the theme', () => {
    expect(DEFAULT_PREFERENCES.theme).toBe('system');
    expect(DEFAULT_PREFERENCES.gridWhite).toBe(false);
    expect(parsePreferences(null).theme).toBe('system');
  });

  it('accepts only known values', () => {
    expect(parsePreferences('{"theme":"dark","gridWhite":true}')).toMatchObject({ theme: 'dark', gridWhite: true });
    expect(parsePreferences('{"theme":"purple","gridWhite":"yes"}')).toMatchObject({ theme: 'system', gridWhite: false });
  });

  it('saving the theme keeps other preferences', () => {
    const merged = mergePreferences(parsePreferences('{"yearMode":"be","showWelcomeGuide":false}'), { theme: 'dark' });
    expect(merged).toMatchObject({ yearMode: 'be', showWelcomeGuide: false, theme: 'dark' });
  });
});

describe('defaultTextColorOnFill', () => {
  it('uses dark text on light fills and light text on dark fills', () => {
    expect(defaultTextColorOnFill('#e2e8f0')).toBe(THEME_TEXT.light); // หัวตารางเทาอ่อนของเทมเพลต
    expect(defaultTextColorOnFill('#fef08a')).toBe(THEME_TEXT.light);
    expect(defaultTextColorOnFill('#1e3a8a')).toBe(THEME_TEXT.dark);
    expect(defaultTextColorOnFill('not-a-color')).toBeUndefined();
  });

  it('the chosen color is readable (>= 4.5:1) on the palette fills offered to users', () => {
    const palette = ['#000000', '#434343', '#666666', '#999999', '#cccccc', '#ffffff', '#fecaca', '#fed7aa', '#bbf7d0', '#bfdbfe', '#e2e8f0', '#dbeafe'];
    for (const fill of palette) {
      expect(contrastRatio(defaultTextColorOnFill(fill)!, fill), fill).toBeGreaterThanOrEqual(4.5);
    }
  });

  it('theme text constants match --grid-fg in theme.css', () => {
    const css = fs.readFileSync(path.resolve(__dirname, '../../src/ui/theme.css'), 'utf8');
    expect(parseTokenBlock(css, ':root[data-theme="light"]')['--grid-fg']).toBe(THEME_TEXT.light);
    expect(parseTokenBlock(css, ':root[data-theme="dark"]')['--grid-fg']).toBe(THEME_TEXT.dark);
  });
});
