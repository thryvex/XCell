import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';

/**
 * สี UI ทั้งหมดต้องมาจาก src/ui/theme.css (SPEC 5.17)
 * ยกเว้น: theme.css เอง, หน้าที่พิมพ์ (src/core/printRender.ts อยู่นอก src/ui),
 * และจานสีให้ผู้ใช้เลือกสีเซล (ทำเครื่องหมาย theme-exempt / theme-exempt-start … theme-exempt-end)
 */
const ROOT = path.resolve(__dirname, '../../src');
const COLOR = /#[0-9a-fA-F]{3,8}\b|\brgba?\(|\bhsla?\(/;

function listFiles(dir: string): string[] {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) return listFiles(p);
    return /\.(tsx?|css)$/.test(e.name) ? [p] : [];
  });
}

function findHardcodedColors(file: string, text: string): string[] {
  const hits: string[] = [];
  let exempt = false;
  text.split(/\r?\n/).forEach((line, i) => {
    if (line.includes('theme-exempt-start')) exempt = true;
    const skip = exempt || line.includes('theme-exempt');
    if (line.includes('theme-exempt-end')) exempt = false;
    if (skip) return;
    // ไม่นับ entity ของ HTML (&#123;) และ anchor ใน URL
    const cleaned = line.replace(/&#\d+;/g, '');
    if (COLOR.test(cleaned)) hits.push(`${path.relative(ROOT, file)}:${i + 1}: ${line.trim()}`);
  });
  return hits;
}

describe('ห้ามมีค่าสี hardcode ใน UI', () => {
  it('detector catches hex/rgb and honours exemptions', () => {
    expect(findHardcodedColors('x.tsx', "color: '#fff'")).toHaveLength(1);
    expect(findHardcodedColors('x.tsx', "boxShadow: '0 1px 2px rgba(0,0,0,.2)'")).toHaveLength(1);
    expect(findHardcodedColors('x.tsx', "color: 'var(--fg-main)'")).toHaveLength(0);
    expect(findHardcodedColors('x.tsx', "a = '#000000' // theme-exempt: palette")).toHaveLength(0);
    const hits = findHardcodedColors(
      'x.tsx',
      "// theme-exempt-start\nconst P = ['#111111'];\n// theme-exempt-end\nconst c = '#222222';"
    );
    expect(hits).toHaveLength(1);
    expect(hits[0]).toMatch(/:4: const c = '#222222';$/);
  });

  it('src/ui and src/index.css use theme tokens only', () => {
    const files = [...listFiles(path.join(ROOT, 'ui')), path.join(ROOT, 'index.css')].filter(
      (f) => path.basename(f) !== 'theme.css'
    );
    expect(files.length).toBeGreaterThan(10);
    const hits = files.flatMap((f) => findHardcodedColors(f, fs.readFileSync(f, 'utf8')));
    expect(hits, 'พบค่าสี hardcode (ย้ายไปเป็น token ใน src/ui/theme.css):\n' + hits.join('\n')).toEqual([]);
  });
});
