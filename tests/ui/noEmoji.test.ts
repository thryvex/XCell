import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';

/**
 * ห้ามมีอีโมจิและสัญลักษณ์ตกแต่งใน src/ (ไอคอนใช้ SVG ใน src/ui/icons.tsx)
 * คงไว้ตามธรรมเนียมโปรแกรมตารางคำนวณ: … (เมนูที่เปิดหน้าต่าง), × (แถว × คอลัมน์),
 * ● (ยังไม่บันทึก ในชื่อหน้าต่าง) และ ƒx (หน้าแถบสูตร) เฉพาะตำแหน่งใน ALLOWED
 * ไม่ตรวจบรรทัดคอมเมนต์
 */
const ROOT = path.resolve(__dirname, '../../src');

const FORBIDDEN = new RegExp(
  [
    '\\p{Extended_Pictographic}', // อีโมจิทั้งหมด (รวม ⚠ ✓ ✕ ✏ ⌨ ⚙ ❓ ✨)
    '\\uFE0F', // ตัวเลือกรูปแบบอีโมจิ
    '[\\u2190-\\u21FF]', // ลูกศร → ➔ ↶ ↷ ⇤ ⇥ ↵ ⇅ ↻
    '[\\u27A1-\\u27BF]', // ลูกศรตกแต่ง ➔
    '[\\u2200-\\u22FF]', // สัญลักษณ์คณิต Σ ⊞ ⊟ − (ยกเว้นที่อยู่ในข้อความทั่วไป)
    '[\\u2300-\\u23FF]', // สัญลักษณ์เทคนิค ⌘ ⌨
    '[\\u2460-\\u24FF]', // ตัวเลขในวงกลม ① ② ③
    '[\\u2500-\\u25FF]', // กล่องและรูปทรง ▲ ▼ ▾ ◀ ▶ ◻ ●
    '[\\u2600-\\u27BF]', // สัญลักษณ์เบ็ดเตล็ด ☵ ✓ ✕
    '[\\uFF01-\\uFF5E]', // ตัวอักษรเต็มความกว้าง ＋
    '[\\u03A3\\u2211]', // Σ
    '[\\u0192]', // ƒ
    '[田回]', // ตัวอักษรจีนที่ใช้เป็นไอคอน
  ].join('|'),
  'u'
);

/** อนุญาตเฉพาะตำแหน่งเหล่านี้ (ไฟล์ + ข้อความในบรรทัด) */
const ALLOWED: { file: string; pattern: RegExp }[] = [
  { file: 'i18n/th.ts', pattern: /unsavedDot: '●'/ },
  { file: 'ui/store.ts', pattern: /isDirty \? ' ●' : ''/ },
  { file: 'i18n/th.ts', pattern: /fx: 'ƒx'/ },
];

function listFiles(dir: string): string[] {
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) return listFiles(p);
    return /\.(tsx?|css|html)$/.test(e.name) ? [p] : [];
  });
}

function findSymbols(rel: string, text: string): string[] {
  const hits: string[] = [];
  text.split(/\r?\n/).forEach((line, i) => {
    const trimmed = line.trim();
    if (trimmed.startsWith('//') || trimmed.startsWith('*') || trimmed.startsWith('/*')) return;
    // ตัดคอมเมนต์ท้ายบรรทัด (// ที่ไม่ได้อยู่ใน URL)
    const code = line.replace(/(^|[^:'"`])\/\/.*$/, '$1');
    if (!FORBIDDEN.test(code)) return;
    if (ALLOWED.some((a) => rel.replace(/\\/g, '/') === a.file && a.pattern.test(code))) return;
    hits.push(`${rel}:${i + 1}: ${trimmed}`);
  });
  return hits;
}

describe('ห้ามมีอีโมจิและสัญลักษณ์ตกแต่งใน src/', () => {
  it('detector catches emoji and decorative symbols, allows the kept ones', () => {
    for (const s of ['💾 บันทึก', '⚠️ เตือน', '✕', '① ขั้น', 'A → Z', '＋ เพิ่ม', '田 เส้นขอบ', '▾', 'Σ รวม', '⌘']) {
      expect(findSymbols('ui/X.tsx', s), s).toHaveLength(1);
    }
    for (const s of ['เรียงลำดับช่วง…', '{rows} แถว × {cols} คอลัมน์', 'ข้อความธรรมดา 50%–200%', '// คอมเมนต์ → ได้']) {
      expect(findSymbols('ui/X.tsx', s), s).toHaveLength(0);
    }
    expect(findSymbols('i18n/th.ts', "  unsavedDot: '●',")).toHaveLength(0);
    expect(findSymbols('ui/Other.tsx', "  unsavedDot: '●',")).toHaveLength(1);
  });

  it('src/ has no emoji or decorative symbols outside the allowed list', () => {
    const hits = listFiles(ROOT).flatMap((f) => findSymbols(path.relative(ROOT, f), fs.readFileSync(f, 'utf8')));
    expect(hits).toEqual([]);
  });
});
