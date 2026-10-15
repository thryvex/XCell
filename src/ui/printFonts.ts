// ฟอนต์ Sarabun สำหรับเอกสารพิมพ์/PDF/ตัวอย่างก่อนพิมพ์ ใช้ไฟล์จาก @fontsource/sarabun (dependency เดิมของแอป)
// ตอน build ไฟล์เหล่านี้ถูกฝังเป็น data URL (vite.config.ts: assetsInlineLimit) จึงไม่ต้องโหลดอะไรจากภายนอก
// ตอน dev ได้เป็น URL จึงแปลงเป็น data URL ด้วยการ fetch ครั้งเดียวแล้วเก็บไว้
import thai400 from '@fontsource/sarabun/files/sarabun-thai-400-normal.woff2';
import thai700 from '@fontsource/sarabun/files/sarabun-thai-700-normal.woff2';
import thai400i from '@fontsource/sarabun/files/sarabun-thai-400-italic.woff2';
import thai700i from '@fontsource/sarabun/files/sarabun-thai-700-italic.woff2';
import latin400 from '@fontsource/sarabun/files/sarabun-latin-400-normal.woff2';
import latin700 from '@fontsource/sarabun/files/sarabun-latin-700-normal.woff2';
import latin400i from '@fontsource/sarabun/files/sarabun-latin-400-italic.woff2';
import latin700i from '@fontsource/sarabun/files/sarabun-latin-700-italic.woff2';

// unicode-range ตาม @fontsource/sarabun
const THAI_RANGE = 'U+02D7,U+0303,U+0331,U+0E01-0E5B,U+200C-200D,U+25CC';
const LATIN_RANGE =
  'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD';

const FACES: Array<{ src: string; weight: number; style: 'normal' | 'italic'; range: string }> = [
  { src: thai400, weight: 400, style: 'normal', range: THAI_RANGE },
  { src: thai700, weight: 700, style: 'normal', range: THAI_RANGE },
  { src: thai400i, weight: 400, style: 'italic', range: THAI_RANGE },
  { src: thai700i, weight: 700, style: 'italic', range: THAI_RANGE },
  { src: latin400, weight: 400, style: 'normal', range: LATIN_RANGE },
  { src: latin700, weight: 700, style: 'normal', range: LATIN_RANGE },
  { src: latin400i, weight: 400, style: 'italic', range: LATIN_RANGE },
  { src: latin700i, weight: 700, style: 'italic', range: LATIN_RANGE },
];

async function toDataUrl(src: string): Promise<string> {
  if (src.startsWith('data:')) return src;
  const res = await fetch(src);
  const bytes = new Uint8Array(await res.arrayBuffer());
  let binary = '';
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return `data:font/woff2;base64,${btoa(binary)}`;
}

let cached: Promise<string> | null = null;

/** CSS @font-face ของ Sarabun ที่ฝังฟอนต์เป็น data URL (ถ้าโหลดไม่ได้คืน '' แล้วใช้ฟอนต์สำรองของระบบ) */
export function getPrintFontCss(): Promise<string> {
  if (!cached) {
    cached = Promise.all(
      FACES.map(async (f) => {
        const url = await toDataUrl(f.src);
        return `@font-face { font-family: 'Sarabun'; src: url(${url}) format('woff2'); font-weight: ${f.weight}; font-style: ${f.style}; unicode-range: ${f.range}; }`;
      })
    )
      .then((faces) => faces.join('\n'))
      .catch(() => '');
  }
  return cached;
}
