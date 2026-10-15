// ไฟล์ฟอนต์ที่ import เป็น asset (Vite) ได้เป็น URL หรือ data URL
declare module '*.woff2' {
  const src: string;
  export default src;
}
