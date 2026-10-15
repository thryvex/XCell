/// <reference types="vitest" />
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  plugins: [react()],
  base: './',
  resolve: {
    alias: {
      '@': path.resolve(__dirname, './src'),
    },
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    // ฝังไฟล์ฟอนต์ Sarabun (woff2 จาก @fontsource) เป็น data URL: เอกสารพิมพ์/PDF ใช้ได้โดยไม่ต้องโหลดไฟล์ภายนอก
    assetsInlineLimit: (filePath: string) =>
      /@fontsource[\\/]sarabun[\\/]files[\\/].*\.woff2$/.test(filePath) ? true : undefined,
  },
  server: {
    port: 5173,
    strictPort: true,
  },
  test: {
    globals: true,
    environment: 'node',
  },
});
