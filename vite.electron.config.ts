import { defineConfig } from 'vite';
import path from 'path';

export default defineConfig({
  build: {
    outDir: 'dist-electron',
    emptyOutDir: false,
    target: 'node20',
    ssr: true,
    lib: {
      entry: {
        main: path.resolve(__dirname, 'electron/main.ts'),
        preload: path.resolve(__dirname, 'electron/preload.ts'),
      },
      formats: ['cjs'],
    },
    rollupOptions: {
      external: [
        'electron',
        'fs',
        'fs/promises',
        'path',
        'os',
        'child_process',
        'crypto',
        'events',
        'stream',
        'util',
        'buffer'
      ],
      output: {
        entryFileNames: '[name].js',
      },
    },
  },
});
