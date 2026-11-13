import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

// Release build: minified, split into vendor chunks + lazily loaded feature panels.
// Formula engine (HyperFormula), file IO (ExcelJS/SheetJS), transforms (Arquero)
// and macros each live in their own Web Worker bundle.
export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    target: ['es2022', 'chrome100'], // Android System WebView is auto-updated on Android 8+
    minify: 'oxc',
    sourcemap: false,
    chunkSizeWarningLimit: 6000,
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (!id.includes('node_modules')) return undefined;
          if (id.includes('@univerjs')) return 'univer';
          if (id.includes('echarts') || id.includes('zrender')) return 'echarts';
          if (id.includes('react') || id.includes('scheduler')) return 'react';
          return undefined;
        },
      },
    },
  },
  worker: { format: 'es' },
  server: { host: true },
});
