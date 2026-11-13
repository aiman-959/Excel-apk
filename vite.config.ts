import react from '@vitejs/plugin-react';
import { defineConfig, type Plugin } from 'vite';
import fs from 'node:fs';

/**
 * Univer's renderer bundles hyphenation dictionaries for ~77 languages (≈5 MB) that are
 * only used by rich-text documents set to those languages. Spreadsheet cells do not
 * hyphenate and English patterns are built in, so each dictionary is replaced by an
 * empty one (same export, no patterns).
 */
const HYPHENATION = /@univerjs[\\/]engine-render[\\/]lib[\\/]es[\\/](?!index|worker)[a-z0-9-]+-[A-Za-z0-9_-]{8}\.js$/;
function dropHyphenation(): Plugin {
  return {
    name: 'petak-drop-hyphenation',
    enforce: 'pre',
    load(id) {
      if (!HYPHENATION.test(id)) return null;
      const name = /export \{ (\w+) \};\s*$/.exec(fs.readFileSync(id, 'utf8'))?.[1];
      return name ? `const ${name} = ["0", "{}", []];\nexport { ${name} };\n` : null;
    },
  };
}

// Release build: minified, split into vendor chunks + lazily loaded feature panels.
// Formula engine (HyperFormula), file IO (ExcelJS/SheetJS), transforms (Arquero)
// and macros each live in their own Web Worker bundle.
const ISOLATION = { 'Cross-Origin-Opener-Policy': 'same-origin', 'Cross-Origin-Embedder-Policy': 'credentialless' };

export default defineConfig({
  plugins: [react(), dropHyphenation()],
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
  // Cross-origin isolation enables SharedArrayBuffer between the UI and workers
  // (the Android app sets the same headers natively, see MainActivity.java).
  server: { host: true, headers: ISOLATION },
  preview: { headers: ISOLATION },
});
