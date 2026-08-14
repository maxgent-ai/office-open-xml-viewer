import { defineConfig } from 'vite';
import wasm from 'vite-plugin-wasm';
import topLevelAwait from 'vite-plugin-top-level-await';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  root: __dirname,
  plugins: [wasm(), topLevelAwait()],
  resolve: {
    alias: {
      '@ooxml-test-three-d-addon': resolve(__dirname, '../../src/three-d.ts'),
      '@ooxml-test-region-map-addon': resolve(__dirname, '../../src/region-map.ts'),
    },
  },
  build: {
    // Serve public/ (sample fixtures) from the dev server for VRT, but don't
    // copy it into the published dist/.
    copyPublicDir: false,
    lib: {
      entry: resolve(__dirname, 'src/index.ts'),
      formats: ['es', 'cjs'],
      fileName: (format) => `index.${format === 'es' ? 'mjs' : 'cjs'}`,
    },
  },
  server: {
    port: 5179,
    fs: {
      allow: [__dirname, resolve(__dirname, '../..')],
    },
  },
});
