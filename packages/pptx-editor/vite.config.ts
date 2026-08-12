import { resolve } from 'node:path';

import { dts } from 'rolldown-plugin-dts';
import { defineConfig } from 'vite';

export default defineConfig({
  plugins: [
    dts({
      generator: 'tsgo',
      sourcemap: true,
      tsconfig: './tsconfig.build.json',
    }),
  ],
  // Declaration modules are already valid output. Oxc must not transform them.
  oxc: {
    exclude: [/\.js$/, /\.d\.[cm]?ts$/],
  },
  build: {
    copyPublicDir: false,
    lib: {
      entry: resolve(import.meta.dirname, 'src/index.ts'),
      formats: ['es'],
      fileName: (_format, name) =>
        name.endsWith('.d') ? `${name.slice(0, -2)}.d.ts` : `${name}.js`,
    },
    minify: false,
    sourcemap: true,
    target: 'esnext',
    rollupOptions: {
      // The editor consumes the host application's matching viewer package.
      // Keep the peer external so the tarball never embeds a duplicate copy.
      external: [/^@maxgent\/ooxml(?:\/.*)?$/],
    },
  },
});
