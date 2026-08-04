import { defineConfig } from 'vitest/config';
import { resolve } from 'node:path';

export default defineConfig({
  root: import.meta.dirname,
  resolve: {
    alias: {
      '@maxgent/ooxml/pptx': resolve(import.meta.dirname, '../pptx/src/types.ts'),
    },
  },
  test: {
    include: ['test/**/*.test.ts'],
  },
});
