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
    // Live OfficeCLI contract tests execute the real binary; they run only
    // through `pnpm test:officecli-live` (vitest.officecli-live.config.ts).
    exclude: ['**/node_modules/**', 'test/officecli-live/**'],
  },
});
