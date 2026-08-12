import { defineConfig } from 'vitest/config';
import { resolve } from 'node:path';

/**
 * Live OfficeCLI contract tests. They execute the real `officecli` binary
 * against temporary decks, so they are opt-in: run them with
 * `OFFICECLI_LIVE=1 pnpm test:officecli-live`. Missing env or binary fails
 * loudly instead of skipping, so a green run always means the contract ran.
 */
export default defineConfig({
  root: import.meta.dirname,
  resolve: {
    alias: {
      '@maxgent/ooxml/pptx': resolve(import.meta.dirname, '../pptx/src/types.ts'),
    },
  },
  test: {
    include: ['test/officecli-live/**/*.live.test.ts'],
    // Every officecli invocation is an external process (0.5-6s each) and
    // spawns a per-file resident; run files serially to keep them stable.
    fileParallelism: false,
    testTimeout: 120_000,
    hookTimeout: 120_000,
  },
});
