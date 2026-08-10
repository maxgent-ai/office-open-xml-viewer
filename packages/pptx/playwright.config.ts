import { defineConfig, devices } from '@playwright/test';

const vrtPort = Number(process.env.VRT_PORT ?? 5173);
if (!Number.isInteger(vrtPort) || vrtPort < 1 || vrtPort > 65_535) {
  throw new Error(`invalid VRT_PORT: ${process.env.VRT_PORT}`);
}
const privateCorpus = process.env.VRT_PRIVATE_CORPUS === '1';

export default defineConfig({
  testDir: './tests/visual',
  testMatch: '**/*.spec.ts',
  // Run slides sequentially for stable output
  fullyParallel: false,
  reporter: [
    ['list'],
    ['html', { outputFolder: 'tests/visual/report', open: 'never' }],
  ],
  use: {
    baseURL: `http://127.0.0.1:${vrtPort}`,
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        // Use the system-installed Google Chrome so fonts (Hiragino etc.)
        // and rendering exactly match what the user sees in the browser.
        channel: 'chrome',
        // Force DPR=1 so canvas physical size matches the 1280×720
        // PowerPoint reference images (toDataURL returns canvas.width × canvas.height).
        deviceScaleFactor: 1,
        viewport: { width: 1280, height: 720 },
      },
    },
  ],
  webServer: {
    command: `pnpm exec vite --host 127.0.0.1 --port ${vrtPort} --strictPort`,
    url: `http://127.0.0.1:${vrtPort}/tests/visual/fixture.html`,
    reuseExistingServer: privateCorpus ? false : !process.env.CI,
    timeout: 60_000,
  },
});
