import { defineConfig, devices } from '@playwright/test';

// Non-well-known port so concurrent agents don't collide. Override with PORT.
const PORT = process.env.PORT || '50650';
// Relative Vite base ('./') serves the app at the server root, so the e2e
// baseURL is the root too (specs navigate baseURL-relative via goto('./')).
const BASE = `http://localhost:${PORT}/`;

export default defineConfig({
  testDir: './tests',
  // Browser E2E specs only. Pure-unit assertions live in *.test.js and run
  // under Node's own test runner (`node --test`), not Playwright — excluding
  // them here keeps Playwright from importing/executing node:test files.
  testMatch: '**/*.spec.js',
  fullyParallel: true,
  reporter: 'list',
  timeout: 30000,
  use: {
    baseURL: BASE,
    trace: 'off',
    // DATEV-Lohn export triggers a file download; tests need to capture it.
    acceptDownloads: true,
  },
  webServer: {
    command: `npm run build && npm run preview -- --port ${PORT} --strictPort`,
    url: BASE,
    reuseExistingServer: true,
    timeout: 120000,
  },
  projects: [
    // Desktop runs the FULL suite — it is the source of truth for behaviour.
    { name: 'desktop', use: { ...devices['Desktop Chrome'], viewport: { width: 1280, height: 800 } } },
    // Mobile only re-runs the specs that assert something genuinely different
    // on a touch/narrow device. Every other spec is viewport-neutral — its
    // mobile run was a pure duplicate (the only per-project branches elsewhere
    // are screenshot side-effects, not extra assertions), so running them twice
    // doubled CI time for zero coverage. The three below earn their mobile run:
    //   • core-flow         — tap-to-log is a touch-only gesture
    //   • drag-resize-move  — touch long-press move + edge resize
    //   • ui                — responsive DAY_COLS + mobile-only menu clipping
    {
      name: 'mobile',
      use: { ...devices['Pixel 5'] },
      testMatch: [
        '**/core-flow.spec.js',
        '**/drag-resize-move.spec.js',
        '**/ui.spec.js',
      ],
    },
  ],
});
