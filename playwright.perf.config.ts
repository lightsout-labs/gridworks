import { defineConfig, devices } from '@playwright/test';
import base from './playwright.config';

// The time limits of the app (R45, R46), measured on Chromium one test at a
// time, so that the tests do not slow each other down.
export default defineConfig({
  ...base,
  testDir: 'tests/perf',
  testMatch: '*.spec.ts',
  fullyParallel: false,
  workers: 1,
  projects: [{ name: 'chromium', use: { ...devices['Desktop Chrome'] } }],
});
