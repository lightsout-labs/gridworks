import { defineConfig, devices } from '@playwright/test';

const browsers = [
  ['chromium', 'Desktop Chrome'],
  ['firefox', 'Desktop Firefox'],
  ['webkit', 'Desktop Safari'],
] as const;

// GRIDWORKS_URL points the tests at a deployed app instead of a local build.
const deployed = process.env.GRIDWORKS_URL;

export default defineConfig({
  testDir: 'tests/e2e',
  fullyParallel: true,
  forbidOnly: Boolean(process.env.CI),
  retries: process.env.CI ? 1 : 0,
  reporter: process.env.CI ? 'github' : 'list',
  use: {
    baseURL: deployed ?? 'http://localhost:4173/gridworks/',
    trace: 'retain-on-failure',
  },
  projects: browsers.map(([name, device]) => ({
    name,
    use: { ...devices[device] },
  })),
  webServer: deployed
    ? undefined
    : {
        command: 'npm run build && npm run preview',
        url: 'http://localhost:4173/gridworks/',
        reuseExistingServer: false,
      },
});
