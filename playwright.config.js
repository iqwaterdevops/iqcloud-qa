// @ts-check
const { defineConfig, devices } = require('@playwright/test');

/** @type {import('@playwright/test').PlaywrightTestConfig} */
module.exports = defineConfig({
  testDir: './tests',
  timeout: 30 * 1000,
  expect: {
    timeout: 5000,
  },
  fullyParallel: true,
  reporter: [
    ['list'],
    ['html', { open: 'never' }],
  ],
  use: {
    baseURL: 'https://iqtrackitng-test.azurewebsites.net',
    headless: true,
    screenshot: 'only-on-failure',
    video: 'retain-on-failure',
    trace: 'retain-on-failure',
    locale: 'en-US',
  },
  projects: [
    {
      name: 'chromium-en',
      use: {
        ...devices['Desktop Chrome'],
        channel: 'chrome',
        locale: 'en-US',
      },
    },
  ],
});

