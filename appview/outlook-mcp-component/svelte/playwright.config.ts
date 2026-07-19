import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: 'e2e',
  timeout: 30_000,
  retries: 1,
  use: {
    baseURL: 'https://outlook.etzhayyim.com',
    headless: true,
  },
});
