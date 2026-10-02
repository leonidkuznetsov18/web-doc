import { defineConfig } from "@playwright/test";
import { existsSync } from "node:fs";

/*
 * Runs the scripted demos (`*.demo.ts`) that capture screenshots for tickets.
 * They are not part of `npm run test:e2e`; run them on demand:
 *
 *   node node_modules/@playwright/test/cli.js test --config=playwright.demo.config.ts
 */

const chromiumPath =
  process.env.CHROMIUM_PATH ??
  (existsSync("/usr/bin/chromium") ? "/usr/bin/chromium" : undefined);

export default defineConfig({
  testDir: "./tests/e2e",
  testMatch: ["*.demo.ts"],
  outputDir: "test-results/demo-run",
  workers: 1,
  retries: 0,
  use: {
    baseURL: "http://127.0.0.1:4173",
    launchOptions: chromiumPath ? { executablePath: chromiumPath } : {},
  },
  webServer: {
    command: "npm run dev --workspace @zrimo/example-vanilla",
    port: 4173,
    reuseExistingServer: !process.env.CI,
  },
});
