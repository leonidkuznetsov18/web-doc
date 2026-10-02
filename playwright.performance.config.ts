import { defineConfig } from "@playwright/test";

import matrixConfig from "./playwright.matrix.config.js";

// Wall-clock budgets need an idle runner. Keep every matrix browser, but
// run the measurements separately from the parallel functional suite.
export default defineConfig({
  ...matrixConfig,
  grep: /@performance/,
  grepInvert: [],
  fullyParallel: false,
  workers: 1,
});
