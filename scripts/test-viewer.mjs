import { execFileSync } from "node:child_process";
import { readdirSync, rmSync } from "node:fs";

const node = process.execPath;
// Compile into a clean directory: a test compiled on another branch would
// otherwise stay behind and run against code it was not written for.
rmSync(".test-dist", { recursive: true, force: true });
execFileSync(
  node,
  ["../../node_modules/typescript/bin/tsc", "-p", "tsconfig.test.json"],
  {
    stdio: "inherit",
  },
);
const tests = readdirSync(".test-dist/test")
  .filter((file) => file.endsWith(".test.js"))
  .map((file) => `.test-dist/test/${file}`);
execFileSync(node, ["--test", ...tests], {
  stdio: "inherit",
});
