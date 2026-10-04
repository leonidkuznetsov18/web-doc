import { copyFile, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createHash } from "node:crypto";

const subject = resolve(process.argv[2]);
const targetImport =
  'import { expect, test, type Page } from "@playwright/test";';
const replacement =
  'import type { Page } from "@playwright/test";\nimport { expect, test } from "./action-915-fixture.js";';
const hashes = [];
for (const name of ["edit-ai.spec.ts", "edit-pptx.spec.ts"]) {
  const file = resolve(subject, "tests/e2e", name);
  const original = await readFile(file, "utf8");
  if (original.split(targetImport).length !== 2)
    throw new Error(`Unexpected test import in ${name}`);
  const instrumented = original.replace(targetImport, replacement);
  if (instrumented.replace(replacement, targetImport) !== original)
    throw new Error(`Test body changed in ${name}`);
  hashes.push({
    name,
    originalSha256: createHash("sha256").update(original).digest("hex"),
  });
  await writeFile(file, instrumented);
}
await copyFile(
  new URL("./fixture.ts", import.meta.url),
  resolve(subject, "tests/e2e/action-915-fixture.ts"),
);
await writeFile(
  resolve(subject, "action-915-source-hashes.json"),
  JSON.stringify(hashes, null, 2) + "\n",
);
