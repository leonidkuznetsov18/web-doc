import assert from "node:assert/strict";
import { test } from "node:test";

import { plan } from "./verify.mjs";

const scripts = (files, mode) => plan(files, mode).map((check) => check.script);

test("a documentation change builds and checks the pages, not the viewer", () => {
  assert.deepEqual(scripts(["docs/api/editing.md"]), [
    "check",
    "pages:build",
    "test:pages",
  ]);
});

test("a viewer change runs the browser suites and the package checks", () => {
  const run = scripts(["packages/viewer/src/viewport.ts"]);
  for (const script of [
    "check",
    "test:pack",
    "fuzz:js",
    "report:size",
    "test:e2e",
    "test:e2e:matrix",
    "test:e2e:performance",
    "pages:build",
  ])
    assert.ok(run.includes(script), script);
  assert.ok(!run.includes("audit:vulnerabilities"));
  assert.ok(!run.includes("fuzz:rust"));
});

test("a lockfile change audits dependencies and a Rust change fuzzes", () => {
  assert.ok(scripts(["package-lock.json"]).includes("audit:vulnerabilities"));
  assert.ok(scripts(["crates/core/src/lib.rs"]).includes("fuzz:rust"));
});

test("a browser spec change runs the browser suites without packaging", () => {
  const run = scripts(["tests/e2e/edit-pdf.spec.ts"]);
  assert.ok(run.includes("test:e2e:matrix"));
  assert.ok(!run.includes("test:pack"));
});

test("full runs every check and quick only the core check", () => {
  assert.equal(
    scripts([], "full").length,
    plan(["packages/x", "crates/x", "docs/x", "package-lock.json"]).length,
  );
  assert.deepEqual(scripts(["packages/viewer/src/index.ts"], "quick"), [
    "check",
  ]);
});
