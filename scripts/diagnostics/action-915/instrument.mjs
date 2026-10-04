import { copyFile, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { pathToFileURL } from "node:url";

export const SUBJECT = "283fd8566398af7b91526c0972d440d80d160943";
const expected = JSON.parse(
  await readFile(new URL("./source-hashes.json", import.meta.url), "utf8"),
);
const sha = (value) => createHash("sha256").update(value).digest("hex");

/** Pure guarded plan; artifact-only validation does not need a Git checkout. */
export function createPlan(files) {
  const output = new Map();
  const guards = [];
  for (const [name, hash] of Object.entries(expected))
    if (typeof files[name] !== "string" || sha(files[name]) !== hash)
      throw new Error(`Historical hash mismatch: ${name}`);
  const patch = (name, anchor, replacement, label) => {
    const value = output.get(name) ?? files[name];
    const count = value.split(anchor).length - 1;
    if (count !== 1)
      throw new Error(`Expected one ${label} anchor, found ${count}`);
    output.set(name, value.replace(anchor, replacement));
    guards.push({ name, label, count });
  };
  const imported = new Set();
  const inject = (name, anchor, lines, label, after = false) => {
    if (!imported.has(name)) {
      const relative = name.includes("/adapters/")
        ? "../action-915-diagnostics.js"
        : "./action-915-diagnostics.js";
      output.set(
        name,
        `import { ${name.endsWith("worker-endpoint.ts") ? "recordPhase, safeMetadata" : "recordPhase"} } from "${relative}";\n${output.get(name) ?? files[name]}`,
      );
      imported.add(name);
    }
    patch(name, anchor, after ? anchor + lines : lines + anchor, label);
  };
  const viewer = "packages/viewer/src/viewer.ts";
  inject(
    viewer,
    "    this.#assertAlive();\n    const generation = ++this.#generation;",
    '    recordPhase("load-before", { format: options.fileName?.split(".").at(-1) });\n',
    "load-before",
  );
  inject(
    viewer,
    '      this.#emit("ready", this.#state);',
    '\n      recordPhase("load-after", { format: detection.format });',
    "load-after",
    true,
  );
  patch(
    viewer,
    "    if (stale) await stale.core.end();",
    '    if (stale) {\n      recordPhase("stale-session-end-before", { format });\n      await stale.core.end();\n      recordPhase("stale-session-end-after", { format });\n    }',
    "stale-end",
  );
  inject(
    viewer,
    "      const engine = await provider.load(this.#original!.slice(), {",
    '      recordPhase("provider-await-before", { format });\n',
    "provider-before",
  );
  inject(
    viewer,
    "      if (generation !== this.#generation || operation.signal.aborted) {\n        await engine.dispose()",
    '      recordPhase("provider-await-after", { format });\n',
    "provider-after",
  );
  inject(
    viewer,
    "      this.#session = { core, session };",
    '\n      recordPhase("session-ready", { format });',
    "session-ready",
    true,
  );
  patch(
    viewer,
    "    if (session) await session.core.end();",
    '    if (session) {\n      recordPhase("session-end-before", { format: session.core.format });\n      await session.core.end();\n      recordPhase("session-end-after", { format: session.core.format });\n    }',
    "session-end",
  );
  const office = "packages/viewer/src/adapters/office.ts";
  const officeStart = files[office].indexOf(
    "    load: async (original, context) =>",
  );
  const officeEnd = files[office].indexOf("    createSession:", officeStart);
  const oldOffice = files[office].slice(officeStart, officeEnd);
  inject(office, oldOffice, "", "office-import-owner");
  patch(
    office,
    oldOffice,
    `    load: async (original, context) => {
      recordPhase("provider-import-before", { format: context.format });
      if (context.format === "docx") {
        const provider = await import("../edit/docx/provider.js");
        recordPhase("provider-import-after", { format: context.format });
        return provider.loadDocxEditEngine(original, context, this.#options.edit ?? {});
      }
      const provider = await import("../edit/pptx/provider.js");
      recordPhase("provider-import-after", { format: context.format });
      return provider.loadPptxEditEngine(original, context, this.#options.edit ?? {});
    },
`,
    "office-import",
  );
  const pdf = "packages/viewer/src/adapters/pdf.ts";
  const pdfStart = files[pdf].indexOf("    load: async (original, context) =>");
  const pdfEnd = files[pdf].indexOf("    createSession:", pdfStart);
  const oldPdf = files[pdf].slice(pdfStart, pdfEnd);
  inject(pdf, oldPdf, "", "pdf-import-owner");
  patch(
    pdf,
    oldPdf,
    `    load: async (original, context) => {
      recordPhase("provider-import-before", { format: context.format });
      const provider = await import("../edit/pdf/provider.js");
      recordPhase("provider-import-after", { format: context.format });
      return provider.loadPdfEditEngine(original, context, this.#options.edit ?? {});
    },
`,
    "pdf-import",
  );
  const client = "packages/viewer/src/worker-client.ts";
  inject(
    client,
    "    const id = this.#nextId++;",
    '\n    recordPhase("request-created", { id, operation, timeoutMs: options.timeoutMs ?? 0 }, this.#worker);',
    "request-arm",
    true,
  );
  inject(
    client,
    "              timeout: setTimeout(() => {",
    '\n                recordPhase("request-timeout-fire", { id, operation, timeoutMs: options.timeoutMs ?? 0 }, this.#worker);',
    "timeout-fire",
    true,
  );
  inject(
    client,
    "      this.#pending.set(id, pending);",
    '\n      if (pending.timeout) recordPhase("timer-armed", { id, operation, timeoutMs: options.timeoutMs ?? 0 }, this.#worker);',
    "timer-armed",
    true,
  );
  const endpoint = "packages/viewer/src/worker-endpoint.ts";
  inject(
    endpoint,
    "    const message = event.data;",
    '\n    recordPhase("worker-receipt", safeMetadata(message));',
    "worker-receipt",
    true,
  );
  inject(
    endpoint,
    '    scope.postMessage(\n      {\n        kind: "success",',
    '    recordPhase("worker-reply", { id: request.id, operation: request.operation, kind: "success" });\n',
    "worker-success",
  );
  inject(
    endpoint,
    "  } catch (error) {",
    '\n    recordPhase("worker-reply", { id: request.id, operation: request.operation, kind: "failure" });',
    "worker-failure",
    true,
  );
  for (const worker of ["ooxml", "pdf"]) {
    const name = `packages/viewer/src/${worker}-edit-worker.ts`;
    inject(
      name,
      "attachWorkerEndpoint(\n",
      'recordPhase("worker-module-body");\n',
      `${worker}-module-body`,
    );
  }
  const oldImport =
    'import { expect, test, type Page } from "@playwright/test";';
  const replacement =
    'import type { Page } from "@playwright/test";\nimport { expect, test } from "./action-915-fixture.js";';
  for (const name of ["edit-ai.spec.ts", "edit-pptx.spec.ts"])
    patch(`tests/e2e/${name}`, oldImport, replacement, "fixture-import");
  const ai = "tests/e2e/edit-ai.spec.ts";
  const target =
    'test("describes 500 pages and 500 slides within the operation budget and records the latency"';
  const targetIndex = output.get(ai).indexOf(target);
  if (targetIndex < 0 || output.get(ai).indexOf(target, targetIndex + 1) >= 0)
    throw new Error("Target test changed");
  // Only the targeted evaluation's phase labels are inserted; all timed
  // expressions, original assertions, operation budgets and awaits stay intact.
  const source = output.get(ai);
  let tail = source.slice(targetIndex);
  const edits = [
    [
      "      const session = await viewer.edit();",
      '      console.debug("ACTION915:" + JSON.stringify({ ...Reflect.get(globalThis, "__action915Context"), event: "edit-before", now: performance.now() }));\n      const session = await viewer.edit();\n      console.debug("ACTION915:" + JSON.stringify({ ...Reflect.get(globalThis, "__action915Context"), event: "edit-after", now: performance.now() }));',
    ],
    [
      "      const started = performance.now();\n      const description = await session.describe();\n      const elapsed = performance.now() - started;",
      '      console.debug("ACTION915:" + JSON.stringify({ ...Reflect.get(globalThis, "__action915Context"), event: "cold-describe-before", now: performance.now() }));\n      const started = performance.now();\n      const description = await session.describe();\n      const elapsed = performance.now() - started;\n      console.debug("ACTION915:" + JSON.stringify({ ...Reflect.get(globalThis, "__action915Context"), event: "cold-describe-after", now: performance.now() }));',
    ],
    [
      "      const again = performance.now();\n      await session.describe();\n      const warm = performance.now() - again;",
      '      console.debug("ACTION915:" + JSON.stringify({ ...Reflect.get(globalThis, "__action915Context"), event: "warm-describe-before", now: performance.now() }));\n      const again = performance.now();\n      await session.describe();\n      const warm = performance.now() - again;\n      console.debug("ACTION915:" + JSON.stringify({ ...Reflect.get(globalThis, "__action915Context"), event: "warm-describe-after", now: performance.now() }));',
    ],
  ];
  for (const [before, after] of edits) {
    const count = tail.split(before).length - 1;
    if (count !== 1) throw new Error("Description phase anchor changed");
    tail = tail.replace(before, after);
    guards.push({ name: ai, label: "description-phase", count });
  }
  output.set(ai, source.slice(0, targetIndex) + tail);
  return { output, guards };
}

async function main() {
  const subject = resolve(process.argv[2]);
  const head = execFileSync("git", ["-C", subject, "rev-parse", "HEAD"], {
    encoding: "utf8",
  }).trim();
  if (head !== SUBJECT) throw new Error("Wrong historical subject");
  const files = Object.fromEntries(
    await Promise.all(
      Object.keys(expected).map(async (name) => [
        name,
        await readFile(resolve(subject, name), "utf8"),
      ]),
    ),
  );
  const plan = createPlan(files);
  for (const [name, value] of plan.output)
    await writeFile(resolve(subject, name), value);
  for (const [from, to] of [
    ["fixture.ts", "tests/e2e/action-915-fixture.ts"],
    ["sink.ts", "tests/e2e/action-915-sink.ts"],
    ["phase.ts", "packages/viewer/src/action-915-diagnostics.ts"],
  ])
    await copyFile(new URL(`./${from}`, import.meta.url), resolve(subject, to));
  await writeFile(
    resolve(subject, "action-915-source-hashes.json"),
    JSON.stringify(
      {
        subject: head,
        expected,
        guards: plan.guards,
        instrumented: Object.fromEntries(
          [...plan.output].map(([name, value]) => [name, sha(value)]),
        ),
      },
      null,
      2,
    ) + "\n",
  );
}
if (
  process.argv[1] &&
  pathToFileURL(resolve(process.argv[1])).href === import.meta.url
)
  await main();
