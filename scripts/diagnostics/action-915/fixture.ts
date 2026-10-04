import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { performance } from "node:perf_hooks";
import { expect, test as base } from "@playwright/test";
import { createSink, sampleResources } from "./action-915-sink.js";

export { expect };

// Diagnostics only; the two historical test bodies retain their budgets/assertions.
export const test = base.extend({
  page: async ({ page }, use, info) => {
    const output = info.outputPath("worker-lifecycle.jsonl");
    const record = createSink(output);
    const hash = (path: string) =>
      createHash("sha256").update(readFileSync(path)).digest("hex");
    const browser = page.context().browser();
    record({
      event: "test-start",
      browserType: browser?.browserType().name() ?? "unavailable",
      browserVersion: browser?.version() ?? "unavailable",
      repeat: info.repeatEachIndex,
      worker: info.workerIndex,
      subject: "283fd8566398af7b91526c0972d440d80d160943",
      kit: process.env.GITHUB_SHA ?? "local-unqualified",
      lockSha256: hash("package-lock.json"),
      packageSha256: hash("packages/viewer/package.json"),
      playwrightSha256: hash("node_modules/@playwright/test/package.json"),
      testSha256: hash(info.file),
    });
    page.on("console", (message) => {
      const value = message.text();
      if (!value.startsWith("ACTION915:") || value.length > 4096) return;
      try {
        record(JSON.parse(value.slice("ACTION915:".length)));
      } catch {
        record({ event: "capture-invalid", status: "malformed-metadata" });
      }
    });
    page.on("crash", () => record({ event: "page-crash" }));
    page.on("pageerror", () => record({ event: "page-error" }));
    page.on("requestfailed", (request) => {
      // Never retain arbitrary document URLs, paths, query strings or messages.
      const path = new URL(request.url()).pathname;
      if (/^\/workers\/(ooxml-edit-worker\.js|pdf-edit-worker\.js)$/.test(path))
        record({
          event: "worker-fetch-failed",
          kind: path.endsWith("pdf-edit-worker.js") ? "pdf" : "ooxml",
        });
    });
    await page.addInitScript(
      ({ testWorker, repeat }) => {
        const navigation = `navigation-${testWorker}-${repeat}-${crypto.randomUUID()}`;
        const context = {
          navigation,
          format: "pending",
          clock: `renderer-${crypto.randomUUID()}`,
        };
        const ids = new WeakMap<object, number>();
        Reflect.set(globalThis, "__action915Context", context);
        Reflect.set(globalThis, "__action915WorkerIds", ids);
        const clock = context.clock;
        const emit = (
          event: string,
          worker: number,
          value: unknown,
          owner = context,
        ) => {
          const metadata: Record<string, string | number> = {
            ...owner,
            event,
            worker,
            clock,
            now: performance.now(),
          };
          if (value !== null && typeof value === "object")
            for (const key of [
              "id",
              "operation",
              "kind",
              "event",
              "clock",
              "now",
              "timeoutMs",
            ]) {
              const field: unknown = Reflect.get(value, key);
              if (typeof field === "number" && Number.isFinite(field))
                metadata[key] = field;
              else if (
                typeof field === "string" &&
                field.length <= 128 &&
                /^[a-zA-Z0-9_.:/-]+$/.test(field)
              )
                metadata[key] = field;
            }
          console.debug("ACTION915:" + JSON.stringify(metadata));
        };
        let beat = performance.now();
        let rendererSamples = 0;
        const rendererHeartbeat = setInterval(() => {
          const now = performance.now();
          emit("renderer-heartbeat", 0, { now });
          // Jitter is a renderer-only difference; worker clocks are never subtracted.
          console.debug(
            "ACTION915:" +
              JSON.stringify({
                ...context,
                event: "renderer-jitter",
                clock,
                now,
                jitterMs: now - beat - 5000,
              }),
          );
          beat = now;
          if (++rendererSamples === 72) clearInterval(rendererHeartbeat);
        }, 5000);
        const NativeWorker = window.Worker;
        let ordinal = 0;
        window.Worker = class extends NativeWorker {
          readonly diagnosticId: number;
          readonly diagnosticEnabled: boolean;
          readonly diagnosticOwner: typeof context;
          constructor(url: string | URL, options?: WorkerOptions) {
            super(url, options); // Original native URL, options and module loading.
            this.diagnosticId = ++ordinal;
            this.diagnosticEnabled =
              options?.name === "web-doc-ooxml-edit" ||
              options?.name === "web-doc-pdf-edit";
            this.diagnosticOwner = { ...context };
            ids.set(this, this.diagnosticId);
            if (!this.diagnosticEnabled) return;
            emit(
              "worker-create",
              this.diagnosticId,
              undefined,
              this.diagnosticOwner,
            );
            this.addEventListener("message", (event: MessageEvent<unknown>) => {
              const value = event.data;
              if (
                value !== null &&
                typeof value === "object" &&
                Reflect.get(value, "action915") === true
              )
                emit(
                  "worker-phase",
                  this.diagnosticId,
                  Reflect.get(value, "metadata"),
                  this.diagnosticOwner,
                );
              else
                emit(
                  "host-reply",
                  this.diagnosticId,
                  value,
                  this.diagnosticOwner,
                );
            });
            this.addEventListener("error", () =>
              emit(
                "worker-error",
                this.diagnosticId,
                undefined,
                this.diagnosticOwner,
              ),
            );
            this.addEventListener("messageerror", () =>
              emit(
                "worker-messageerror",
                this.diagnosticId,
                undefined,
                this.diagnosticOwner,
              ),
            );
          }
          postMessage(
            message: unknown,
            options?: Transferable[] | StructuredSerializeOptions,
          ) {
            if (this.diagnosticEnabled)
              emit(
                "host-send",
                this.diagnosticId,
                message,
                this.diagnosticOwner,
              );
            if (Array.isArray(options)) super.postMessage(message, options);
            else super.postMessage(message, options);
          }
          terminate() {
            if (this.diagnosticEnabled)
              emit(
                "worker-terminate",
                this.diagnosticId,
                undefined,
                this.diagnosticOwner,
              );
            super.terminate();
          }
        };
      },
      { testWorker: info.workerIndex, repeat: info.repeatEachIndex },
    );
    let beat = performance.now();
    let samples = 0;
    const heartbeat = setInterval(() => {
      const now = performance.now();
      record({
        event: "node-heartbeat",
        nodeNow: now,
        jitterMs: now - beat - 5000,
      });
      beat = now;
      sampleResources(record);
      if (++samples === 72) clearInterval(heartbeat);
    }, 5000);
    sampleResources(record);
    try {
      await use(page);
    } finally {
      clearInterval(heartbeat);
      record({ event: "test-end", status: info.status });
      await info.attach("worker-lifecycle", {
        path: output,
        contentType: "application/x-ndjson",
      });
    }
  },
});
