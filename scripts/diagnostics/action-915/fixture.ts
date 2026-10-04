import { appendFileSync, mkdirSync } from "node:fs";
import { dirname } from "node:path";
import { expect, test as base } from "@playwright/test";

export { expect };

// A dedicated diagnostic lane. Original tests retain their assertions and budgets.
export const test = base.extend({
  page: async ({ page }, use, info) => {
    const output = info.outputPath("worker-lifecycle.jsonl");
    mkdirSync(dirname(output), { recursive: true });
    const record = (details: object) => {
      appendFileSync(
        output,
        JSON.stringify({ at: Date.now(), ...details }) + "\n",
      );
    };
    record({
      event: "test-start",
      repeat: info.repeatEachIndex,
      worker: info.workerIndex,
    });
    page.on("console", (message) => {
      if (message.text().startsWith("ACTION915:")) {
        record({
          event: "worker",
          metadata: message.text().slice("ACTION915:".length),
        });
      }
    });
    page.on("crash", () => record({ event: "page-crash" }));
    page.on("pageerror", (error) =>
      record({ event: "page-error", name: error.name }),
    );
    page.on("requestfailed", (request) =>
      record({
        event: "request-failed",
        path: new URL(request.url()).pathname,
        reason: request.failure()?.errorText,
      }),
    );
    await page.route("**/workers/ooxml-edit-worker.js", async (route) => {
      const response = await route.fetch();
      const bootstrap = () => {
        console.debug(
          "ACTION915:" +
            JSON.stringify({ event: "boot", ms: performance.now() }),
        );
        self.addEventListener("message", (event: MessageEvent<unknown>) => {
          const value = event.data;
          if (
            value === null ||
            typeof value !== "object" ||
            !("kind" in value) ||
            value.kind !== "request"
          )
            return;
          const details: Record<string, unknown> = {
            event: "worker-request-received",
            ms: performance.now(),
          };
          for (const key of ["kind", "id", "operation"]) {
            const field: unknown = Reflect.get(value, key);
            if (typeof field === "string" || typeof field === "number")
              details[key] = field;
          }
          console.debug("ACTION915:" + JSON.stringify(details));
        });
      };
      await route.fulfill({
        response,
        body: `(${bootstrap.toString()})();\n` + (await response.text()),
      });
    });
    await page.addInitScript(() => {
      const NativeWorker = window.Worker;
      let ordinal = 0;
      const emit = (event: string, worker: number, value?: unknown) => {
        const details: Record<string, unknown> = {
          event,
          worker,
          ms: performance.now(),
        };
        if (value !== null && typeof value === "object") {
          for (const key of ["kind", "id", "operation"]) {
            if (!(key in value)) continue;
            const field: unknown = Reflect.get(value, key);
            if (typeof field === "number" || typeof field === "string")
              details[key] = field;
          }
        }
        console.debug("ACTION915:" + JSON.stringify(details));
      };
      window.Worker = class extends NativeWorker {
        readonly diagnosticId: number;
        readonly diagnosticEnabled: boolean;
        constructor(url: string | URL, options?: WorkerOptions) {
          super(url, options);
          this.diagnosticId = ++ordinal;
          this.diagnosticEnabled = options?.name === "web-doc-ooxml-edit";
          if (!this.diagnosticEnabled) return;
          emit("create", this.diagnosticId);
          this.addEventListener("message", (event) =>
            emit("received", this.diagnosticId, event.data),
          );
          this.addEventListener("error", () =>
            emit("error", this.diagnosticId),
          );
          this.addEventListener("messageerror", () =>
            emit("messageerror", this.diagnosticId),
          );
        }
        postMessage(
          message: unknown,
          options?: Transferable[] | StructuredSerializeOptions,
        ) {
          if (this.diagnosticEnabled) emit("sent", this.diagnosticId, message);
          if (Array.isArray(options)) super.postMessage(message, options);
          else super.postMessage(message, options);
        }
        terminate() {
          if (this.diagnosticEnabled) emit("terminate", this.diagnosticId);
          super.terminate();
        }
      };
    });
    try {
      await use(page);
    } finally {
      record({ event: "test-end", status: info.status });
      await info.attach("worker-lifecycle", {
        path: output,
        contentType: "application/x-ndjson",
      });
    }
  },
});
