import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import { loadPdfEditEngine } from "../src/edit/pdf/provider.js";
import type {
  EditSession,
  EditStateChange,
  PdfEditSession,
  PdfBackend,
} from "../src/index.js";
import {
  createPdfAdapter,
  defaultResourceLimits,
  ViewerClient,
  ViewerError,
} from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";

/**
 * The worker handler, backed by the Node-loaded PDFium instance. While `held`
 * is set, text searches never answer, so a crash can land mid-request.
 */
function workerPair() {
  const handler = createPdfEditHandler(() => fixturePdfium());
  const state = { held: false };
  const pair = loopbackWorker(async (operation, payload, context) => {
    if (operation === "edit-find-text" && state.held)
      await new Promise<never>(() => {});
    return handler(operation, payload, context);
  });
  return { ...pair, state };
}

function context(signal = new AbortController().signal) {
  return { format: "pdf" as const, limits: defaultResourceLimits, signal };
}

function rejectsWith(code: string) {
  return (error: unknown): boolean => {
    assert.ok(error instanceof ViewerError, String(error));
    assert.equal(error.code, code, error.message);
    return true;
  };
}

/** A PDF.js-free backend so the viewer can open the fixture in Node. */
function fakeBackend(pageCount: number): PdfBackend {
  return {
    pageCount,
    async renderPage() {},
    async pageText() {
      return [];
    },
    close() {},
  };
}

describe("PDF edit engine over the worker protocol", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf(["Page one", "Page two"]);
  });

  it("initialises PDFium, opens the document and hands the original back", async () => {
    const pair = workerPair();
    const engine = await loadPdfEditEngine(original, context(), {
      createWorker: () => pair.worker,
    });
    assert.equal(engine.schemas.format, "pdf");
    const signal = new AbortController().signal;
    assert.deepEqual(await engine.materialize(signal), original);
    assert.deepEqual(
      (await engine.getElements({}, signal)).map((element) => element.id),
      ["p0:o0", "p1:o0"],
    );
    assert.deepEqual(
      (await engine.findText("Page", {}, signal)).map((hit) => hit.pageIndex),
      [0, 1],
    );
    assert.equal(
      (await engine.getElement?.("p1:o0", signal))?.text,
      "Page two",
    );
    const issues = await engine.validate([{ op: "nope" }], signal);
    assert.equal(issues[0]?.code, "unknown-operation");
    await engine.restore([], signal);
    await engine.dispose();
    assert.equal(pair.terminated(), true);
  });

  it("rejects a file PDFium cannot open and tears the worker down", async () => {
    const pair = workerPair();
    await assert.rejects(
      loadPdfEditEngine(new TextEncoder().encode("%PDF-1.7 nope"), context(), {
        createWorker: () => pair.worker,
      }),
      rejectsWith("invalid-file"),
    );
    assert.equal(pair.terminated(), true);
  });

  it("drives a PDF session through the viewer", async () => {
    const pairs: ReturnType<typeof workerPair>[] = [];
    const adapter = createPdfAdapter({
      open: async () => fakeBackend(2),
      edit: {
        createWorker: () => {
          const pair = workerPair();
          pairs.push(pair);
          return pair.worker;
        },
      },
    });
    const viewer = ViewerClient.create({ adapters: [adapter] }).createViewer();
    const states: EditStateChange[] = [];
    viewer.on("editstatechange", (state) => states.push(state));
    await viewer.load(original, { fileName: "fixture.pdf" });
    assert.equal(viewer.getDocumentInfo().capabilities?.editing, true);
    assert.equal(pairs.length, 0);

    const session = await viewer.edit();
    assert.equal(pairs.length, 1);
    assert.equal(session.format, "pdf");
    const narrowed: EditSession = session;
    if (narrowed.format === "pdf") {
      // Compile-time: the union narrows to the PDF session by `format`.
      const pdf: PdfEditSession = narrowed;
      assert.equal(pdf.schemas.format, "pdf");
    }
    assert.deepEqual(await session.save(), original);
    assert.equal(session.state.dirty, false);
    await assert.rejects(
      session.apply([{ op: "nope" } as never]),
      rejectsWith("invalid-operation"),
    );

    // A crashed worker ends this session; the next edit() starts a fresh one.
    pairs[0]!.state.held = true;
    const pending = session.findText("Page");
    await new Promise((resolve) => setTimeout(resolve, 0));
    pairs[0]!.crash();
    await assert.rejects(pending, rejectsWith("worker-crashed"));
    await assert.rejects(session.getElements(), rejectsWith("edit-failed"));
    const replacement = await viewer.edit();
    assert.notEqual(replacement, session);
    assert.equal(pairs.length, 2);
    assert.deepEqual(await replacement.save(), original);

    await viewer.close();
    assert.equal(pairs[1]!.terminated(), true);
    assert.equal(states.at(-1)?.active, false);
  });
});
