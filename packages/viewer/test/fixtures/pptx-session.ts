import type { TextRun } from "../../src/contracts.js";
import { PptxEditEngine } from "../../src/edit/pptx/engine.js";
import { createOoxmlEditHandler } from "../../src/edit/pptx/handler.js";
import { loadPptxEditEngine } from "../../src/edit/pptx/provider.js";
import { PptxSession } from "../../src/edit/pptx/session.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../../src/edit/session.js";
import {
  defaultResourceLimits,
  type PptxEditSession,
} from "../../src/index.js";
import { loopbackWorker } from "./loopback-worker.js";

async function slideCountOf(bytes: Uint8Array): Promise<number> {
  const engine = await PptxEditEngine.open(bytes, defaultResourceLimits);
  try {
    return engine.pageCount;
  } finally {
    await engine.dispose();
  }
}

/**
 * A PPTX session over the loopback worker, with a host that only counts
 * slides by reading the package again, as the renderer would. With `runs`,
 * the session reads those text runs of each slide, as the shown deck
 * reports them; without, it is headless.
 */
export async function pptxSession(
  original: Uint8Array,
  runs?: readonly (readonly TextRun[])[],
): Promise<{ session: PptxEditSession; end(): Promise<void> }> {
  const signal = new AbortController().signal;
  const pair = loopbackWorker(createOoxmlEditHandler());
  const engine = await loadPptxEditEngine(
    original,
    { format: "pptx", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "pptx",
    limits: defaultResourceLimits,
    prepareDocument: async (bytes) => ({
      pageCount: await slideCountOf(bytes),
    }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
  };
  const core = new EditSessionController(
    engine,
    host,
    original,
    await slideCountOf(original),
  );
  const access = runs && {
    getTextRuns: async (pageIndex: number) => runs[pageIndex] ?? [],
    cachedPages: () => runs.map((_, index) => index),
  };
  return { session: new PptxSession(core, access), end: () => core.end() };
}
