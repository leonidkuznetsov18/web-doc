import type { TextRun } from "../../src/contracts.js";
import { loadDocxEditEngine } from "../../src/edit/docx/provider.js";
import { DocxSession } from "../../src/edit/docx/session.js";
import type { EditSessionAccess } from "../../src/edit/engine.js";
import { createOoxmlEditHandler } from "../../src/edit/pptx/handler.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../../src/edit/session.js";
import {
  defaultResourceLimits,
  type DocxEditSession,
} from "../../src/index.js";
import { loopbackWorker } from "./loopback-worker.js";

/**
 * The runs a fake renderer reports per page: the geometry join is tested
 * without a layout engine by handing the session runs that carry
 * paragraph ids.
 */
export type FakePages = readonly (readonly TextRun[])[];

/** A text run of a paragraph on a page, as the Office adapter would report it. */
export function run(
  paragraphId: string,
  text: string,
  x: number,
  y: number,
  width = text.length * 6,
  height = 12,
): TextRun {
  return { text, x, y, width, height, paragraphId, textLayer: "docx" };
}

/**
 * A DOCX session over the loopback worker, with a host that counts pages
 * from the fake runs and an access object that serves them.
 */
export async function docxSession(
  original: Uint8Array,
  pages: FakePages,
  options: { readonly cached?: readonly number[] } = {},
): Promise<{
  session: DocxSession & DocxEditSession;
  reads: number[];
  end(): Promise<void>;
}> {
  const signal = new AbortController().signal;
  const pair = loopbackWorker(createOoxmlEditHandler());
  const engine = await loadDocxEditEngine(
    original,
    { format: "docx", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "docx",
    limits: defaultResourceLimits,
    prepareDocument: async () => ({ pageCount: pages.length }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
    pageOf: (paragraphId) => {
      const index = pages.findIndex((runs) =>
        runs.some((run) => run.paragraphId === paragraphId),
      );
      return index < 0 ? undefined : index;
    },
  };
  const core = new EditSessionController(engine, host, original, pages.length);
  const reads: number[] = [];
  const cached = new Set(options.cached ?? []);
  const access: EditSessionAccess = {
    getTextRuns: async (pageIndex) => {
      reads.push(pageIndex);
      cached.add(pageIndex);
      return pages[pageIndex] ?? [];
    },
    cachedPages: () => [...cached],
  };
  return {
    session: new DocxSession(core, access),
    reads,
    end: () => core.end(),
  };
}
