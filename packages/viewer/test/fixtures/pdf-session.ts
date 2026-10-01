import { createPdfEditHandler } from "../../src/edit/pdf/engine/handler.js";
import { loadPdfEditEngine } from "../../src/edit/pdf/provider.js";
import { PdfSession } from "../../src/edit/pdf/session.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../../src/edit/session.js";
import { defaultResourceLimits, type PdfEditSession } from "../../src/index.js";
import { loopbackWorker } from "./loopback-worker.js";
import { fixturePdfium } from "./pdf-builder.js";

async function pageCountOf(bytes: Uint8Array): Promise<number> {
  const pdfium = await fixturePdfium();
  const document = pdfium.openDocument(bytes);
  try {
    return pdfium.lib.FPDF_GetPageCount(document.handle);
  } finally {
    document.close();
  }
}

/** A PDF session over the loopback worker, with a host that only counts pages. */
export async function pdfSession(
  original: Uint8Array,
): Promise<{ session: PdfEditSession; end(): Promise<void> }> {
  const signal = new AbortController().signal;
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      fetchBytes: async () => {
        throw new Error("no fonts");
      },
      decodeImage: async () => {
        throw new Error("no images");
      },
    }),
  );
  const engine = await loadPdfEditEngine(
    original,
    { format: "pdf", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "pdf",
    limits: defaultResourceLimits,
    prepareDocument: async (bytes) => ({ pageCount: await pageCountOf(bytes) }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
  };
  const core = new EditSessionController(
    engine,
    host,
    original,
    await pageCountOf(original),
  );
  return { session: new PdfSession(core), end: () => core.end() };
}
