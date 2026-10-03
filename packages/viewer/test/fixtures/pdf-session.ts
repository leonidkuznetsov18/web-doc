import { readFileSync } from "node:fs";
import { createRequire } from "node:module";

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

const FALLBACK_URL = "https://fonts.test/noto.ttf";
let fallbackFont: Uint8Array | undefined;

/**
 * A PDF session over the loopback worker, with a host that only counts
 * pages. With `fallbackFont`, the packaged Noto Sans subset answers the
 * engine's fallback-font fetch; without it no font can be fetched.
 */
export async function pdfSession(
  original: Uint8Array,
  options: {
    readonly fallbackFont?: boolean;
    readonly compact?: (bytes: Uint8Array) => Uint8Array;
  } = {},
): Promise<{ session: PdfEditSession; end(): Promise<void> }> {
  const signal = new AbortController().signal;
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      ...(options.compact ? { compact: options.compact } : {}),
      fetchBytes: async (url) => {
        if (options.fallbackFont && url === FALLBACK_URL)
          return (fallbackFont ??= new Uint8Array(
            readFileSync(
              new URL(
                "../../../fonts/noto-sans-latin-cyrillic.ttf",
                import.meta.url,
              ),
            ),
          ));
        const file = new URL(url).pathname.split("/").at(-1);
        if (
          options.fallbackFont &&
          file &&
          /^LiberationSans-(Regular|Bold|Italic|BoldItalic)\.ttf$/.test(file)
        )
          return new Uint8Array(
            readFileSync(
              createRequire(import.meta.url).resolve(
                `pdfjs-dist/standard_fonts/${file}`,
              ),
            ),
          );
        throw new Error(`No font at ${url}`);
      },
      decodeImage: async () => {
        throw new Error("no images");
      },
    }),
  );
  const engine = await loadPdfEditEngine(
    original,
    { format: "pdf", limits: defaultResourceLimits, signal },
    {
      createWorker: () => pair.worker,
      ...(options.fallbackFont ? { fallbackFontUrl: FALLBACK_URL } : {}),
    },
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
