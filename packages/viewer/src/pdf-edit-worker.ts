import { createPdfEditHandler } from "./edit/pdf/engine/handler.js";
import { Pdfium } from "./edit/pdf/engine/pdfium.js";
import { attachWorkerEndpoint } from "./worker-endpoint.js";

/*
 * Module worker that owns a PDFium instance for one edit session. The WASM is
 * fetched from the URL the main thread resolved, so hosting the assets
 * elsewhere (assetBaseUrl) needs no change here.
 */

attachWorkerEndpoint(
  self as unknown as DedicatedWorkerGlobalScope,
  createPdfEditHandler(async (wasmUrl) => {
    const response = await fetch(wasmUrl);
    if (!response.ok)
      throw new Error(
        `Fetching ${wasmUrl} failed with HTTP ${response.status}`,
      );
    return Pdfium.load(await response.arrayBuffer());
  }),
);
