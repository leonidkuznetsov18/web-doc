import { createPdfEditHandler } from "./edit/pdf/engine/handler.js";
import { Pdfium } from "./edit/pdf/engine/pdfium.js";
import { attachWorkerEndpoint } from "./worker-endpoint.js";

/*
 * Module worker that owns a PDFium instance for one edit session. The WASM is
 * fetched from the URL the main thread resolved, so hosting the assets
 * elsewhere (assetBaseUrl) needs no change here.
 */

async function fetchBytes(url: string): Promise<Uint8Array> {
  const response = await fetch(url);
  if (!response.ok)
    throw new Error(`Fetching ${url} failed with HTTP ${response.status}`);
  return new Uint8Array(await response.arrayBuffer());
}

attachWorkerEndpoint(
  self as unknown as DedicatedWorkerGlobalScope,
  createPdfEditHandler({
    loadPdfium: async (wasmUrl) => Pdfium.load(await fetchBytes(wasmUrl)),
    fetchBytes,
  }),
);
