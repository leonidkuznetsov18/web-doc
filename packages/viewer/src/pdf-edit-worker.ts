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
    decodeImage: async (bytes, mimeType) => {
      const bitmap = await createImageBitmap(
        new Blob([new Uint8Array(bytes)], { type: mimeType }),
      );
      try {
        const canvas = new OffscreenCanvas(bitmap.width, bitmap.height);
        const context = canvas.getContext("2d")!;
        context.drawImage(bitmap, 0, 0);
        const { data } = context.getImageData(
          0,
          0,
          bitmap.width,
          bitmap.height,
        );
        return {
          width: bitmap.width,
          height: bitmap.height,
          rgba: new Uint8Array(data.buffer),
        };
      } finally {
        bitmap.close();
      }
    },
  }),
);
