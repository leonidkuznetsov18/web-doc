import { createOoxmlEditHandler } from "./edit/pptx/handler.js";
import { attachWorkerEndpoint } from "./worker-endpoint.js";

/*
 * Module worker that owns one OOXML package for one edit session: the ZIP
 * container, the scanned parts and the overlay of changes. It needs no
 * WebAssembly; inflation uses the platform DecompressionStream.
 */

attachWorkerEndpoint(
  self as unknown as DedicatedWorkerGlobalScope,
  createOoxmlEditHandler(),
);
