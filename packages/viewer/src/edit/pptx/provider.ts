import { ViewerError } from "../../errors.js";
import { WorkerRpcClient, type WorkerLike } from "../../worker-client.js";
import type { EditWorkerOpenPayload } from "../../worker-protocol.js";
import type { EditEngineContext } from "../engine.js";
import {
  packageRelativeUrl,
  resolveAssetUrl,
  WorkerEngineClient,
} from "../worker-engine.js";
import type { PptxEngineReads } from "./engine.js";
import { pptxOperationSchemas } from "./schemas.js";
import type { PptxLayoutInfo, PptxSlideInfo } from "./types.js";

export interface PptxEditProviderOptions {
  readonly workerUrl?: string | URL;
  /** Test hook: supplies the worker instead of the packaged script. */
  readonly createWorker?: () => WorkerLike;
}

/**
 * Starts the OOXML edit worker for `original` as a PPTX engine. Loaded
 * lazily by the Office adapter's provider on the first `edit()`.
 */
export async function loadPptxEditEngine(
  original: Uint8Array,
  context: EditEngineContext,
  options: PptxEditProviderOptions = {},
): Promise<PptxEditEngineClient> {
  const worker = options.createWorker
    ? options.createWorker()
    : createPackagedWorker(options, context);
  const rpc = new WorkerRpcClient(worker);
  const engine = new PptxEditEngineClient(rpc, context);
  try {
    await engine.start(original);
    return engine;
  } catch (error) {
    rpc.destroy();
    throw error;
  }
}

export class PptxEditEngineClient
  extends WorkerEngineClient
  implements PptxEngineReads
{
  readonly schemas = pptxOperationSchemas;

  async start(original: Uint8Array): Promise<void> {
    const data = original.slice().buffer;
    const open: EditWorkerOpenPayload = {
      data,
      limits: this.context.limits,
      format: "pptx",
      ...(this.context.fileName ? { fileName: this.context.fileName } : {}),
    };
    await this.request("edit-open", open, this.context.signal, [data]);
  }

  slides(signal: AbortSignal): Promise<readonly PptxSlideInfo[]> {
    return this.request("edit-pptx-slides", undefined, signal);
  }

  layouts(signal: AbortSignal): Promise<readonly PptxLayoutInfo[]> {
    return this.request("edit-pptx-layouts", undefined, signal);
  }
}

function createPackagedWorker(
  options: PptxEditProviderOptions,
  context: EditEngineContext,
): WorkerLike {
  if (typeof Worker === "undefined")
    throw new ViewerError(
      "edit-unsupported",
      "PPTX editing requires a browser module Worker",
      { details: { format: "pptx", reason: "no-worker" } },
    );
  const url = resolveAssetUrl(
    options.workerUrl,
    context.assetBaseUrl,
    context.assetBaseUrl
      ? new URL("workers/ooxml-edit-worker.js", context.assetBaseUrl)
      : packageRelativeUrl(
          "../../../workers/ooxml-edit-worker.js",
          import.meta.url,
        ),
  );
  return new Worker(url, { type: "module", name: "web-doc-ooxml-edit" });
}
