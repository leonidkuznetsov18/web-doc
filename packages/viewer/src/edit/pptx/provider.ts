import { WorkerRpcClient } from "../../worker-client.js";
import type { EditWorkerOpenPayload } from "../../worker-protocol.js";
import type { EditEngineContext } from "../engine.js";
import {
  createOoxmlEditWorker,
  type OoxmlEditProviderOptions,
} from "../ooxml/worker.js";
import { WorkerEngineClient } from "../worker-engine.js";
import type { TextSpan } from "../range-style.js";
import type { PptxEngineReads } from "./engine.js";
import { pptxOperationSchemas } from "./schemas.js";
import type { PptxLayoutInfo, PptxSlideInfo, PptxTextStyle } from "./types.js";

export type PptxEditProviderOptions = OoxmlEditProviderOptions;

/**
 * Starts the OOXML edit worker for `original` as a PPTX engine. Loaded
 * lazily by the Office adapter's provider on the first `edit()`.
 */
export async function loadPptxEditEngine(
  original: Uint8Array,
  context: EditEngineContext,
  options: PptxEditProviderOptions = {},
): Promise<PptxEditEngineClient> {
  const worker = createOoxmlEditWorker(options, context);
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

  textStyle(
    id: string,
    span: TextSpan | undefined,
    signal: AbortSignal,
  ): Promise<Partial<PptxTextStyle> | undefined> {
    return this.request("edit-pptx-text-style", { id, span }, signal);
  }
}
