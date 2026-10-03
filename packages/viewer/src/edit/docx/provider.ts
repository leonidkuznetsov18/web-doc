import { WorkerRpcClient } from "../../worker-client.js";
import type { EditWorkerOpenPayload } from "../../worker-protocol.js";
import type { EditEngineContext } from "../engine.js";
import {
  createOoxmlEditWorker,
  type OoxmlEditProviderOptions,
} from "../ooxml/worker.js";
import { WorkerEngineClient } from "../worker-engine.js";
import type { DocxDraftDocument, DocxEngineReads } from "./engine.js";
import { docxOperationSchemas } from "./schemas.js";
import type { TextSpan } from "../range-style.js";
import type {
  DocxTextPreviewFields,
  DocxRevision,
  DocxTextStyle,
} from "./types.js";

export type DocxEditProviderOptions = OoxmlEditProviderOptions;

/**
 * Starts the OOXML edit worker for `original` as a DOCX engine. Loaded
 * lazily by the Office adapter's provider on the first `edit()`.
 */
export async function loadDocxEditEngine(
  original: Uint8Array,
  context: EditEngineContext,
  options: DocxEditProviderOptions = {},
): Promise<DocxEditEngineClient> {
  const worker = createOoxmlEditWorker(options, context);
  const rpc = new WorkerRpcClient(worker);
  const engine = new DocxEditEngineClient(rpc, context);
  try {
    await engine.start(original);
    return engine;
  } catch (error) {
    rpc.destroy();
    throw error;
  }
}

export class DocxEditEngineClient
  extends WorkerEngineClient
  implements DocxEngineReads
{
  readonly schemas = docxOperationSchemas;

  async previewText(
    fields: DocxTextPreviewFields,
    signal: AbortSignal,
  ): Promise<Uint8Array> {
    return (await this.previewDraft(fields, signal)).bytes;
  }

  async previewDraft(
    fields: DocxTextPreviewFields,
    signal: AbortSignal,
  ): Promise<DocxDraftDocument> {
    const draft = await this.request<{
      readonly bytes: ArrayBuffer;
      readonly paragraph?: DocxDraftDocument["paragraph"];
    }>("edit-docx-preview-text", { fields }, signal);
    return {
      bytes: new Uint8Array(draft.bytes),
      ...(draft.paragraph ? { paragraph: draft.paragraph } : {}),
    };
  }

  revisions(id: string, signal: AbortSignal): Promise<readonly DocxRevision[]> {
    return this.request("edit-docx-revisions", { id }, signal);
  }

  textStyle(
    id: string,
    span: TextSpan | undefined,
    signal: AbortSignal,
  ): Promise<Partial<DocxTextStyle> | undefined> {
    return this.request("edit-docx-text-style", { id, span }, signal);
  }

  textColors(
    id: string,
    spans: readonly TextSpan[],
    signal: AbortSignal,
  ): Promise<readonly string[] | undefined> {
    return this.request("edit-docx-text-colors", { id, spans }, signal);
  }

  async start(original: Uint8Array): Promise<void> {
    const data = original.slice().buffer;
    const open: EditWorkerOpenPayload = {
      data,
      limits: this.context.limits,
      format: "docx",
      ...(this.context.fileName ? { fileName: this.context.fileName } : {}),
    };
    await this.request("edit-open", open, this.context.signal, [data]);
  }
}
