import { ViewerError } from "../../errors.js";
import { WorkerRpcClient, type WorkerLike } from "../../worker-client.js";
import type {
  EditWorkerInitPayload,
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
  EditWorkerOperation,
} from "../../worker-protocol.js";
import type { EditEngine, EditEngineContext, EngineChange } from "../engine.js";
import type {
  EditElement,
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  TextTarget,
} from "../types.js";
import { pdfOperationSchemas } from "./schemas.js";

/** How long disposal waits for the worker's answer before terminating it. */
const DISPOSE_GRACE_MS = 1000;

export interface PdfEditProviderOptions {
  readonly workerUrl?: string | URL;
  readonly wasmUrl?: string | URL;
  /** Test hook: supplies the worker instead of the packaged script. */
  readonly createWorker?: () => WorkerLike;
}

/**
 * Starts the PDF edit worker for `original` and returns the engine the core
 * drives. Loaded lazily by the PDF adapter's provider on the first `edit()`.
 */
export async function loadPdfEditEngine(
  original: Uint8Array,
  context: EditEngineContext,
  options: PdfEditProviderOptions = {},
): Promise<EditEngine> {
  const worker = options.createWorker
    ? options.createWorker()
    : createPackagedWorker(options, context);
  const rpc = new WorkerRpcClient(worker);
  const engine = new PdfEditEngineClient(rpc, context);
  try {
    await engine.start(original, options);
    return engine;
  } catch (error) {
    rpc.destroy();
    throw error;
  }
}

class PdfEditEngineClient implements EditEngine {
  readonly schemas = pdfOperationSchemas;
  readonly #rpc: WorkerRpcClient;
  readonly #context: EditEngineContext;

  constructor(rpc: WorkerRpcClient, context: EditEngineContext) {
    this.#rpc = rpc;
    this.#context = context;
  }

  async start(
    original: Uint8Array,
    options: PdfEditProviderOptions,
  ): Promise<void> {
    const init: EditWorkerInitPayload = {
      wasmUrl: resolveAssetUrl(
        options.wasmUrl,
        this.#context.assetBaseUrl,
        this.#context.assetBaseUrl
          ? new URL("assets/pdfium/pdfium.wasm", this.#context.assetBaseUrl)
          : packageRelativeUrl("../../../assets/pdfium/pdfium.wasm"),
      ).href,
    };
    await this.#request("edit-init", init, this.#context.signal);
    const data = original.slice().buffer;
    const open: EditWorkerOpenPayload = {
      data,
      limits: this.#context.limits,
      ...(this.#context.fileName ? { fileName: this.#context.fileName } : {}),
    };
    await this.#request<EditWorkerOpenResult>(
      "edit-open",
      open,
      this.#context.signal,
      [data],
    );
  }

  validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly OperationIssue[]> {
    return this.#request("edit-validate", { operations }, signal);
  }

  apply(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange> {
    return this.#request("edit-apply", { operations }, signal);
  }

  async materialize(signal: AbortSignal): Promise<Uint8Array> {
    const buffer = await this.#request<ArrayBuffer>(
      "edit-materialize",
      undefined,
      signal,
    );
    return new Uint8Array(buffer);
  }

  restore(
    batches: readonly (readonly EditOperation[])[],
    signal: AbortSignal,
  ): Promise<void> {
    return this.#request("edit-restore", { batches }, signal);
  }

  getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]> {
    return this.#request("edit-elements", { query }, signal);
  }

  getElement(
    id: string,
    signal: AbortSignal,
  ): Promise<EditElement | undefined> {
    return this.#request("edit-element", { id }, signal);
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]> {
    return this.#request("edit-elements-at", { pageIndex, point }, signal);
  }

  findText(
    query: string,
    options: EditFindOptions,
    signal: AbortSignal,
  ): Promise<readonly TextTarget[]> {
    return this.#request("edit-find-text", { query, options }, signal);
  }

  async dispose(): Promise<void> {
    // A healthy worker answers at once; a dead one never would, and nothing
    // is lost by terminating it without an answer.
    const request = this.#request("edit-dispose", undefined, undefined).catch(
      () => undefined,
    );
    await Promise.race([
      request,
      new Promise((resolve) => setTimeout(resolve, DISPOSE_GRACE_MS)),
    ]);
    this.#rpc.destroy();
  }

  #request<T>(
    operation: EditWorkerOperation,
    payload: unknown,
    signal: AbortSignal | undefined,
    transfer?: Transferable[],
  ): Promise<T> {
    return this.#rpc.request<T>(operation, payload, {
      ...(signal ? { signal } : {}),
      ...(transfer ? { transfer } : {}),
      timeoutMs: this.#context.limits.maxOperationMs,
    });
  }
}

function createPackagedWorker(
  options: PdfEditProviderOptions,
  context: EditEngineContext,
): WorkerLike {
  if (typeof Worker === "undefined")
    throw new ViewerError(
      "edit-unsupported",
      "PDF editing requires a browser module Worker",
      { details: { format: "pdf", reason: "no-worker" } },
    );
  const url = resolveAssetUrl(
    options.workerUrl,
    context.assetBaseUrl,
    context.assetBaseUrl
      ? new URL("workers/pdf-edit-worker.js", context.assetBaseUrl)
      : packageRelativeUrl("../../../workers/pdf-edit-worker.js"),
  );
  return new Worker(url, { type: "module", name: "web-doc-pdf-edit" });
}

function resolveAssetUrl(
  value: string | URL | undefined,
  base: URL | undefined,
  fallback: URL,
): URL {
  return value === undefined ? fallback : new URL(value, base);
}

function packageRelativeUrl(path: string): URL {
  // Data-driven so bundlers do not try to hash or inline the package-owned
  // directory; hosts that bundle the facade normally provide assetBaseUrl.
  const moduleUrl: string = import.meta.url;
  return new URL(path, moduleUrl);
}
