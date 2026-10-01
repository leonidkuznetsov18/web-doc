import type { RegisteredFont, ViewerWarning } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { WorkerRpcClient, type WorkerLike } from "../../worker-client.js";
import type {
  EditWorkerBitmap,
  EditWorkerFont,
  EditWorkerInitPayload,
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
  EditWorkerOperation,
} from "../../worker-protocol.js";
import type {
  EditEngine,
  EditEngineContext,
  EngineBatch,
  EngineChange,
  MaterializedDocument,
  MaterializeOptions,
  RestoreTarget,
} from "../engine.js";
import type {
  EditElement,
  EditFindOptions,
  EditOperation,
  ElementQuery,
  OperationIssue,
  PagePoint,
  PageRect,
  TextPosition,
  TextRange,
  TextTarget,
} from "../types.js";
import { pdfOperationSchemas } from "./schemas.js";
import type { PageBitmap, PageLayout, TextLayout } from "./types.js";

/** The reads behind the overlay primitives, beyond the core engine interface. */
export interface PdfEngineReads {
  textLayout(id: string, signal: AbortSignal): Promise<TextLayout | undefined>;
  positionAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<TextPosition | undefined>;
  rangeRects(
    range: TextRange,
    signal: AbortSignal,
  ): Promise<readonly PageRect[]>;
  renderWithout(
    pageIndex: number,
    elementIds: readonly string[],
    scale: number,
    signal: AbortSignal,
  ): Promise<PageBitmap>;
  pageLayout(
    pageIndex: number,
    signal: AbortSignal,
  ): Promise<PageLayout | undefined>;
}

/** How long disposal waits for the worker's answer before terminating it. */
const DISPOSE_GRACE_MS = 1000;

export interface PdfEditProviderOptions {
  readonly workerUrl?: string | URL;
  readonly wasmUrl?: string | URL;
  /** TrueType font for text the standard fonts cannot encode. */
  readonly fallbackFontUrl?: string | URL;
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
): Promise<PdfEditEngineClient> {
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

export class PdfEditEngineClient implements EditEngine, PdfEngineReads {
  readonly schemas = pdfOperationSchemas;
  readonly #rpc: WorkerRpcClient;
  readonly #context: EditEngineContext;
  /** Highest state id seen, for batches passed as plain arrays. */
  #stateId = 0;

  constructor(rpc: WorkerRpcClient, context: EditEngineContext) {
    this.#rpc = rpc;
    this.#context = context;
  }

  async start(
    original: Uint8Array,
    options: PdfEditProviderOptions,
  ): Promise<void> {
    const base = this.#context.assetBaseUrl;
    const init: EditWorkerInitPayload = {
      wasmUrl: resolveAssetUrl(
        options.wasmUrl,
        base,
        base
          ? new URL("assets/pdfium/pdfium.wasm", base)
          : packageRelativeUrl("../../../assets/pdfium/pdfium.wasm"),
      ).href,
      fallbackFontUrl: resolveAssetUrl(
        options.fallbackFontUrl,
        base,
        base
          ? new URL("fonts/noto-sans-latin-cyrillic.ttf", base)
          : packageRelativeUrl("../../../fonts/noto-sans-latin-cyrillic.ttf"),
      ).href,
    };
    await this.#request("edit-init", init, this.#context.signal);
    const data = original.slice().buffer;
    const fonts = (this.#context.fonts ?? []).map(toWorkerFont);
    const open: EditWorkerOpenPayload = {
      data,
      limits: this.#context.limits,
      ...(this.#context.fileName ? { fileName: this.#context.fileName } : {}),
      ...(fonts.length > 0 ? { fonts } : {}),
    };
    await this.#request<EditWorkerOpenResult>(
      "edit-open",
      open,
      this.#context.signal,
      [
        data,
        ...fonts
          .map((font) => font.source)
          .filter(
            (source): source is ArrayBuffer => source instanceof ArrayBuffer,
          ),
      ],
    );
  }

  validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly OperationIssue[]> {
    return this.#request("edit-validate", { operations }, signal);
  }

  /** Plain operation arrays, as the unit tests pass them, become the next batch. */
  apply(
    input: EngineBatch | readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange> {
    const batch: EngineBatch = Array.isArray(input)
      ? { stateId: ++this.#stateId, operations: input }
      : (input as EngineBatch);
    this.#stateId = Math.max(this.#stateId, batch.stateId);
    return this.#request("edit-apply", { batch }, signal);
  }

  /** A bare signal, as the unit tests pass it, asks for the shown form. */
  async materialize(
    purposeOrSignal: "show" | "save" | AbortSignal = "show",
    options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<Uint8Array> {
    return (await this.materializeDocument(purposeOrSignal, options, signal))
      .bytes;
  }

  async materializeDocument(
    purposeOrSignal: "show" | "save" | AbortSignal = "show",
    options: MaterializeOptions = {},
    signal: AbortSignal = new AbortController().signal,
  ): Promise<MaterializedDocument> {
    const purpose =
      purposeOrSignal instanceof AbortSignal ? "show" : purposeOrSignal;
    const own =
      purposeOrSignal instanceof AbortSignal ? purposeOrSignal : signal;
    const result = await this.#request<{
      readonly data: ArrayBuffer;
      readonly warnings: readonly ViewerWarning[];
    }>("edit-materialize", { purpose, options }, own);
    return { bytes: new Uint8Array(result.data), warnings: result.warnings };
  }

  restore(
    input: RestoreTarget | readonly (readonly EditOperation[])[],
    signal: AbortSignal,
  ): Promise<void> {
    const target: RestoreTarget = Array.isArray(input)
      ? {
          batches: (input as readonly (readonly EditOperation[])[]).map(
            (operations, index) => ({ stateId: index + 1, operations }),
          ),
        }
      : (input as RestoreTarget);
    const base = target.base?.slice().buffer;
    return this.#request(
      "edit-restore",
      { batches: target.batches, ...(base ? { base } : {}) },
      signal,
      base ? [base] : undefined,
    );
  }

  putAsset(id: string, data: Uint8Array, signal: AbortSignal): Promise<void> {
    const buffer = data.slice().buffer;
    return this.#request("edit-put-asset", { id, data: buffer }, signal, [
      buffer,
    ]);
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

  textLayout(id: string, signal: AbortSignal): Promise<TextLayout | undefined> {
    return this.#request("edit-text-layout", { id }, signal);
  }

  positionAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<TextPosition | undefined> {
    return this.#request("edit-position-at", { pageIndex, point }, signal);
  }

  rangeRects(
    range: TextRange,
    signal: AbortSignal,
  ): Promise<readonly PageRect[]> {
    return this.#request("edit-range-rects", { range }, signal);
  }

  async renderWithout(
    pageIndex: number,
    elementIds: readonly string[],
    scale: number,
    signal: AbortSignal,
  ): Promise<PageBitmap> {
    const bitmap = await this.#request<EditWorkerBitmap>(
      "edit-render-without",
      { pageIndex, elementIds, scale },
      signal,
    );
    return { ...bitmap, data: new Uint8Array(bitmap.data) };
  }

  pageLayout(
    pageIndex: number,
    signal: AbortSignal,
  ): Promise<PageLayout | undefined> {
    return this.#request("edit-page-layout", { pageIndex }, signal);
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

/** A host font in the worker's shape: bytes are copied, URLs made absolute. */
function toWorkerFont(font: RegisteredFont): EditWorkerFont {
  const source =
    font.source instanceof ArrayBuffer
      ? font.source.slice(0)
      : font.source instanceof Uint8Array
        ? font.source.slice().buffer
        : new URL(font.source, globalThis.location?.href ?? "http://localhost/")
            .href;
  return {
    family: font.family,
    weight: font.weight ?? 400,
    style: font.style ?? "normal",
    source,
  };
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
