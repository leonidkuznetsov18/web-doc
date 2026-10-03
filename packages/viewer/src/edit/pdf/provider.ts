import type { RegisteredFont } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { WorkerRpcClient, type WorkerLike } from "../../worker-client.js";
import type {
  EditWorkerBitmap,
  EditWorkerFont,
  EditWorkerInitPayload,
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
} from "../../worker-protocol.js";
import type { EditEngineContext } from "../engine.js";
import type { PagePoint, PageRect, TextPosition, TextRange } from "../types.js";
import {
  packageRelativeUrl,
  resolveAssetUrl,
  WorkerEngineClient,
} from "../worker-engine.js";
import { pdfOperationSchemas } from "./schemas.js";
import type {
  PageBitmap,
  PageLayout,
  TextLayout,
  PdfTextParagraph,
} from "./types.js";

/** The reads behind the overlay primitives, beyond the core engine interface. */
export interface PdfEngineReads {
  textParagraph?(
    id: string,
    signal: AbortSignal,
  ): Promise<PdfTextParagraph | undefined>;
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

export class PdfEditEngineClient
  extends WorkerEngineClient
  implements PdfEngineReads
{
  readonly schemas = pdfOperationSchemas;

  async start(
    original: Uint8Array,
    options: PdfEditProviderOptions,
  ): Promise<void> {
    const base = this.context.assetBaseUrl;
    const init: EditWorkerInitPayload = {
      wasmUrl: resolveAssetUrl(
        options.wasmUrl,
        base,
        base
          ? new URL("assets/pdfium/pdfium.wasm", base)
          : packageRelativeUrl(
              "../../../assets/pdfium/pdfium.wasm",
              import.meta.url,
            ),
      ).href,
      fallbackFontUrl: resolveAssetUrl(
        options.fallbackFontUrl,
        base,
        base
          ? new URL("fonts/noto-sans-latin-cyrillic.ttf", base)
          : packageRelativeUrl(
              "../../../fonts/noto-sans-latin-cyrillic.ttf",
              import.meta.url,
            ),
      ).href,
    };
    await this.request("edit-init", init, this.context.signal);
    const data = original.slice().buffer;
    const fonts = (this.context.fonts ?? []).map(toWorkerFont);
    const open: EditWorkerOpenPayload = {
      data,
      limits: this.context.limits,
      format: "pdf",
      ...(this.context.fileName ? { fileName: this.context.fileName } : {}),
      ...(fonts.length > 0 ? { fonts } : {}),
    };
    await this.request<EditWorkerOpenResult>(
      "edit-open",
      open,
      this.context.signal,
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

  textParagraph(
    id: string,
    signal: AbortSignal,
  ): Promise<PdfTextParagraph | undefined> {
    return this.request("edit-text-paragraph", { id }, signal);
  }

  textLayout(id: string, signal: AbortSignal): Promise<TextLayout | undefined> {
    return this.request("edit-text-layout", { id }, signal);
  }

  positionAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<TextPosition | undefined> {
    return this.request("edit-position-at", { pageIndex, point }, signal);
  }

  rangeRects(
    range: TextRange,
    signal: AbortSignal,
  ): Promise<readonly PageRect[]> {
    return this.request("edit-range-rects", { range }, signal);
  }

  async renderWithout(
    pageIndex: number,
    elementIds: readonly string[],
    scale: number,
    signal: AbortSignal,
  ): Promise<PageBitmap> {
    const bitmap = await this.request<EditWorkerBitmap>(
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
    return this.request("edit-page-layout", { pageIndex }, signal);
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
      : packageRelativeUrl(
          "../../../workers/pdf-edit-worker.js",
          import.meta.url,
        ),
  );
  return new Worker(url, { type: "module", name: "web-doc-pdf-edit" });
}
