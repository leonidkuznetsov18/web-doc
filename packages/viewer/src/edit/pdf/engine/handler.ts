import { ViewerError } from "../../../errors.js";
import type { WorkerOperationHandler } from "../../../worker-endpoint.js";
import type {
  EditWorkerInitPayload,
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
  WorkerOperation,
} from "../../../worker-protocol.js";
import type { EngineBatch } from "../../engine.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  PagePoint,
} from "../../types.js";
import { pdfOperationSchemas } from "../schemas.js";
import { PdfEditDocument } from "./document.js";
import { AssetStore } from "../../assets.js";
import { FontLibrary } from "./fonts.js";
import type { ImageDecoder } from "./images.js";
import type { Pdfium } from "./pdfium.js";

export interface PdfEditHost {
  /** Instantiates PDFium for a WASM URL. */
  loadPdfium(wasmUrl: string): Promise<Pdfium>;
  /** Fetches font bytes for a URL. */
  fetchBytes(url: string): Promise<Uint8Array>;
  /** Decodes an image into RGBA pixels. */
  decodeImage: ImageDecoder;
}

/**
 * Serves the edit worker protocol for one PDF document. It runs inside the
 * edit worker in the browser and directly in Node tests; the host supplies
 * how WebAssembly and fonts are obtained.
 */
export function createPdfEditHandler(
  host: PdfEditHost,
): WorkerOperationHandler {
  let pdfium: Pdfium | undefined;
  let state: PdfEditDocument | undefined;
  const fonts = new FontLibrary((url) => host.fetchBytes(url));
  const assets = new AssetStore();

  const engine = (): PdfEditDocument => {
    if (!state)
      throw new ViewerError("lifecycle-error", "No PDF is open for editing");
    return state;
  };

  return async (operation: WorkerOperation, payload: unknown) => {
    switch (operation) {
      case "edit-init": {
        const init = payload as EditWorkerInitPayload;
        pdfium ??= await host.loadPdfium(init.wasmUrl);
        fonts.setFallbackUrl(init.fallbackFontUrl);
        return undefined;
      }
      case "edit-open": {
        if (!pdfium)
          throw new ViewerError("lifecycle-error", "PDFium is not initialised");
        const open = payload as EditWorkerOpenPayload;
        state?.dispose();
        fonts.register(open.fonts ?? []);
        state = new PdfEditDocument(
          pdfium,
          new Uint8Array(open.data),
          fonts,
          open.limits,
          assets,
        );
        const result: EditWorkerOpenResult = { pageCount: state.pageCount };
        return result;
      }
      case "edit-validate": {
        const { operations } = payload as {
          readonly operations: readonly EditOperation[];
        };
        await fonts.prepare(engine().fontRequests(operations));
        await engine().images.prepare(operations, host.decodeImage, assets);
        return engine().validate(operations);
      }
      case "edit-apply": {
        const { batch } = payload as { readonly batch: EngineBatch };
        await fonts.prepare(engine().fontRequests(batch.operations));
        await engine().images.prepare(
          batch.operations,
          host.decodeImage,
          assets,
        );
        return engine().apply(batch);
      }
      case "edit-put-asset": {
        const { id, data } = payload as {
          readonly id: string;
          readonly data: ArrayBuffer;
        };
        assets.set(id, new Uint8Array(data));
        return undefined;
      }
      case "edit-materialize": {
        const { purpose, options } = (payload ?? {}) as {
          readonly purpose?: "show" | "save";
          readonly options?: { readonly mode?: unknown };
        };
        const mode = options?.mode;
        return engine().materialize(
          purpose ?? "show",
          mode === "full" || mode === "incremental" ? mode : undefined,
        ).buffer;
      }
      case "edit-restore": {
        const { batches, base } = payload as {
          readonly batches: readonly EngineBatch[];
          readonly base?: ArrayBuffer;
        };
        const operations = batches.flatMap((batch) => batch.operations);
        await fonts.prepare(engine().fontRequests(operations));
        await engine().images.prepare(operations, host.decodeImage, assets);
        engine().restore({
          batches,
          ...(base ? { base: new Uint8Array(base) } : {}),
        });
        return undefined;
      }
      case "edit-elements":
        return engine().getElements(
          (payload as { readonly query: ElementQuery }).query,
        );
      case "edit-element":
        return engine().getElement((payload as { readonly id: string }).id);
      case "edit-elements-at": {
        const { pageIndex, point } = payload as {
          readonly pageIndex: number;
          readonly point: PagePoint;
        };
        return engine().elementsAt(pageIndex, point);
      }
      case "edit-find-text": {
        const { query, options } = payload as {
          readonly query: string;
          readonly options: EditFindOptions;
        };
        return engine().findText(query, options);
      }
      case "edit-dispose":
        state?.dispose();
        state = undefined;
        return undefined;
      default:
        throw new ViewerError(
          "internal",
          `The PDF edit worker does not handle ${operation}`,
        );
    }
  };
}

export { pdfOperationSchemas };
