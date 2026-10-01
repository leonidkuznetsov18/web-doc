import { ViewerError } from "../../../errors.js";
import type { WorkerOperationHandler } from "../../../worker-endpoint.js";
import type {
  EditWorkerInitPayload,
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
  WorkerOperation,
} from "../../../worker-protocol.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  PagePoint,
} from "../../types.js";
import { pdfOperationSchemas } from "../schemas.js";
import { PdfEditDocument } from "./document.js";
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
        );
        const result: EditWorkerOpenResult = { pageCount: state.pageCount };
        return result;
      }
      case "edit-validate": {
        const { operations } = payload as {
          readonly operations: readonly EditOperation[];
        };
        await fonts.prepare(engine().fontRequests(operations));
        await engine().images.prepare(operations, host.decodeImage);
        return engine().validate(operations);
      }
      case "edit-apply": {
        const { operations } = payload as {
          readonly operations: readonly EditOperation[];
        };
        await fonts.prepare(engine().fontRequests(operations));
        await engine().images.prepare(operations, host.decodeImage);
        return engine().apply(operations);
      }
      case "edit-materialize":
        return engine().materialize().buffer;
      case "edit-restore": {
        const { batches } = payload as {
          readonly batches: readonly (readonly EditOperation[])[];
        };
        await fonts.prepare(engine().fontRequests(batches.flat()));
        await engine().images.prepare(batches.flat(), host.decodeImage);
        engine().restore(batches);
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
