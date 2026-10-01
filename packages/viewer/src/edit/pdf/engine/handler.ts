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
import type { Pdfium } from "./pdfium.js";
import { pdfOperationSchemas } from "../schemas.js";
import { PdfEditDocument } from "./document.js";

/**
 * Serves the edit worker protocol for one PDF document. It runs inside the
 * edit worker in the browser and directly in Node tests; `loadPdfium` is how
 * the host obtains an engine instance for a WASM URL.
 */
export function createPdfEditHandler(
  loadPdfium: (wasmUrl: string) => Promise<Pdfium>,
): WorkerOperationHandler {
  let pdfium: Pdfium | undefined;
  let state: PdfEditDocument | undefined;

  const engine = (): PdfEditDocument => {
    if (!state)
      throw new ViewerError("lifecycle-error", "No PDF is open for editing");
    return state;
  };

  return async (operation: WorkerOperation, payload: unknown) => {
    switch (operation) {
      case "edit-init": {
        pdfium ??= await loadPdfium((payload as EditWorkerInitPayload).wasmUrl);
        return undefined;
      }
      case "edit-open": {
        if (!pdfium)
          throw new ViewerError("lifecycle-error", "PDFium is not initialised");
        state?.dispose();
        state = new PdfEditDocument(
          pdfium,
          new Uint8Array((payload as EditWorkerOpenPayload).data),
        );
        const result: EditWorkerOpenResult = { pageCount: state.pageCount };
        return result;
      }
      case "edit-validate":
        return engine().validate(
          (payload as { readonly operations: readonly EditOperation[] })
            .operations,
        );
      case "edit-apply":
        return engine().apply(
          (payload as { readonly operations: readonly EditOperation[] })
            .operations,
        );
      case "edit-materialize":
        return engine().materialize().buffer;
      case "edit-restore":
        engine().restore(
          (
            payload as {
              readonly batches: readonly (readonly EditOperation[])[];
            }
          ).batches,
        );
        return undefined;
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
