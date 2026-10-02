import { ViewerError } from "../../errors.js";
import type { WorkerOperationHandler } from "../../worker-endpoint.js";
import type {
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
  WorkerOperation,
} from "../../worker-protocol.js";
import { DocxEditEngine } from "../docx/engine.js";
import type { BatchMode, EngineBatch } from "../engine.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  PagePoint,
} from "../types.js";
import { PptxEditEngine } from "./engine.js";

type OoxmlEngine = PptxEditEngine | DocxEditEngine;

/**
 * Serves the edit worker protocol for one OOXML document. It runs inside
 * the OOXML edit worker in the browser and directly in Node tests. The open
 * payload names the format: PPTX (the default) or DOCX.
 */
export function createOoxmlEditHandler(): WorkerOperationHandler {
  let state: OoxmlEngine | undefined;

  const engine = (): OoxmlEngine => {
    if (!state)
      throw new ViewerError(
        "lifecycle-error",
        "No package is open for editing",
      );
    return state;
  };

  const docx = (): DocxEditEngine => {
    const current = engine();
    if (!(current instanceof DocxEditEngine))
      throw new ViewerError(
        "edit-unsupported",
        "The open document is not a Word document",
        { details: { format: "pptx", reason: "no-reads" } },
      );
    return current;
  };

  const pptx = (): PptxEditEngine => {
    const current = engine();
    if (!(current instanceof PptxEditEngine))
      throw new ViewerError(
        "edit-unsupported",
        "The open document is not a presentation",
        { details: { format: "docx", reason: "no-reads" } },
      );
    return current;
  };

  return async (operation: WorkerOperation, payload: unknown, context) => {
    const { signal } = context;
    switch (operation) {
      case "edit-init":
        return undefined;
      case "edit-open": {
        const open = payload as EditWorkerOpenPayload;
        const format = open.format ?? "pptx";
        if (format !== "pptx" && format !== "docx")
          throw new ViewerError(
            "edit-unsupported",
            `The OOXML edit worker does not serve ${open.format} yet`,
            { details: { format: open.format, reason: "no-engine" } },
          );
        // The live engine stays until the next one opened, so a failed or
        // aborted open leaves the worker serving what it served before.
        const bytes = new Uint8Array(open.data);
        const next: OoxmlEngine =
          format === "docx"
            ? await DocxEditEngine.open(bytes, open.limits, signal)
            : await PptxEditEngine.open(bytes, open.limits, signal);
        await state?.dispose();
        state = next;
        const result: EditWorkerOpenResult = { pageCount: state.pageCount };
        return result;
      }
      case "edit-validate": {
        const { operations, mode } = payload as {
          readonly operations: readonly EditOperation[];
          readonly mode?: BatchMode;
        };
        return engine().validate(operations, signal, mode);
      }
      case "edit-apply": {
        const { batch } = payload as { readonly batch: EngineBatch };
        return engine().apply(batch, signal);
      }
      case "edit-put-asset": {
        const { id, data } = payload as {
          readonly id: string;
          readonly data: ArrayBuffer;
        };
        await engine().putAsset(id, new Uint8Array(data), signal);
        return undefined;
      }
      case "edit-materialize": {
        const { purpose, options } = (payload ?? {}) as {
          readonly purpose?: "show" | "save";
          readonly options?: Record<string, unknown>;
        };
        const { bytes, warnings } = await engine().materializeDocument(
          purpose ?? "show",
          options ?? {},
          signal,
        );
        const data = bytes.slice().buffer;
        return { data, warnings };
      }
      case "edit-restore": {
        const { batches, base } = payload as {
          readonly batches: readonly EngineBatch[];
          readonly base?: ArrayBuffer;
        };
        await engine().restore(
          { batches, ...(base ? { base: new Uint8Array(base) } : {}) },
          signal,
        );
        return undefined;
      }
      case "edit-elements":
        return engine().getElements(
          (payload as { readonly query: ElementQuery }).query,
          signal,
        );
      case "edit-element":
        return engine().getElement(
          (payload as { readonly id: string }).id,
          signal,
        );
      case "edit-elements-at": {
        const { pageIndex, point } = payload as {
          readonly pageIndex: number;
          readonly point: PagePoint;
        };
        return engine().elementsAt(pageIndex, point, signal);
      }
      case "edit-find-text": {
        const { query, options } = payload as {
          readonly query: string;
          readonly options: EditFindOptions;
        };
        return engine().findText(query, options, signal);
      }
      case "edit-pptx-slides":
        return pptx().slides(signal);
      case "edit-pptx-layouts":
        return pptx().layouts(signal);
      case "edit-docx-revisions":
        return docx().revisions(
          (payload as { readonly id: string }).id,
          signal,
        );
      case "edit-dispose":
        await state?.dispose();
        state = undefined;
        return undefined;
      default:
        throw new ViewerError(
          "internal",
          `The OOXML edit worker does not handle ${operation}`,
        );
    }
  };
}
