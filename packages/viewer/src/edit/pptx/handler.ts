import { ViewerError } from "../../errors.js";
import type { WorkerOperationHandler } from "../../worker-endpoint.js";
import type {
  EditWorkerOpenPayload,
  EditWorkerOpenResult,
  WorkerOperation,
} from "../../worker-protocol.js";
import type { EngineBatch } from "../engine.js";
import type {
  EditFindOptions,
  EditOperation,
  ElementQuery,
  PagePoint,
} from "../types.js";
import { PptxEditEngine } from "./engine.js";

/**
 * Serves the edit worker protocol for one OOXML document. It runs inside
 * the OOXML edit worker in the browser and directly in Node tests. The open
 * payload names the format; PPTX is served now, DOCX joins with module 06.
 */
export function createOoxmlEditHandler(): WorkerOperationHandler {
  let state: PptxEditEngine | undefined;

  const engine = (): PptxEditEngine => {
    if (!state)
      throw new ViewerError(
        "lifecycle-error",
        "No package is open for editing",
      );
    return state;
  };

  return async (operation: WorkerOperation, payload: unknown, context) => {
    const { signal } = context;
    switch (operation) {
      case "edit-init":
        return undefined;
      case "edit-open": {
        const open = payload as EditWorkerOpenPayload;
        if (open.format !== undefined && open.format !== "pptx")
          throw new ViewerError(
            "edit-unsupported",
            `The OOXML edit worker does not serve ${open.format} yet`,
            { details: { format: open.format, reason: "no-engine" } },
          );
        await state?.dispose();
        state = await PptxEditEngine.open(
          new Uint8Array(open.data),
          open.limits,
          signal,
        );
        const result: EditWorkerOpenResult = { pageCount: state.pageCount };
        return result;
      }
      case "edit-validate": {
        const { operations } = payload as {
          readonly operations: readonly EditOperation[];
        };
        return engine().validate(operations, signal);
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
        return engine().slides(signal);
      case "edit-pptx-layouts":
        return engine().layouts(signal);
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
