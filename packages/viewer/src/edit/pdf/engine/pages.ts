import type {
  DeletePageOperation,
  InsertPageOperation,
  MovePageOperation,
  RotatePageOperation,
} from "../types.js";
import type { OperationContext, OperationHandler } from "./operations.js";

/*
 * Page structure. The model's page records move with the pages, so element
 * ids keep pointing at the same content whatever happens to page order.
 */

export const insertPage: OperationHandler<InsertPageOperation> = {
  validate(operation, context, issue) {
    if (operation.index > context.pageCount)
      issue(
        "/index",
        "range",
        `The index must be at most the page count (${context.pageCount})`,
      );
  },
  apply(operation, context) {
    const { lib } = context.pdfium;
    const size = operation.size ?? neighbourSize(operation.index, context);
    const page = lib.FPDFPage_New(
      context.document,
      operation.index,
      size.width,
      size.height,
    );
    lib.FPDFPage_GenerateContent(page);
    lib.FPDF_ClosePage(page);
    context.insertPageRecord(operation.index);
    return {
      createdIds: [],
      changedPages: fromIndex(operation.index, context.pageCount + 1),
      warnings: [],
    };
  },
};

export const deletePage: OperationHandler<DeletePageOperation> = {
  validate(operation, context, issue) {
    if (operation.pageIndex >= context.pageCount)
      issue("/pageIndex", "unknown-target", `No page ${operation.pageIndex}`);
    else if (context.pageCount === 1)
      issue("/pageIndex", "last-page", "The last page cannot be deleted");
  },
  apply(operation, context) {
    context.pdfium.lib.FPDFPage_Delete(context.document, operation.pageIndex);
    context.removePageRecord(operation.pageIndex);
    return {
      createdIds: [],
      changedPages: fromIndex(operation.pageIndex, context.pageCount - 1),
      warnings: [],
    };
  },
};

export const movePage: OperationHandler<MovePageOperation> = {
  validate(operation, context, issue) {
    if (operation.from >= context.pageCount)
      issue("/from", "unknown-target", `No page ${operation.from}`);
    if (operation.to >= context.pageCount)
      issue(
        "/to",
        "range",
        `The index must be below the page count (${context.pageCount})`,
      );
  },
  apply(operation, context) {
    const { pdfium } = context;
    if (operation.from !== operation.to) {
      const indexes = pdfium.writeInt32Array([operation.from]);
      try {
        // PDFium takes the page out first, so the destination is the index
        // the page will have afterwards.
        pdfium.lib.FPDF_MovePages(context.document, indexes, 1, operation.to);
      } finally {
        pdfium.free(indexes);
      }
      context.movePageRecord(operation.from, operation.to);
    }
    const low = Math.min(operation.from, operation.to);
    const high = Math.max(operation.from, operation.to);
    return {
      createdIds: [],
      changedPages: Array.from({ length: high - low + 1 }, (_, i) => low + i),
      warnings: [],
    };
  },
};

export const rotatePage: OperationHandler<RotatePageOperation> = {
  validate(operation, context, issue) {
    if (operation.pageIndex >= context.pageCount)
      issue("/pageIndex", "unknown-target", `No page ${operation.pageIndex}`);
  },
  apply(operation, context) {
    const { lib } = context.pdfium;
    context.withPage(operation.pageIndex, (page) => {
      lib.FPDFPage_SetRotation(page, operation.rotation / 90);
    });
    context.invalidatePage(operation.pageIndex);
    return {
      createdIds: [],
      changedPages: [operation.pageIndex],
      warnings: [],
    };
  },
};

function neighbourSize(
  index: number,
  context: OperationContext,
): { readonly width: number; readonly height: number } {
  if (context.pageCount === 0) return { width: 612, height: 792 };
  return context.pageSize(index > 0 ? index - 1 : 0);
}

function fromIndex(start: number, pageCount: number): number[] {
  return Array.from(
    { length: Math.max(0, pageCount - start) },
    (_, i) => start + i,
  );
}
