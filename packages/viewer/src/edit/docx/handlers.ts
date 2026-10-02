import type { DocxOperationHandler } from "./operations.js";

/** Handlers by operation name; `IMPLEMENTED_OPERATIONS` in schemas.ts lists the same names. */
export const docxHandlers: ReadonlyMap<string, DocxOperationHandler> = new Map<
  string,
  DocxOperationHandler
>([]);
