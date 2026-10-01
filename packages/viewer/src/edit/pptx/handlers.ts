import type { PptxOperationHandler } from "./operations.js";
import { replaceTextHandler, setTextStyleHandler } from "./text-ops.js";

/** Handlers by operation name; `IMPLEMENTED_OPERATIONS` in schemas.ts lists the same names. */
export const pptxHandlers: ReadonlyMap<string, PptxOperationHandler> = new Map<
  string,
  PptxOperationHandler
>([
  ["replaceText", replaceTextHandler as PptxOperationHandler],
  ["setTextStyle", setTextStyleHandler as PptxOperationHandler],
]);
