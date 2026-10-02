import type { DocxOperationHandler } from "./operations.js";
import {
  replaceTextHandler,
  setParagraphStyleHandler,
  setTextStyleHandler,
} from "./text-ops.js";

/** Handlers by operation name; `IMPLEMENTED_OPERATIONS` in schemas.ts lists the same names. */
export const docxHandlers: ReadonlyMap<string, DocxOperationHandler> = new Map<
  string,
  DocxOperationHandler
>([
  ["replaceText", replaceTextHandler as DocxOperationHandler],
  ["setTextStyle", setTextStyleHandler as DocxOperationHandler],
  ["setParagraphStyle", setParagraphStyleHandler as DocxOperationHandler],
]);
