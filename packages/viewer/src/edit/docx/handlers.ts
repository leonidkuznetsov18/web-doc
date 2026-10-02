import type { DocxOperationHandler } from "./operations.js";
import {
  deleteElementHandler,
  insertImageHandler,
  insertParagraphHandler,
  moveElementHandler,
} from "./structure-ops.js";
import { insertTableHandler, setTableCellHandler } from "./table-ops.js";
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
  ["insertParagraph", insertParagraphHandler as DocxOperationHandler],
  ["deleteElement", deleteElementHandler as DocxOperationHandler],
  ["moveElement", moveElementHandler as DocxOperationHandler],
  ["insertImage", insertImageHandler as DocxOperationHandler],
  ["insertTable", insertTableHandler as DocxOperationHandler],
  ["setTableCell", setTableCellHandler as DocxOperationHandler],
]);
