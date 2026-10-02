import type { PptxOperationHandler } from "./operations.js";
import {
  deleteElementHandler,
  insertTextBoxHandler,
  moveElementHandler,
  resizeElementHandler,
  setShapeStyleHandler,
} from "./shape-ops.js";
import {
  insertImageHandler,
  insertTableHandler,
  setTableCellHandler,
} from "./image-table-ops.js";
import {
  deleteSlideHandler,
  duplicateSlideHandler,
  insertSlideHandler,
  moveSlideHandler,
} from "./slide-ops.js";
import { replaceTextHandler, setTextStyleHandler } from "./text-ops.js";

/** Handlers by operation name; `IMPLEMENTED_OPERATIONS` in schemas.ts lists the same names. */
export const pptxHandlers: ReadonlyMap<string, PptxOperationHandler> = new Map<
  string,
  PptxOperationHandler
>([
  ["replaceText", replaceTextHandler as PptxOperationHandler],
  ["setTextStyle", setTextStyleHandler as PptxOperationHandler],
  ["setShapeStyle", setShapeStyleHandler as PptxOperationHandler],
  ["moveElement", moveElementHandler as PptxOperationHandler],
  ["resizeElement", resizeElementHandler as PptxOperationHandler],
  ["deleteElement", deleteElementHandler as PptxOperationHandler],
  ["insertTextBox", insertTextBoxHandler as PptxOperationHandler],
  ["insertImage", insertImageHandler as PptxOperationHandler],
  ["insertTable", insertTableHandler as PptxOperationHandler],
  ["setTableCell", setTableCellHandler as PptxOperationHandler],
  ["insertSlide", insertSlideHandler as PptxOperationHandler],
  ["duplicateSlide", duplicateSlideHandler as PptxOperationHandler],
  ["deleteSlide", deleteSlideHandler as PptxOperationHandler],
  ["moveSlide", moveSlideHandler as PptxOperationHandler],
]);
