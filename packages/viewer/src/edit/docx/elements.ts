import type { PageRect } from "../types.js";
import {
  tableRows,
  tableText,
  type AnyRecord,
  type DocxModel,
  type ParagraphRecord,
} from "./model.js";
import { IMPLEMENTED_OPERATIONS } from "./schemas.js";
import { resolveParagraphStyle, resolveTextStyle } from "./style.js";
import { firstTextItem } from "./text.js";
import type { DocxElement, DocxElementKind } from "./types.js";

/*
 * `DocxElement` values from the model's records, without geometry: the
 * engine never lays out, so `pageIndex` is −1 and `bounds` empty until
 * the session joins the renderer's runs on the main thread.
 */

export const NO_PAGE = -1;
export const EMPTY_RECT: PageRect = { x: 0, y: 0, width: 0, height: 0 };

/** The operations each kind accepts once its handler ships. */
const OPERATIONS: Readonly<Record<DocxElementKind, readonly string[]>> = {
  paragraph: [
    "replaceText",
    "setTextStyle",
    "setParagraphStyle",
    "insertParagraph",
    "insertTable",
    "insertImage",
    "moveElement",
    "deleteElement",
  ],
  table: [
    "setTableCell",
    "insertParagraph",
    "insertTable",
    "insertImage",
    "moveElement",
    "deleteElement",
  ],
  image: ["deleteElement"],
  other: [],
};
/** A read-only paragraph only accepts siblings placed next to it. */
const INSERTIONS = ["insertParagraph", "insertTable", "insertImage"];

function operationsOf(
  kind: DocxElementKind,
  readOnly: boolean,
): readonly string[] {
  return OPERATIONS[kind].filter(
    (name) =>
      IMPLEMENTED_OPERATIONS.includes(name) &&
      (!readOnly || INSERTIONS.includes(name)),
  );
}

export function toElement(model: DocxModel, record: AnyRecord): DocxElement {
  const base = {
    pageIndex: NO_PAGE,
    bounds: EMPTY_RECT,
    fragments: [],
    story: { kind: "body" as const },
  };
  if (record.kind === "paragraph") return paragraphElement(model, record, base);
  if (record.kind === "table")
    return {
      ...base,
      id: record.elementId,
      kind: "table",
      text: tableText(record),
      table: { rows: tableRows(record) },
      operations: operationsOf("table", false),
    };
  return {
    ...base,
    id: record.elementId,
    kind: record.kind,
    parentId: record.paragraph.elementId,
    operations: operationsOf(record.kind, false),
  };
}

function paragraphElement(
  model: DocxModel,
  record: ParagraphRecord,
  base: Pick<DocxElement, "pageIndex" | "bounds" | "fragments" | "story">,
): DocxElement {
  const first = firstTextItem(record.text);
  return {
    ...base,
    id: record.elementId,
    kind: "paragraph",
    text: record.text.text,
    ...(record.table ? { parentId: record.table.elementId } : {}),
    textStyle: resolveTextStyle(model.styles, record.pPr, first?.rPr),
    paragraphStyle: resolveParagraphStyle(model.styles, record.pPr),
    ...(record.readOnlyReason ? { readOnlyReason: record.readOnlyReason } : {}),
    operations: operationsOf("paragraph", record.readOnlyReason !== undefined),
  };
}
