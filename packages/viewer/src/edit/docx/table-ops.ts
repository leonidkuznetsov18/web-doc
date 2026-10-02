import { patches, type XmlPatch } from "../ooxml/patch.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import { W_NS } from "./ids.js";
import type { DocxModel, ParagraphRecord, TableRecord } from "./model.js";
import type {
  DocxOperationContext,
  DocxOperationHandler,
  DocxOperationResult,
  Issue,
} from "./operations.js";
import { LINE_BREAK } from "./text.js";
import { replacedParagraph } from "./text-ops.js";
import type {
  DocxInsertTableOperation,
  DocxSetTableCellOperation,
} from "./types.js";
import {
  namespacePatches,
  normalizeText,
  paragraphXml,
  runContentXml,
  runXml,
  textProblem,
} from "./write.js";

/*
 * Tables: a new `w:tbl` next to a paragraph or table of the body, with a
 * grid over the section's content width, one paragraph per cell and a
 * paragraph after it where the body needs one; a cell's text replaced in
 * its first paragraph with the cell's other paragraphs removed.
 */

/** US Letter with one-inch margins: the content width Word assumes without a section. */
const DEFAULT_CONTENT_WIDTH = 12240 - 2 * 1440;

/** Exactly one of `before` and `after`, resolved to a body-level block. */
function placement(
  operation: { readonly before?: string; readonly after?: string },
  context: DocxOperationContext,
  issue: Issue,
):
  | { record: ParagraphRecord | TableRecord; side: "before" | "after" }
  | undefined {
  const named = [operation.before, operation.after].filter(
    (value) => value !== undefined,
  );
  if (named.length !== 1) {
    issue(
      "/before",
      "invalid-value",
      "Exactly one of before and after names the reference element",
    );
    return undefined;
  }
  const side = operation.before !== undefined ? "before" : "after";
  const id = operation[side]!;
  const record = context.model.byId.get(id);
  if (!record || (record.kind !== "paragraph" && record.kind !== "table")) {
    issue(
      `/${side}`,
      record ? "invalid-target" : "unknown-target",
      record
        ? `Element ${id} is not a paragraph or a table`
        : `No element ${id}`,
    );
    return undefined;
  }
  if (record.kind === "paragraph" && record.table) {
    issue(
      `/${side}`,
      "invalid-target",
      "A table cannot be placed inside a cell",
    );
    return undefined;
  }
  return { record, side };
}

/** The content width, in twentieths of a point, of the section a block belongs to. */
function contentWidth(
  model: DocxModel,
  node: XmlElement,
  side: "before" | "after",
): number {
  const part = model.document;
  const geometry = (sectPr: XmlElement | undefined): number | undefined => {
    if (!sectPr) return undefined;
    const read = (local: string, name: string): number | undefined => {
      const holder = sectPr.children.find(
        (child) => child.local === local && child.namespace === W_NS,
      );
      const value = Number(holder ? part.attribute(holder, name) : NaN);
      return Number.isFinite(value) && value >= 0 ? value : undefined;
    };
    const width = read("pgSz", "w:w");
    if (width === undefined) return undefined;
    return (
      width -
      (read("pgMar", "w:left") ?? 1440) -
      (read("pgMar", "w:right") ?? 1440)
    );
  };
  // The section a block belongs to ends at the next paragraph carrying a
  // w:sectPr, else at the body's own; a table placed after a paragraph
  // that ends a section lands in the section after it.
  let passed = false;
  for (const block of model.blocks) {
    const own = block.node === node;
    if (own) passed = true;
    if (own && side === "after") continue;
    if (passed && block.kind === "paragraph" && block.sectPr) {
      const width = geometry(block.sectPr);
      if (width !== undefined) return width;
      break;
    }
  }
  return geometry(model.bodySectPr) ?? DEFAULT_CONTENT_WIDTH;
}

/** Column widths in twentieths of a point from relative weights over `total`. */
function columnWidths(weights: readonly number[], total: number): number[] {
  const sum = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((weight) => Math.floor((weight / sum) * total));
  const remainder = total - widths.reduce((a, b) => a + b, 0);
  widths[widths.length - 1]! += remainder;
  return widths;
}

/** `w:tblPr`: the TableGrid style when the document defines it, else single borders. */
function tablePropertiesXml(model: DocxModel): string {
  const grid = model.styles.styles.get("TableGrid");
  const look =
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>';
  if (grid && grid.type === "table")
    return `<w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/>${look}</w:tblPr>`;
  const edge = (name: string): string =>
    `<w:${name} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`;
  return `<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map(edge).join("")}</w:tblBorders>${look}</w:tblPr>`;
}

/** The text of a cell as one paragraph: newlines become line breaks. */
function cellRunXml(text: string): string {
  return runXml(
    "",
    runContentXml(normalizeText(text).replaceAll("\n", LINE_BREAK)),
  );
}

function commit(
  context: DocxOperationContext,
  items: readonly XmlPatch[],
  result: Omit<DocxOperationResult, "warnings">,
): Promise<DocxOperationResult> {
  const part = context.model.document;
  const transaction = context.pkg.transaction();
  transaction.patch(part, [...namespacePatches(part), ...items]);
  return transaction
    .commit()
    .then((change) => ({ ...result, warnings: change.warnings }));
}

export const insertTableHandler: DocxOperationHandler<DocxInsertTableOperation> =
  {
    async validate(operation, context, issue) {
      placement(operation, context, issue);
      const columns = operation.rows[0]?.length ?? 0;
      if (operation.rows.some((row) => row.length !== columns))
        issue(
          "/rows",
          "invalid-value",
          "Every row needs the same number of cells",
        );
      if (operation.columnWidths && operation.columnWidths.length !== columns)
        issue(
          "/columnWidths",
          "invalid-value",
          "One weight per column is needed",
        );
      operation.rows.forEach((row, rowIndex) =>
        row.forEach((cell, columnIndex) => {
          const problem = textProblem(cell);
          if (problem)
            issue(
              `/rows/${rowIndex}/${columnIndex}`,
              "invalid-text",
              `The text holds ${problem}`,
            );
        }),
      );
    },
    async apply(operation, context) {
      const { model } = context;
      const { record, side } = placement(operation, context, () => {})!;
      const part = model.document;
      const node = record.node;
      const widths = columnWidths(
        operation.columnWidths ?? operation.rows[0]!.map(() => 1),
        contentWidth(model, node, side),
      );
      const createdIds: string[] = [];
      const cellIds: string[] = [];
      const rows = operation.rows
        .map(
          (row) =>
            `<w:tr>${row
              .map((cell, column) => {
                const id = context.freshParagraphId();
                cellIds.push(`p:${id}`);
                return `<w:tc><w:tcPr><w:tcW w:w="${widths[column]}" w:type="dxa"/></w:tcPr>${paragraphXml(undefined, id, "", cellRunXml(cell))}</w:tc>`;
              })
              .join("")}</w:tr>`,
        )
        .join("");
      const table = `<w:tbl>${tablePropertiesXml(model)}<w:tblGrid>${widths
        .map((width) => `<w:gridCol w:w="${width}"/>`)
        .join("")}</w:tblGrid>${rows}</w:tbl>`;
      createdIds.push(`tbl:${cellIds[0]!.slice(2)}`, ...cellIds);
      // A paragraph must follow the table when the next sibling would be a
      // table or the body's section properties.
      const siblings = model.blocks;
      const index = siblings.indexOf(record);
      const next = side === "before" ? record : siblings[index + 1];
      let trailing = "";
      if (!next || next.kind === "table") {
        const id = context.freshParagraphId();
        createdIds.push(`p:${id}`);
        trailing = paragraphXml(undefined, id, "", "");
      }
      return commit(
        context,
        [
          side === "before"
            ? patches.insertBefore(part, node, table)
            : patches.insertAfter(part, node, table),
          ...(trailing
            ? [
                side === "before"
                  ? patches.insertBefore(part, node, trailing)
                  : patches.insertAfter(part, node, trailing),
              ]
            : []),
        ],
        { createdIds, reflowFrom: record.id },
      );
    },
  };

export const setTableCellHandler: DocxOperationHandler<DocxSetTableCellOperation> =
  {
    async validate(operation, context, issue) {
      const problem = textProblem(normalizeText(operation.text));
      if (problem) issue("/text", "invalid-text", `The text holds ${problem}`);
      const record = context.model.byId.get(operation.target);
      if (!record || record.kind !== "table") {
        issue(
          "/target",
          record ? "invalid-target" : "unknown-target",
          record
            ? `Element ${operation.target} is not a table`
            : `No element ${operation.target}`,
        );
        return;
      }
      const row = record.rows[operation.row];
      if (!row) {
        issue("/row", "range", `The table has ${record.rows.length} rows`);
        return;
      }
      const cell = row[operation.column];
      if (!cell) {
        issue(
          "/column",
          "range",
          `Row ${operation.row} has ${row.length} cells`,
        );
        return;
      }
      const first = cell.paragraphs[0];
      if (!first)
        issue("/target", "invalid-target", "The cell holds no paragraph");
      else if (first.readOnlyReason)
        issue(
          "/target",
          "invalid-target",
          `The cell's paragraph is read-only (${first.readOnlyReason})`,
        );
    },
    async apply(operation, context) {
      const { model } = context;
      const record = model.byId.get(operation.target) as TableRecord;
      const cell = record.rows[operation.row]![operation.column]!;
      const [first, ...rest] = cell.paragraphs;
      const text = normalizeText(operation.text).replaceAll("\n", LINE_BREAK);
      const replaced = replacedParagraph(
        context,
        first!,
        0,
        first!.text.text.length,
        text,
      );
      const part: XmlPart = model.document;
      const items: XmlPatch[] = [...replaced.items];
      const removedIds = [...replaced.removedIds];
      const removedParagraphIds: string[] = [...replaced.removedParagraphIds];
      for (const paragraph of rest) {
        items.push(patches.removeElement(part, paragraph.node));
        removedIds.push(
          paragraph.elementId,
          ...paragraph.inlines.map((inline) => inline.elementId),
        );
        removedParagraphIds.push(paragraph.id);
        for (const nested of part.findAll("p", paragraph.node)) {
          const id = model.paragraphIds.get(nested);
          if (id && nested !== paragraph.node) removedParagraphIds.push(id);
        }
      }
      return commit(context, items, {
        createdIds: replaced.createdIds,
        ...(removedIds.length > 0 ? { removedIds } : {}),
        ...(removedParagraphIds.length > 0 ? { removedParagraphIds } : {}),
        ...(model.unauthoredSet.has(first!.id) ? { stamped: [first!.id] } : {}),
        reflowFrom: record.id,
      });
    },
  };
