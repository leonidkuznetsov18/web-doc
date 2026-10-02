import { resolveBinary } from "../assets.js";
import { extensionForMime, IMAGE_RELATIONSHIP_TYPE } from "../ooxml/opc.js";
import { patches, type XmlPatch } from "../ooxml/patch.js";
import { relativeTarget } from "../ooxml/transaction.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import { OFFICE_RELATIONSHIPS, W_NS } from "./ids.js";
import type {
  DocxModel,
  InlineRecord,
  ParagraphRecord,
  TableRecord,
} from "./model.js";
import type {
  DocxOperationContext,
  DocxOperationHandler,
  DocxOperationResult,
  Issue,
} from "./operations.js";
import { firstTextItem } from "./text.js";
import type {
  DocxDeleteElementOperation,
  DocxInsertImageOperation,
  DocxInsertParagraphOperation,
  DocxMoveElementOperation,
} from "./types.js";
import {
  changedRunProperties,
  colorProblem,
  isAuthored,
  namespacePatches,
  normalizeText,
  paragraphPropertiesWithoutSection,
  paragraphXml,
  runContentXml,
  runXml,
  sliceOf,
  textProblem,
} from "./write.js";

/*
 * Structure: paragraphs inserted next to other elements, paragraphs,
 * tables and inline pictures deleted, paragraphs and tables moved within
 * their container, inline pictures inserted as a paragraph of their own.
 * Every operation is a patch of the body part; a moved or new paragraph
 * carries its id, so ids survive the move.
 */

const MEDIA_FOLDER = "/word/media";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const R_NS = OFFICE_RELATIONSHIPS.slice(0, -1);
/** EMU per point. */
const EMU_PER_PT = 12_700;

/** The node a block occupies among its siblings: the paragraph or table itself. */
function blockNode(record: ParagraphRecord | TableRecord): XmlElement {
  return record.node;
}

/** Exactly one of `before` and `after`, resolved to its record. */
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
  return { record, side };
}

/** The `w:rPr` bytes new text next to a reference takes: its first text run's, else its paragraph mark's. */
function referenceRunProperties(
  part: XmlPart,
  record: ParagraphRecord | TableRecord,
): XmlElement | undefined {
  const paragraph =
    record.kind === "paragraph" ? record : record.rows[0]?.[0]?.paragraphs[0];
  if (!paragraph) return undefined;
  const first = firstTextItem(paragraph.text);
  if (first?.rPr) return first.rPr;
  return paragraph.pPr?.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
}

/** Ids of every `w:p` under a node, the node included when it is one. */
function paragraphIdsUnder(model: DocxModel, node: XmlElement): string[] {
  const out: string[] = [];
  if (node.local === "p" && node.namespace === W_NS) {
    const own = model.paragraphIds.get(node);
    if (own) out.push(own);
  }
  for (const paragraph of model.document.findAll("p", node)) {
    if (paragraph === node || paragraph.namespace !== W_NS) continue;
    const id = model.paragraphIds.get(paragraph);
    if (id) out.push(id);
  }
  return out;
}

/** Element ids of every listed element under a node, for `removedIds`. */
function elementIdsUnder(model: DocxModel, node: XmlElement): string[] {
  const out: string[] = [];
  for (const record of model.records) {
    let current: XmlElement | undefined = record.node;
    while (current && current !== node) current = current.parent;
    if (current === node) out.push(record.elementId);
  }
  return out;
}

/**
 * The bytes of a node with a `w14:paraId` written on every paragraph in
 * it that lacks one, the ids being those the engine knows them by, so a
 * copy placed elsewhere keeps the same ids.
 */
function stampedSlice(
  model: DocxModel,
  node: XmlElement,
): { xml: string; stamped: string[] } {
  const part = model.document;
  const targets = [
    ...(node.local === "p" && node.namespace === W_NS ? [node] : []),
    ...part
      .findAll("p", node)
      .filter((p) => p !== node && p.namespace === W_NS),
  ]
    .filter((paragraph) => !isAuthored(part, paragraph))
    .sort((a, b) => b.start - a.start);
  let xml = sliceOf(part, node);
  const stamped: string[] = [];
  for (const paragraph of targets) {
    const id = model.paragraphIds.get(paragraph);
    if (!id) continue;
    // After the attributes of the start tag, before ">" or "/>".
    const last = paragraph.attributes.at(-1);
    const from = last ? last.end : paragraph.start + 1 + paragraph.name.length;
    const close = part.text.indexOf(">", from);
    let at = close;
    if (part.text[close - 1] === "/") at = close - 1;
    const offset = at - node.start;
    xml = `${xml.slice(0, offset)} w14:paraId="${id}"${xml.slice(offset)}`;
    stamped.push(id);
  }
  return { xml, stamped: stamped.reverse() };
}

/** The table a cell paragraph names and whether it is the table's first paragraph. */
function firstCellParagraph(record: ParagraphRecord): TableRecord | undefined {
  const table = record.table;
  return table && table.rows[0]?.[0]?.paragraphs[0] === record
    ? table
    : undefined;
}

/**
 * A table is named after its first paragraph; when that paragraph
 * changes, the table's id changes with it and the receipt says so.
 */
function tableRemap(
  table: TableRecord,
  newFirstParagraphId: string,
): Readonly<Record<string, string>> | undefined {
  return table.id === newFirstParagraphId
    ? undefined
    : { [table.elementId]: `tbl:${newFirstParagraphId}` };
}

function commit(
  context: DocxOperationContext,
  items: readonly XmlPatch[],
  result: Omit<DocxOperationResult, "warnings">,
  transaction = context.pkg.transaction(),
): Promise<DocxOperationResult> {
  const part = context.model.document;
  transaction.patch(part, [...namespacePatches(part), ...items]);
  return transaction
    .commit()
    .then((change) => ({ ...result, warnings: change.warnings }));
}

export const insertParagraphHandler: DocxOperationHandler<DocxInsertParagraphOperation> =
  {
    async validate(operation, context, issue) {
      const problem = textProblem(normalizeText(operation.text));
      if (problem) issue("/text", "invalid-text", `The text holds ${problem}`);
      if (operation.style?.color !== undefined) {
        const colour = colorProblem(operation.style.color);
        if (colour) issue("/style/color", "invalid-value", `Colour: ${colour}`);
      }
      if (operation.style?.fontFamily !== undefined) {
        const font = textProblem(operation.style.fontFamily);
        if (font) issue("/style/fontFamily", "invalid-value", `Font: ${font}`);
      }
      placement(operation, context, issue);
    },
    async apply(operation, context) {
      const { record, side } = placement(operation, context, () => {})!;
      const part = context.model.document;
      const node = blockNode(record);
      const sourceRPr = referenceRunProperties(part, record);
      const rPr = operation.style
        ? changedRunProperties(
            part,
            sourceRPr,
            operation.style,
            context.model.styles,
          )
        : sliceOf(part, sourceRPr);
      const pPr =
        record.kind === "paragraph"
          ? paragraphPropertiesWithoutSection(part, record.pPr)
          : "";
      const segments = normalizeText(operation.text).split("\n");
      const createdIds: string[] = [];
      const xml = segments.map((segment) => {
        const id = context.freshParagraphId();
        createdIds.push(`p:${id}`);
        return paragraphXml(
          undefined,
          id,
          pPr,
          runXml(rPr, runContentXml(segment)),
        );
      });
      // Inserts at one position land in call order, so the paragraphs
      // keep their order on either side of the reference.
      const items =
        side === "before"
          ? xml.map((paragraph) => patches.insertBefore(part, node, paragraph))
          : xml.map((paragraph) => patches.insertAfter(part, node, paragraph));
      const table =
        side === "before" && record.kind === "paragraph"
          ? firstCellParagraph(record)
          : undefined;
      const remappedIds = table
        ? tableRemap(table, createdIds[0]!.slice(2))
        : undefined;
      return commit(context, items, {
        createdIds,
        reflowFrom: record.id,
        ...(remappedIds ? { remappedIds } : {}),
      });
    },
  };

/** Whether the container of a block would keep a paragraph without it. */
function otherParagraphsIn(model: DocxModel, record: ParagraphRecord): boolean {
  if (record.table) {
    for (const row of record.table.rows)
      for (const cell of row)
        if (cell.paragraphs.includes(record))
          return cell.paragraphs.some((candidate) => candidate !== record);
    return false;
  }
  return model.blocks.some(
    (block) => block.kind === "paragraph" && block !== record,
  );
}

export const deleteElementHandler: DocxOperationHandler<DocxDeleteElementOperation> =
  {
    async validate(operation, context, issue) {
      const record = context.model.byId.get(operation.target);
      if (!record || record.kind === "other") {
        issue(
          "/target",
          record ? "invalid-target" : "unknown-target",
          record
            ? `Element ${operation.target} cannot be deleted`
            : `No element ${operation.target}`,
        );
        return;
      }
      if (record.kind === "paragraph") {
        if (record.sectPr)
          issue(
            "/target",
            "section-break",
            "A paragraph that ends a section cannot be deleted",
          );
        else if (record.readOnlyReason)
          issue(
            "/target",
            "invalid-target",
            `Paragraph ${operation.target} is read-only (${record.readOnlyReason})`,
          );
        else if (!otherParagraphsIn(context.model, record))
          issue(
            "/target",
            "last-paragraph",
            "The last paragraph of a body or a cell cannot be deleted",
          );
      }
    },
    async apply(operation, context) {
      const { model } = context;
      const found = model.byId.get(operation.target)!;
      const part = model.document;
      if (found.kind !== "paragraph" && found.kind !== "table")
        return deleteInline(context, found);
      const record: ParagraphRecord | TableRecord = found;
      const node = blockNode(record);
      const items: XmlPatch[] = [patches.removeElement(part, node)];
      const createdIds: string[] = [];
      if (record.kind === "table" && model.blocks.includes(record)) {
        // The body must not end with a table: a paragraph follows if it would.
        const blocks = model.blocks;
        const index = blocks.indexOf(record);
        const next = blocks[index + 1];
        const previous = blocks[index - 1];
        if (!next && previous?.kind === "table") {
          const id = context.freshParagraphId();
          createdIds.push(`p:${id}`);
          items.push(
            patches.insertAfter(
              part,
              node,
              paragraphXml(undefined, id, "", ""),
            ),
          );
        }
      }
      const removedParagraphIds = paragraphIdsUnder(model, node);
      const table =
        record.kind === "paragraph" ? firstCellParagraph(record) : undefined;
      const successor = table?.rows[0]?.[0]?.paragraphs[1];
      const remappedIds =
        table && successor ? tableRemap(table, successor.id) : undefined;
      return commit(context, items, {
        createdIds,
        removedIds: elementIdsUnder(model, node),
        removedParagraphIds,
        reflowFrom: record.id,
        ...(remappedIds ? { remappedIds } : {}),
      });
    },
  };

/** Removes an inline object's drawing and the relationship only it used. */
async function deleteInline(
  context: DocxOperationContext,
  record: InlineRecord,
): Promise<DocxOperationResult> {
  const { model } = context;
  const part = model.document;
  const transaction = context.pkg.transaction();
  const blip = part.find("blip", record.node);
  const rId = blip
    ? (part.attribute(blip, "r:embed") ?? part.attribute(blip, "r:link"))
    : undefined;
  if (rId) {
    const uses = part
      .findAll("blip")
      .filter(
        (candidate) =>
          candidate !== blip &&
          (part.attribute(candidate, "r:embed") === rId ||
            part.attribute(candidate, "r:link") === rId),
      );
    const elsewhere = new RegExp(`r:(?:id|embed|link|pict)="${rId}"`);
    const others =
      uses.length > 0 || elsewhere.test(stripNode(part, record.node));
    if (!others) transaction.removeRelationship(model.mainPart, rId);
  }
  return commit(
    context,
    [patches.removeElement(part, record.node)],
    {
      createdIds: [],
      removedIds: [record.elementId],
      reflowFrom: record.paragraph.id,
    },
    transaction,
  );
}

/** The part's text without a node, to look for other uses of a relationship. */
function stripNode(part: XmlPart, node: XmlElement): string {
  return part.text.slice(0, node.start) + part.text.slice(node.end);
}

export const moveElementHandler: DocxOperationHandler<DocxMoveElementOperation> =
  {
    async validate(operation, context, issue) {
      const record = context.model.byId.get(operation.target);
      if (!record || (record.kind !== "paragraph" && record.kind !== "table")) {
        issue(
          "/target",
          record ? "invalid-target" : "unknown-target",
          record
            ? `Element ${operation.target} is not a paragraph or a table`
            : `No element ${operation.target}`,
        );
        return;
      }
      if (record.kind === "paragraph" && record.sectPr) {
        issue(
          "/target",
          "section-break",
          "A paragraph that ends a section cannot move",
        );
        return;
      }
      const reference = placement(operation, context, issue);
      if (!reference) return;
      if (reference.record === record) {
        issue(
          `/${reference.side}`,
          "invalid-target",
          "An element cannot move next to itself",
        );
        return;
      }
      if (blockNode(reference.record).parent !== blockNode(record).parent)
        issue(
          `/${reference.side}`,
          "invalid-target",
          "The reference must be in the same body or cell as the element",
        );
    },
    async apply(operation, context) {
      const { model } = context;
      const record = model.byId.get(operation.target) as
        ParagraphRecord | TableRecord;
      const { record: reference, side } = placement(
        operation,
        context,
        () => {},
      )!;
      const part = model.document;
      const node = blockNode(record);
      const target = blockNode(reference);
      const { xml, stamped } = stampedSlice(model, node);
      const items: XmlPatch[] = [
        patches.removeElement(part, node),
        side === "before"
          ? patches.insertBefore(part, target, xml)
          : patches.insertAfter(part, target, xml),
      ];
      // The document reflows from whichever of the two comes first.
      const order = (candidate: ParagraphRecord | TableRecord): number =>
        model.records.indexOf(candidate);
      const first = order(record) < order(reference) ? record : reference;
      // Within a table's first cell, the move may change which paragraph
      // comes first, and so the table's id.
      let remappedIds: Readonly<Record<string, string>> | undefined;
      if (record.kind === "paragraph" && record.table) {
        const cell = record.table.rows[0]?.[0];
        if (cell && cell.paragraphs.includes(record)) {
          const without = cell.paragraphs.filter((p) => p !== record);
          const at = without.indexOf(reference as ParagraphRecord);
          const reordered = [...without];
          if (at >= 0)
            reordered.splice(side === "before" ? at : at + 1, 0, record);
          const head = reordered[0];
          if (head) remappedIds = tableRemap(record.table, head.id);
        }
      }
      return commit(context, items, {
        createdIds: [],
        stamped,
        reflowFrom: first.id,
        ...(remappedIds ? { remappedIds } : {}),
      });
    },
  };

/** A media part of the package with these bytes and this type's extension, if any. */
async function existingMedia(
  context: DocxOperationContext,
  bytes: Uint8Array,
  mimeType: string,
): Promise<string | undefined> {
  const extension = `.${extensionForMime(mimeType)}`;
  for (const name of context.pkg.currentPartNames) {
    if (!name.startsWith(`${MEDIA_FOLDER}/`) || !name.endsWith(extension))
      continue;
    const entry = context.pkg.entry(name);
    if (entry && entry.uncompressedSize !== bytes.length) continue;
    const candidate = await context.pkg.part(name);
    if (candidate.length !== bytes.length) continue;
    let same = true;
    for (let index = 0; index < bytes.length; index += 1)
      if (candidate[index] !== bytes[index]) {
        same = false;
        break;
      }
    if (same) return name;
  }
  return undefined;
}

/** Whether the bytes carry the signature of the declared image type. */
function matchesMime(bytes: Uint8Array, mimeType: string): boolean {
  if (mimeType === "image/png")
    return (
      bytes.length >= 8 &&
      bytes[0] === 0x89 &&
      bytes[1] === 0x50 &&
      bytes[2] === 0x4e &&
      bytes[3] === 0x47
    );
  return (
    bytes.length >= 3 &&
    bytes[0] === 0xff &&
    bytes[1] === 0xd8 &&
    bytes[2] === 0xff
  );
}

/** One above every `wp:docPr/@id` of the part. */
function nextDrawingId(part: XmlPart): number {
  let max = 0;
  for (const node of part.findAll("docPr")) {
    if (node.namespace !== WP_NS) continue;
    const id = Number(part.attribute(node, "id") ?? NaN);
    if (Number.isInteger(id)) max = Math.max(max, id);
  }
  return max + 1;
}

/** Patches that declare the namespaces an inline picture's markup uses on the root. */
function drawingNamespacePatches(part: XmlPart): XmlPatch[] {
  const root = part.root;
  const has = (name: string): boolean =>
    root.attributes.some((attribute) => attribute.name === name);
  const items: XmlPatch[] = [];
  if (!has("xmlns:wp"))
    items.push(patches.setAttribute(part, root, "xmlns:wp", WP_NS));
  if (!has("xmlns:r"))
    items.push(patches.setAttribute(part, root, "xmlns:r", R_NS));
  return items;
}

export const insertImageHandler: DocxOperationHandler<DocxInsertImageOperation> =
  {
    async validate(operation, context, issue) {
      placement(operation, context, issue);
      try {
        const bytes = resolveBinary(operation.data, context.assets);
        if (bytes.length === 0)
          issue("/data", "invalid-value", "The image has no bytes");
        else if (!matchesMime(bytes, operation.mimeType))
          issue(
            "/data",
            "invalid-value",
            `The bytes are not a ${operation.mimeType} image`,
          );
      } catch {
        issue(
          "/data",
          "unknown-asset",
          "The asset is not known to the session",
        );
      }
    },
    async apply(operation, context) {
      const { model } = context;
      const { record, side } = placement(operation, context, () => {})!;
      const part = model.document;
      const bytes = resolveBinary(operation.data, context.assets);
      const transaction = context.pkg.transaction();
      // The same bytes already in the package are related, not stored again.
      const existing = await existingMedia(context, bytes, operation.mimeType);
      const rId = existing
        ? await transaction.addRelationship(
            model.mainPart,
            IMAGE_RELATIONSHIP_TYPE,
            relativeTarget(model.mainPart, existing),
          )
        : (
            await transaction.addMedia(
              model.mainPart,
              MEDIA_FOLDER,
              bytes,
              operation.mimeType,
            )
          ).rId;
      const cx = Math.round(operation.size.width * EMU_PER_PT);
      const cy = Math.round(operation.size.height * EMU_PER_PT);
      const drawingId = nextDrawingId(part);
      const id = context.freshParagraphId();
      const drawing =
        `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${drawingId}" name="Picture ${drawingId}"/>` +
        `<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="${A_NS}" noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
        `<a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${PIC_NS}"><pic:pic xmlns:pic="${PIC_NS}"><pic:nvPicPr><pic:cNvPr id="0" name="Picture ${drawingId}"/><pic:cNvPicPr/></pic:nvPicPr>` +
        `<pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
        `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
      const xml = paragraphXml(undefined, id, "", drawing);
      const node = blockNode(record);
      const items: XmlPatch[] = [
        ...drawingNamespacePatches(part),
        side === "before"
          ? patches.insertBefore(part, node, xml)
          : patches.insertAfter(part, node, xml),
      ];
      return commit(
        context,
        items,
        { createdIds: [`p:${id}`, `img:${id}.0`], reflowFrom: record.id },
        transaction,
      );
    },
  };
