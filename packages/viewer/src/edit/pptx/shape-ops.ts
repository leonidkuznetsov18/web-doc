import { patches, type XmlPatch } from "../ooxml/patch.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { EditColor, PageRect } from "../types.js";
import type { ShapeRecord, SlideElements } from "./elements.js";
import {
  apply as applyMatrix,
  invert,
  pxToEmu,
  type Matrix,
} from "./geometry.js";
import type {
  Issue,
  PptxOperationContext,
  PptxOperationHandler,
  PptxOperationResult,
} from "./operations.js";
import {
  alignValue,
  changedRunProperties,
  colorProblem,
  itemsOfSegment,
  paragraphXml,
  solidFillXml,
  textProblem,
  PARAGRAPH_BREAK,
} from "./text-write.js";
import type {
  PptxDeleteElementOperation,
  PptxInsertTextBoxOperation,
  PptxMoveElementOperation,
  PptxResizeElementOperation,
  PptxSetShapeStyleOperation,
} from "./types.js";

/*
 * Shape operations: fill and line in p:spPr, frames written to a:xfrm (an
 * inheriting placeholder gets an explicit one), deletion with the slide's
 * relationships only the element used, and text boxes appended to the
 * shape tree.
 */

const FILLS = new Set([
  "noFill",
  "solidFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);
/** Children of p:spPr that follow the fill. */
const AFTER_SHAPE_FILL = new Set([
  "ln",
  "effectLst",
  "effectDag",
  "scene3d",
  "sp3d",
  "extLst",
]);
/** Children of p:spPr that follow a:ln. */
const AFTER_LINE = new Set([
  "effectLst",
  "effectDag",
  "scene3d",
  "sp3d",
  "extLst",
]);

async function editableTarget(
  id: string,
  context: PptxOperationContext,
  issue: Issue,
): Promise<ShapeRecord | undefined> {
  const record = await context.locate(id);
  if (!record) {
    issue("/target", "unknown-target", `No element ${id}`);
    return undefined;
  }
  if (record.readOnly) {
    issue("/target", "invalid-target", `Element ${id} is read-only`);
    return undefined;
  }
  return record;
}

function sliceOf(part: XmlPart, node: XmlElement): string {
  return part.text.slice(node.start, node.end);
}

function commit(
  context: PptxOperationContext,
  record: ShapeRecord,
  items: readonly XmlPatch[],
  extra: {
    readonly createdIds?: readonly string[];
    readonly removedIds?: readonly string[];
    readonly relationships?: readonly string[];
  } = {},
): Promise<PptxOperationResult> {
  const transaction = context.pkg.transaction();
  transaction.patch(record.part, items);
  for (const id of extra.relationships ?? [])
    transaction.removeRelationship(record.slide.part, id);
  return transaction.commit().then((change) => ({
    createdIds: extra.createdIds ?? [],
    ...(extra.removedIds ? { removedIds: extra.removedIds } : {}),
    changedPages: [record.element.pageIndex],
    warnings: change.warnings,
  }));
}

/* Frames */

interface FrameEmu {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/** The node whose `a:xfrm` positions the element: p:spPr, p:grpSpPr or the frame itself. */
function frameHolder(record: ShapeRecord): XmlElement | undefined {
  switch (record.node.local) {
    case "graphicFrame":
      return record.node;
    case "grpSp":
      return record.node.children.find((child) => child.local === "grpSpPr");
    default:
      return record.spPr;
  }
}

/** Patches that give the element the frame, creating `a:xfrm` when it inherits one. */
function framePatches(record: ShapeRecord, frame: FrameEmu): XmlPatch[] {
  const { part } = record;
  const xfrm = record.xfrmNode;
  if (xfrm) {
    const off = xfrm.children.find((child) => child.local === "off");
    const ext = xfrm.children.find((child) => child.local === "ext");
    if (off && ext)
      return [
        patches.setAttribute(part, off, "x", String(frame.x)),
        patches.setAttribute(part, off, "y", String(frame.y)),
        patches.setAttribute(part, ext, "cx", String(frame.cx)),
        patches.setAttribute(part, ext, "cy", String(frame.cy)),
      ];
  }
  const xml = `<a:xfrm><a:off x="${frame.x}" y="${frame.y}"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/></a:xfrm>`;
  const holder = frameHolder(record);
  if (!holder) {
    // A shape without p:spPr: add one before the text body or at the end.
    const spPr = `<p:spPr>${xml}</p:spPr>`;
    const txBody = record.node.children.find(
      (child) => child.local === "txBody",
    );
    return [
      txBody
        ? patches.insertBefore(part, txBody, spPr)
        : patches.appendChild(part, record.node, spPr),
    ];
  }
  if (xfrm) return [patches.replaceElement(part, xfrm, xml)];
  const first = holder.children[0];
  if (first) return [patches.insertBefore(part, first, xml)];
  if (holder.selfClosing) return [patches.replaceContent(part, holder, xml)];
  return [patches.appendChild(part, holder, xml)];
}

/** The element's current frame in its own space, in EMU. */
function currentFrame(record: ShapeRecord): FrameEmu | undefined {
  return record.xfrm
    ? {
        x: record.xfrm.x,
        y: record.xfrm.y,
        cx: record.xfrm.cx,
        cy: record.xfrm.cy,
      }
    : undefined;
}

/** A slide-space offset mapped into the element's space through its groups. */
function offsetInElementSpace(
  parents: Matrix,
  dx: number,
  dy: number,
): { x: number; y: number } {
  const inverse = invert(parents);
  return {
    x: inverse.a * dx + inverse.c * dy,
    y: inverse.b * dx + inverse.d * dy,
  };
}

/** The element-space frame whose rotated, group-mapped box fills `rect` (slide px). */
function frameForBounds(record: ShapeRecord, rect: PageRect): FrameEmu {
  const placed = record.placed!;
  const rotation = (placed.frame.rotation * Math.PI) / 180;
  const cos = Math.abs(Math.cos(rotation));
  const sin = Math.abs(Math.sin(rotation));
  const det = cos * cos - sin * sin;
  let width: number;
  let height: number;
  if (Math.abs(det) < 1e-6) {
    // At 45° the box does not determine both sides; keep the aspect ratio.
    const scale =
      placed.bounds.width > 0 ? rect.width / placed.bounds.width : 1;
    width = placed.frame.width * scale;
    height = placed.frame.height * scale;
  } else {
    width = (rect.width * cos - rect.height * sin) / det;
    height = (rect.height * cos - rect.width * sin) / det;
  }
  width = Math.max(width, 0);
  height = Math.max(height, 0);
  const scaleX = Math.hypot(record.parents.a, record.parents.b) || 1;
  const scaleY = Math.hypot(record.parents.c, record.parents.d) || 1;
  const centre = applyMatrix(invert(record.parents), {
    x: pxToEmu(rect.x + rect.width / 2),
    y: pxToEmu(rect.y + rect.height / 2),
  });
  const cx = Math.round(pxToEmu(width) / scaleX);
  const cy = Math.round(pxToEmu(height) / scaleY);
  return {
    x: Math.round(centre.x - cx / 2),
    y: Math.round(centre.y - cy / 2),
    cx,
    cy,
  };
}

async function framedTarget(
  id: string,
  context: PptxOperationContext,
  issue: Issue,
): Promise<ShapeRecord | undefined> {
  const record = await editableTarget(id, context, issue);
  if (!record) return undefined;
  if (!record.xfrm || !record.placed) {
    issue("/target", "invalid-target", `Element ${id} has no frame`);
    return undefined;
  }
  return record;
}

export const moveElementHandler: PptxOperationHandler<PptxMoveElementOperation> =
  {
    async validate(operation, context, issue) {
      if (operation.to === undefined && operation.by === undefined)
        issue("/to", "required", "One of to and by is required");
      if (operation.to !== undefined && operation.by !== undefined)
        issue("/by", "conflict", "Only one of to and by may be given");
      await framedTarget(operation.target, context, issue);
    },
    async apply(operation, context) {
      const record = (await framedTarget(operation.target, context, () => {}))!;
      const bounds = record.placed!.bounds;
      const dx = operation.to ? operation.to.x - bounds.x : operation.by!.dx;
      const dy = operation.to ? operation.to.y - bounds.y : operation.by!.dy;
      const delta = offsetInElementSpace(
        record.parents,
        pxToEmu(dx),
        pxToEmu(dy),
      );
      const own = currentFrame(record)!;
      return commit(
        context,
        record,
        framePatches(record, {
          ...own,
          x: Math.round(own.x + delta.x),
          y: Math.round(own.y + delta.y),
        }),
      );
    },
  };

export const resizeElementHandler: PptxOperationHandler<PptxResizeElementOperation> =
  {
    async validate(operation, context, issue) {
      await framedTarget(operation.target, context, issue);
    },
    async apply(operation, context) {
      const record = (await framedTarget(operation.target, context, () => {}))!;
      // A group scales its children with its extent; its child space stays.
      return commit(
        context,
        record,
        framePatches(record, frameForBounds(record, operation.rect)),
      );
    },
  };

/* Deletion */

/** Relationship ids referenced by attributes in the `r:` namespace under a node. */
function relationshipIdsIn(part: XmlPart, node: XmlElement): Set<string> {
  const ids = new Set<string>();
  const stack: XmlElement[] = [node];
  while (stack.length > 0) {
    const current = stack.pop()!;
    for (const attribute of current.attributes)
      if (attribute.name.startsWith("r:") && attribute.value)
        ids.add(attribute.value);
    for (const child of current.children) stack.push(child);
  }
  return ids;
}

function descendantIds(elements: SlideElements, record: ShapeRecord): string[] {
  const ids: string[] = [];
  for (const candidate of elements.records) {
    let parent = candidate.parent;
    while (parent) {
      if (parent === record) {
        ids.push(candidate.element.id);
        break;
      }
      parent = parent.parent;
    }
  }
  return ids;
}

export const deleteElementHandler: PptxOperationHandler<PptxDeleteElementOperation> =
  {
    async validate(operation, context, issue) {
      await editableTarget(operation.target, context, issue);
    },
    async apply(operation, context) {
      const record = (await editableTarget(
        operation.target,
        context,
        () => {},
      ))!;
      const elements = await context.elements(record.element.pageIndex);
      const { part } = record;
      const inside = relationshipIdsIn(part, record.node);
      const outside = relationshipIdsIn(part, part.root);
      // Ids the rest of the part still uses stay; count their uses outside the element.
      const exclusive = [...inside].filter((id) => {
        let uses = 0;
        const stack: XmlElement[] = [part.root];
        while (stack.length > 0) {
          const current = stack.pop()!;
          if (current === record.node) continue;
          for (const attribute of current.attributes)
            if (attribute.name.startsWith("r:") && attribute.value === id)
              uses += 1;
          for (const child of current.children) stack.push(child);
        }
        return uses === 0 && outside.has(id);
      });
      const relationships = await context.pkg.relationships(record.slide.part);
      return commit(
        context,
        record,
        [patches.removeElement(part, record.node)],
        {
          removedIds: [record.element.id, ...descendantIds(elements, record)],
          relationships: exclusive.filter((id) => relationships.byId(id)),
        },
      );
    },
  };

/* Fill and line */

function fillXml(value: EditColor | "none"): string {
  return value === "none" ? "<a:noFill/>" : solidFillXml(value);
}

/** The spPr content with the fill and line changes applied. */
function styledShapeProperties(
  part: XmlPart,
  spPr: XmlElement | undefined,
  operation: PptxSetShapeStyleOperation,
): string {
  const children: { local: string; xml: string; node?: XmlElement }[] = [];
  if (spPr)
    for (const child of spPr.children)
      children.push({
        local: child.local,
        xml: sliceOf(part, child),
        node: child,
      });
  const replace = (
    matches: (local: string) => boolean,
    after: ReadonlySet<string>,
    local: string,
    xml: string | undefined,
  ): void => {
    const first = children.findIndex((child) => matches(child.local));
    for (let index = children.length - 1; index >= 0; index -= 1)
      if (matches(children[index]!.local)) children.splice(index, 1);
    if (xml === undefined) return;
    let at =
      first >= 0
        ? first
        : children.findIndex((child) => after.has(child.local));
    if (at < 0) at = children.length;
    children.splice(at, 0, { local, xml });
  };
  if (operation.fill !== undefined)
    replace(
      (local) => FILLS.has(local),
      AFTER_SHAPE_FILL,
      "fill",
      operation.fill === null ? undefined : fillXml(operation.fill),
    );
  if (operation.line !== undefined) {
    const existing = children.find((child) => child.local === "ln");
    let xml: string | undefined;
    if (operation.line === null) xml = undefined;
    else {
      const color: EditColor | "none" =
        operation.line === "none" ? "none" : operation.line.color;
      const width =
        operation.line === "none" ? undefined : operation.line.width;
      xml = lineXml(part, existing?.node, color, width);
    }
    replace((local) => local === "ln", AFTER_LINE, "ln", xml);
  }
  return children.map((child) => child.xml).join("");
}

/** `a:ln` with the fill replaced and the width set, other children and attributes kept. */
function lineXml(
  part: XmlPart,
  existing: XmlElement | undefined,
  color: EditColor | "none",
  width: number | undefined,
): string {
  const attributes = new Map<string, string>();
  const children: { local: string; xml: string }[] = [];
  if (existing) {
    for (const attribute of existing.attributes)
      attributes.set(attribute.name, attribute.rawValue);
    for (const child of existing.children)
      children.push({ local: child.local, xml: sliceOf(part, child) });
  }
  if (width !== undefined)
    attributes.set("w", String(Math.round(width * 12700)));
  const first = children.findIndex((child) => FILLS.has(child.local));
  for (let index = children.length - 1; index >= 0; index -= 1)
    if (FILLS.has(children[index]!.local)) children.splice(index, 1);
  children.splice(first >= 0 ? first : 0, 0, {
    local: "fill",
    xml: fillXml(color),
  });
  const attributeText = [...attributes]
    .map(([name, raw]) => ` ${name}="${raw}"`)
    .join("");
  return `<a:ln${attributeText}>${children.map((child) => child.xml).join("")}</a:ln>`;
}

export const setShapeStyleHandler: PptxOperationHandler<PptxSetShapeStyleOperation> =
  {
    async validate(operation, context, issue) {
      if (
        operation.fill !== undefined &&
        operation.fill !== null &&
        operation.fill !== "none"
      ) {
        const problem = colorProblem(operation.fill);
        if (problem) issue("/fill", "invalid-value", `Colour: ${problem}`);
      }
      if (
        operation.line !== undefined &&
        operation.line !== null &&
        operation.line !== "none" &&
        operation.line.color !== "none"
      ) {
        const problem = colorProblem(operation.line.color);
        if (problem)
          issue("/line/color", "invalid-value", `Colour: ${problem}`);
      }
      const record = await editableTarget(operation.target, context, issue);
      if (!record) return;
      const { kind } = record.element;
      if (kind !== "shape" && kind !== "connector")
        issue(
          "/target",
          "invalid-target",
          `Element ${operation.target} has no shape style`,
        );
      else if (kind === "connector" && operation.fill !== undefined)
        issue("/fill", "invalid-value", "A connector has no fill");
    },
    async apply(operation, context) {
      const record = (await editableTarget(
        operation.target,
        context,
        () => {},
      ))!;
      const { part } = record;
      const content = styledShapeProperties(part, record.spPr, operation);
      if (record.spPr)
        return commit(context, record, [
          patches.replaceContent(part, record.spPr, content),
        ]);
      const txBody = record.node.children.find(
        (child) => child.local === "txBody",
      );
      const spPr = `<p:spPr>${content}</p:spPr>`;
      return commit(context, record, [
        txBody
          ? patches.insertBefore(part, txBody, spPr)
          : patches.appendChild(part, record.node, spPr),
      ]);
    },
  };

/* Text boxes */

/** One more than the largest p:cNvPr id of the slide, at least 2. */
export function nextShapeId(elements: SlideElements): number {
  let max = 1;
  for (const record of elements.records) max = Math.max(max, record.cNvPrId);
  const tree = elements.part.find("spTree");
  if (tree)
    for (const node of elements.part.findAll("cNvPr", tree)) {
      const id = Number(elements.part.attribute(node, "id") ?? NaN);
      if (Number.isInteger(id)) max = Math.max(max, id);
    }
  return max + 1;
}

/** Paragraph XML for a text and a style, as a new text box or table cell carries it. */
export function paragraphsXml(
  part: XmlPart,
  text: string,
  style: PptxInsertTextBoxOperation["style"],
): string {
  const { align, ...runChange } = style ?? {};
  const rPr =
    Object.keys(runChange).length > 0
      ? changedRunProperties(part, undefined, runChange)
      : "";
  const endParaRPr = rPr
    ? rPr
        .replace(/^<a:rPr/, "<a:endParaRPr")
        .replace(/<\/a:rPr>$/, "</a:endParaRPr>")
    : "";
  const pPr = align ? `<a:pPr algn="${alignValue(align)}"/>` : "";
  return text
    .split(PARAGRAPH_BREAK)
    .map((segment) =>
      paragraphXml({ pPr, items: itemsOfSegment(segment, rPr), endParaRPr }),
    )
    .join("");
}

export const insertTextBoxHandler: PptxOperationHandler<PptxInsertTextBoxOperation> =
  {
    async validate(operation, context, issue) {
      if (operation.pageIndex < 0 || operation.pageIndex >= context.pageCount)
        issue(
          "/pageIndex",
          "unknown-target",
          `No slide ${operation.pageIndex}`,
        );
      const problem = textProblem(operation.text);
      if (problem) issue("/text", "invalid-text", `The text holds ${problem}`);
      if (operation.style?.color !== undefined) {
        const colour = colorProblem(operation.style.color);
        if (colour) issue("/style/color", "invalid-value", `Colour: ${colour}`);
      }
    },
    async apply(operation, context) {
      const elements = await context.elements(operation.pageIndex);
      const { part, slide } = elements;
      const tree = part.find("spTree");
      if (!tree) throw new Error(`Slide ${slide.part} has no shape tree`);
      const id = nextShapeId(elements);
      const { rect } = operation;
      const xml =
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${pxToEmu(rect.x)}" y="${pxToEmu(rect.y)}"/><a:ext cx="${pxToEmu(rect.width)}" cy="${pxToEmu(rect.height)}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphsXml(part, operation.text, operation.style)}</p:txBody></p:sp>`;
      const transaction = context.pkg.transaction();
      transaction.patch(part, [patches.appendChild(part, tree, xml)]);
      const change = await transaction.commit();
      return {
        createdIds: [`${slide.key}:${id}`],
        changedPages: [operation.pageIndex],
        warnings: change.warnings,
      };
    },
  };
