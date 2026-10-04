import { invalidOperationError } from "../operations.js";
import { patches, type XmlPatch } from "../ooxml/patch.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { ViewerWarning } from "../../contracts.js";
import type { EditColor, PageRect } from "../types.js";
import type { ShapeRecord, SlideElements } from "./elements.js";
import {
  apply as applyMatrix,
  invert,
  pxToEmu,
  EMU_PER_PX,
  groupMatrix,
  multiply,
  placeFrame,
  type Matrix,
  type Xfrm,
} from "./geometry.js";
import {
  committedParts,
  type Issue,
  type PptxOperationContext,
  type PptxOperationHandler,
  type PptxOperationResult,
} from "./operations.js";
import {
  alignValue,
  changedRunProperties,
  colorProblem,
  itemsOfSegment,
  normalizeText,
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
    readonly warnings?: readonly ViewerWarning[];
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
    warnings: [...(extra.warnings ?? []), ...change.warnings],
    parts: committedParts(change),
  }));
}

/* Frames */

interface FrameEmu {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  readonly child?: NonNullable<Xfrm["child"]>;
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
    if (off && ext) {
      const result: XmlPatch[] = [
        patches.setAttribute(part, off, "x", String(frame.x)),
        patches.setAttribute(part, off, "y", String(frame.y)),
        patches.setAttribute(part, ext, "cx", String(frame.cx)),
        patches.setAttribute(part, ext, "cy", String(frame.cy)),
      ];
      if (frame.child) {
        const childOff = xfrm.children.find((child) => child.local === "chOff");
        const childExt = xfrm.children.find((child) => child.local === "chExt");
        const offXml = `<a:chOff x="${frame.child.x}" y="${frame.child.y}"/>`;
        const extXml = `<a:chExt cx="${frame.child.cx}" cy="${frame.child.cy}"/>`;
        if (childOff) {
          result.push(
            patches.setAttribute(part, childOff, "x", String(frame.child.x)),
            patches.setAttribute(part, childOff, "y", String(frame.child.y)),
          );
        } else {
          result.push(
            childExt
              ? patches.insertBefore(part, childExt, offXml)
              : patches.appendChild(part, xfrm, offXml),
          );
        }
        if (childExt) {
          result.push(
            patches.setAttribute(part, childExt, "cx", String(frame.child.cx)),
            patches.setAttribute(part, childExt, "cy", String(frame.child.cy)),
          );
        } else {
          result.push(patches.appendChild(part, xfrm, extXml));
        }
      }
      return result;
    }
  }
  const tag = record.node.local === "graphicFrame" ? "p:xfrm" : "a:xfrm";
  const childXml = frame.child
    ? `<a:chOff x="${frame.child.x}" y="${frame.child.y}"/><a:chExt cx="${frame.child.cx}" cy="${frame.child.cy}"/>`
    : "";
  const xml = `<${tag}><a:off x="${frame.x}" y="${frame.y}"/><a:ext cx="${frame.cx}" cy="${frame.cy}"/>${childXml}</${tag}>`;
  const holder = frameHolder(record);
  if (!holder) {
    // A shape without p:spPr: add one where the schema puts it.
    return [
      insertShapeProperties(part, record.node, `<p:spPr>${xml}</p:spPr>`),
    ];
  }
  if (xfrm) return [patches.replaceElement(part, xfrm, xml)];
  if (record.node.local === "graphicFrame") {
    // p:xfrm follows p:nvGraphicFramePr and precedes a:graphic.
    const graphic = holder.children.find((child) => child.local === "graphic");
    return [
      graphic
        ? patches.insertBefore(part, graphic, xml)
        : patches.appendChild(part, holder, xml),
    ];
  }
  const first = holder.children[0];
  if (first) return [patches.insertBefore(part, first, xml)];
  if (holder.selfClosing) return [patches.replaceContent(part, holder, xml)];
  return [patches.appendChild(part, holder, xml)];
}

/** A patch that adds `p:spPr` after the non-visual properties, before p:style and the text body. */
function insertShapeProperties(
  part: XmlPart,
  node: XmlElement,
  xml: string,
): XmlPatch {
  const next = node.children.find(
    (child) => child.local === "style" || child.local === "txBody",
  );
  return next
    ? patches.insertBefore(part, next, xml)
    : patches.appendChild(part, node, xml);
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

/**
 * EMU frame rounding moves the centre by at most one EMU and each extent
 * by half an EMU. Two EMU per parent axis conservatively bounds the resulting
 * slide-space corner error, including conversion of the requested centre.
 */
const FRAME_ROUNDING_EMU = 2;
const MATRIX_SINGULAR_TOLERANCE = 1e-10;
const MATRIX_FLOAT_TOLERANCE = 16 * Number.EPSILON;
const UNREPRESENTABLE_FRAME_MESSAGE =
  "The box cannot preserve the element frame and group placement";

/** Solve the actual affine corner bounds, rather than the displayed frame angle. */
function frameForBounds(
  record: ShapeRecord,
  rect: PageRect,
): FrameEmu | undefined {
  const own = record.xfrm;
  const placed = record.placed;
  if (!own || !placed) return undefined;
  const parents = record.parents;
  if (!Object.values(parents).every(Number.isFinite)) return undefined;
  const parentDet = parents.a * parents.d - parents.b * parents.c;
  if (!Number.isFinite(parentDet) || parentDet === 0) return undefined;
  const rotation = (own.rotation * Math.PI) / 180;
  const cos = Math.cos(rotation);
  const sin = Math.sin(rotation);
  // Flips negate columns; they cannot change an axis-aligned extent.
  const a = Math.abs(parents.a * cos + parents.c * sin);
  const b = Math.abs(parents.b * cos + parents.d * sin);
  const c = Math.abs(-parents.a * sin + parents.c * cos);
  const d = Math.abs(-parents.b * sin + parents.d * cos);
  if (record.node.local === "grpSp") {
    const sx = rect.width / placed.bounds.width;
    const sy = rect.height / placed.bounds.height;
    // An anisotropic world scale preserves a basis direction only when it
    // lies on a world axis. Test representability before quantization: a
    // small shear at nominal corners can grow arbitrarily on overflowing
    // descendants, and must not be excused as EMU rounding.
    const uniform =
      Math.abs(sx - sy) <=
      MATRIX_FLOAT_TOLERANCE * Math.max(Math.abs(sx), Math.abs(sy));
    const firstAxis = Math.min(a, b) <= MATRIX_FLOAT_TOLERANCE * Math.max(a, b);
    const secondAxis =
      Math.min(c, d) <= MATRIX_FLOAT_TOLERANCE * Math.max(c, d);
    if (
      ![sx, sy].every(Number.isFinite) ||
      (!uniform && !(firstAxis && secondAxis))
    )
      return undefined;
  }
  const det = a * d - b * c;
  const width = rect.width * EMU_PER_PX;
  const height = rect.height * EMU_PER_PX;
  let cx: number;
  let cy: number;
  if (Math.abs(det) <= MATRIX_SINGULAR_TOLERANCE * Math.max(a * d, b * c)) {
    // A rank-one AABB only permits proportional dimensions. Preserve aspect
    // ratio; the forward check below also verifies the requested height.
    const scale = rect.width / placed.bounds.width;
    cx = own.cx * scale;
    cy = own.cy * scale;
  } else {
    cx = (width * d - height * c) / det;
    cy = (height * a - width * b) / det;
  }
  const centre = applyMatrix(invert(parents), {
    x: (rect.x + rect.width / 2) * EMU_PER_PX,
    y: (rect.y + rect.height / 2) * EMU_PER_PX,
  });
  // Freeze the original effective child space before changing group extents.
  // The renderer's absent/zero extent fallback is identity on that axis.
  const child =
    record.node.local === "grpSp" &&
    (!own.child || own.child.cx === 0 || own.child.cy === 0)
      ? {
          x: own.child?.x ?? 0,
          y: own.child?.y ?? 0,
          cx: own.child?.cx || own.cx,
          cy: own.child?.cy || own.cy,
        }
      : undefined;
  const frame: FrameEmu = {
    ...(child ? { child } : {}),
    cx: Math.round(cx),
    cy: Math.round(cy),
    x: Math.round(centre.x - Math.round(cx) / 2),
    y: Math.round(centre.y - Math.round(cy) / 2),
  };
  if (
    ![frame.x, frame.y, frame.cx, frame.cy].every(Number.isFinite) ||
    frame.cx <= 0 ||
    frame.cy <= 0
  )
    return undefined;
  const candidate = { ...own, ...frame };
  const actual = placeFrame(candidate, parents).bounds;
  if (!Object.values(actual).every(Number.isFinite)) return undefined;
  const toleranceX =
    (FRAME_ROUNDING_EMU * (Math.abs(parents.a) + Math.abs(parents.c))) /
      EMU_PER_PX +
    1e-7;
  const toleranceY =
    (FRAME_ROUNDING_EMU * (Math.abs(parents.b) + Math.abs(parents.d))) /
      EMU_PER_PX +
    1e-7;
  if (
    Math.abs(actual.x - rect.x) > toleranceX ||
    Math.abs(actual.width - rect.width) > toleranceX ||
    Math.abs(actual.y - rect.y) > toleranceY ||
    Math.abs(actual.height - rect.height) > toleranceY
  )
    return undefined;
  if (record.node.local === "grpSp") {
    // Matching a group's AABB alone can shear the requested composition.
    // Require its complete child-space mapping to equal world-axis scaling.
    const previous = multiply(parents, groupMatrix(own));
    const next = multiply(parents, groupMatrix(candidate));
    const child = own.child ?? { x: 0, y: 0, cx: own.cx, cy: own.cy };
    const sx = rect.width / placed.bounds.width;
    const sy = rect.height / placed.bounds.height;
    for (const point of [
      { x: child.x, y: child.y },
      { x: child.x + (child.cx || own.cx), y: child.y },
      { x: child.x, y: child.y + (child.cy || own.cy) },
      { x: child.x + (child.cx || own.cx), y: child.y + (child.cy || own.cy) },
    ]) {
      const before = applyMatrix(previous, point);
      const after = applyMatrix(next, point);
      const expectedX =
        rect.x * EMU_PER_PX + (before.x - placed.bounds.x * EMU_PER_PX) * sx;
      const expectedY =
        rect.y * EMU_PER_PX + (before.y - placed.bounds.y * EMU_PER_PX) * sy;
      if (
        ![before.x, before.y, after.x, after.y, expectedX, expectedY].every(
          Number.isFinite,
        )
      )
        return undefined;
      if (
        Math.abs(after.x - expectedX) > toleranceX * EMU_PER_PX ||
        Math.abs(after.y - expectedY) > toleranceY * EMU_PER_PX
      )
        return undefined;
    }
  }
  return frame;
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
      const record = await framedTarget(operation.target, context, issue);
      if (!record) return;
      const frame = frameForBounds(record, operation.rect);
      if (!frame)
        issue("/rect", "invalid-value", UNREPRESENTABLE_FRAME_MESSAGE);
    },
    async apply(operation, context) {
      const record = (await framedTarget(operation.target, context, () => {}))!;
      const frame = frameForBounds(record, operation.rect);
      if (!frame)
        throw invalidOperationError([
          {
            operationIndex: context.operationIndex,
            path: "/rect",
            code: "invalid-value",
            message: UNREPRESENTABLE_FRAME_MESSAGE,
          },
        ]);
      // A group scales its children with its extent; its child space stays.
      return commit(context, record, framePatches(record, frame));
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
      const removedIds = [
        record.element.id,
        ...descendantIds(elements, record),
      ];
      const items = [patches.removeElement(part, record.node)];
      const warnings: ViewerWarning[] = [];
      // An animation that targets a removed shape makes PowerPoint repair
      // the file; the slide's timing goes with the shape, and the host hears.
      const timing = part.root.children.find(
        (child) => child.local === "timing",
      );
      const shapeIds = new Set(
        [
          record,
          ...elements.records.filter((candidate) =>
            removedIds.includes(candidate.element.id),
          ),
        ].map((candidate) => String(candidate.cNvPrId)),
      );
      if (
        timing &&
        part
          .findAll("spTgt", timing)
          .some((node) => shapeIds.has(part.attribute(node, "spid") ?? ""))
      ) {
        items.push(patches.removeElement(part, timing));
        warnings.push({
          code: "fidelity-degraded",
          message:
            "The slide's animations were removed with the shape they targeted",
          details: {
            reason: "animations-removed",
            elementId: record.element.id,
          },
        });
      }
      return commit(context, record, items, {
        removedIds,
        relationships: exclusive.filter((id) => relationships.byId(id)),
        warnings,
      });
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
    },
    async apply(operation, context) {
      const elements = await context.elements(operation.pageIndex);
      const { part, slide } = elements;
      const tree = part.find("spTree");
      if (!tree) throw new Error(`Slide ${slide.part} has no shape tree`);
      const id = context.allocateShapeId(elements);
      const { rect } = operation;
      const xml =
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${pxToEmu(rect.x)}" y="${pxToEmu(rect.y)}"/><a:ext cx="${pxToEmu(rect.width)}" cy="${pxToEmu(rect.height)}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphsXml(part, normalizeText(operation.text), operation.style)}</p:txBody></p:sp>`;
      const transaction = context.pkg.transaction();
      transaction.patch(part, [patches.appendChild(part, tree, xml)]);
      const change = await transaction.commit();
      return {
        createdIds: [`${slide.key}:${id}`],
        changedPages: [operation.pageIndex],
        warnings: change.warnings,
        parts: committedParts(change),
      };
    },
  };
