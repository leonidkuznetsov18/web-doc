import { patches, type XmlPatch } from "../ooxml/patch.js";
import { encodePart, escapeAttribute } from "../ooxml/xml.js";
import type { XmlElement } from "../ooxml/xml.js";
import type { OoxmlPackage } from "../ooxml/package.js";
import {
  relativeTarget,
  type PackageTransaction,
} from "../ooxml/transaction.js";
import { relationshipsPartOf } from "../ooxml/names.js";
import { readPlaceholders } from "./elements.js";
import { partNumber, RELATIONSHIP_TYPES, type LayoutRecord } from "./model.js";
import type {
  Issue,
  PptxOperationContext,
  PptxOperationHandler,
  PptxOperationResult,
} from "./operations.js";
import type {
  PptxDeleteSlideOperation,
  PptxDuplicateSlideOperation,
  PptxInsertSlideOperation,
  PptxMoveSlideOperation,
} from "./types.js";

/*
 * Slide operations: a new slide instantiates its layout's placeholders, a
 * duplicate copies the slide part and clones the parts only it may own,
 * deletion takes the notes slide along, and reordering touches only
 * p:sldIdLst. Every one keeps presentation.xml, its relationships and the
 * content types consistent through one transaction.
 */

const SLIDE_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NAMESPACES =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
const TREE_HEAD =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
/** Placeholder types PowerPoint instantiates on a new slide. */
const INSTANTIATED = new Set([
  "title",
  "ctrTitle",
  "body",
  "subTitle",
  "obj",
  "pic",
  "tbl",
  "chart",
  "dgm",
  "media",
  "clipArt",
]);
/** Relationship types whose targets two slides may share. */
const SHARED_TARGETS = new Set([
  RELATIONSHIP_TYPES.slideLayout,
  RELATIONSHIP_TYPES.image,
  RELATIONSHIP_TYPES.slide,
  `${RELATIONSHIP_TYPES.image.slice(0, -5)}audio`,
  `${RELATIONSHIP_TYPES.image.slice(0, -5)}video`,
  `${RELATIONSHIP_TYPES.image.slice(0, -5)}media`,
  `${RELATIONSHIP_TYPES.image.slice(0, -5)}font`,
  `${RELATIONSHIP_TYPES.image.slice(0, -5)}hyperlink`,
  "http://schemas.microsoft.com/office/2007/relationships/media",
]);
/** Relationship types a duplicate leaves behind. */
const NOT_COPIED = new Set([
  RELATIONSHIP_TYPES.notesSlide,
  `${RELATIONSHIP_TYPES.image.slice(0, -5)}comments`,
  "http://schemas.microsoft.com/office/2018/10/relationships/comments",
]);

function checkIndex(
  value: number,
  path: string,
  maximum: number,
  issue: Issue,
): boolean {
  if (!Number.isInteger(value) || value < 0 || value > maximum) {
    issue(path, "range", `${path.slice(1)} must be between 0 and ${maximum}`);
    return false;
  }
  return true;
}

/** One more than the largest slide number the package ever held. */
function nextSlideNumber(pkg: OoxmlPackage): number {
  let max = 0;
  for (const name of [...pkg.partNames, ...pkg.currentPartNames])
    if (/^\/ppt\/slides\/slide\d+\.xml$/i.test(name))
      max = Math.max(max, partNumber(name));
  return max + 1;
}

/** One more than the largest p:sldId/@id, at least 256. */
function nextSlideId(context: PptxOperationContext): number {
  let max = 255;
  for (const slide of context.model.slides)
    max = Math.max(max, Number(slide.id) || 0);
  return max + 1;
}

/** Patches that put a `p:sldId` at `index` of the list. */
function sldIdPatches(
  context: PptxOperationContext,
  index: number,
  xml: string,
): XmlPatch[] {
  const presentation = context.model.presentation;
  const list = presentation.find("sldIdLst");
  if (!list) throw new Error("The presentation has no slide list");
  const at = context.model.slides[index];
  return [
    at
      ? patches.insertBefore(presentation, at.node, xml)
      : patches.appendChild(presentation, list, xml),
  ];
}

/** Registers a new slide part with the presentation; returns its sldId XML. */
async function registerSlide(
  context: PptxOperationContext,
  transaction: PackageTransaction,
  part: string,
): Promise<string> {
  const rId = await transaction.addRelationship(
    context.model.presentationPart,
    RELATIONSHIP_TYPES.slide,
    relativeTarget(context.model.presentationPart, part),
  );
  return `<p:sldId id="${nextSlideId(context)}" r:id=${escapeAttribute(rId)}/>`;
}

function slideRange(from: number, to: number): number[] {
  const pages: number[] = [];
  for (let page = from; page <= to; page += 1) pages.push(page);
  return pages;
}

export const insertSlideHandler: PptxOperationHandler<PptxInsertSlideOperation> =
  {
    async validate(operation, context, issue) {
      checkIndex(operation.index, "/index", context.pageCount, issue);
      if (
        operation.layout !== undefined &&
        !context.model.layoutById(operation.layout)
      )
        issue("/layout", "unknown-layout", `No layout ${operation.layout}`);
    },
    async apply(operation, context) {
      const { model, pkg } = context;
      const layout: LayoutRecord | undefined = operation.layout
        ? model.layoutById(operation.layout)
        : (model.slides[operation.index - 1]?.layout ??
          model.slides[0]?.layout ??
          model.layouts[0]);
      if (!layout) throw new Error("The deck has no layout to instantiate");
      const number = nextSlideNumber(pkg);
      const part = `/ppt/slides/slide${number}.xml`;
      const placeholders = await readPlaceholders(pkg, layout.part);
      const shapes: string[] = [];
      const createdIds: string[] = [];
      let id = 2;
      for (const placeholder of placeholders.placeholders) {
        if (!INSTANTIATED.has(placeholder.type)) continue;
        const cNvPr = placeholder.node.children
          .find((child) => child.local === "nvSpPr")
          ?.children.find((child) => child.local === "cNvPr");
        const name = cNvPr
          ? placeholders.part.attribute(cNvPr, "name")
          : undefined;
        const ph = placeholder.node.children
          .find((child) => child.local === "nvSpPr")
          ?.children.find((child) => child.local === "nvPr")
          ?.children.find((child) => child.local === "ph");
        const type = ph ? placeholders.part.attribute(ph, "type") : undefined;
        const idx = ph ? placeholders.part.attribute(ph, "idx") : undefined;
        shapes.push(
          `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name=${escapeAttribute(name ?? `Placeholder ${id - 1}`)}/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${
            type ? ` type=${escapeAttribute(type)}` : ""
          }${idx ? ` idx=${escapeAttribute(idx)}` : ""}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`,
        );
        createdIds.push(`sld${number}:${id}`);
        id += 1;
      }
      const xml = `${XML_HEADER}<p:sld ${NAMESPACES}><p:cSld><p:spTree>${TREE_HEAD}${shapes.join(
        "",
      )}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
      const transaction = pkg.transaction();
      transaction.setPart(part, encodePart(xml), SLIDE_CONTENT_TYPE);
      await transaction.addRelationship(
        part,
        RELATIONSHIP_TYPES.slideLayout,
        relativeTarget(part, layout.part),
      );
      const sldId = await registerSlide(context, transaction, part);
      transaction.patch(
        model.presentation,
        sldIdPatches(context, operation.index, sldId),
      );
      const change = await transaction.commit();
      return {
        createdIds,
        changedPages: slideRange(operation.index, model.pageCount),
        warnings: change.warnings,
      };
    },
  };

/** Copies a part and, recursively, the parts only it may own; returns the copy's name. */
async function clonePart(
  context: PptxOperationContext,
  transaction: PackageTransaction,
  source: string,
  target: string,
  contentType: string | undefined,
  seen: Map<string, string>,
): Promise<void> {
  const { pkg } = context;
  seen.set(source, target);
  transaction.setPart(target, await pkg.part(source), contentType);
  const relationships = await pkg.relationships(source);
  if (relationships.items.length === 0) return;
  const copied: string[] = [];
  for (const item of relationships.items) {
    if (NOT_COPIED.has(item.type)) continue;
    let targetValue = item.target;
    if (
      item.targetMode === "Internal" &&
      item.targetPart &&
      !SHARED_TARGETS.has(item.type)
    ) {
      let clone = seen.get(item.targetPart);
      if (!clone) {
        const match = /^(.*?)(\d*)(\.[^./]+)$/.exec(item.targetPart);
        clone = match
          ? transaction.uniquePartName(match[1]!, match[3]!)
          : transaction.uniquePartName(item.targetPart, "");
        await clonePart(
          context,
          transaction,
          item.targetPart,
          clone,
          await pkg.contentTypeOf(item.targetPart),
          seen,
        );
      }
      targetValue = relativeTarget(target, clone);
    }
    copied.push(
      `<Relationship Id=${escapeAttribute(item.id)} Type=${escapeAttribute(item.type)} Target=${escapeAttribute(targetValue)}${
        item.targetMode === "External" ? ' TargetMode="External"' : ""
      }/>`,
    );
  }
  transaction.setPart(
    relationshipsPartOf(target),
    encodePart(
      `${XML_HEADER}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${copied.join("")}</Relationships>`,
    ),
  );
}

export const duplicateSlideHandler: PptxOperationHandler<PptxDuplicateSlideOperation> =
  {
    async validate(operation, context, issue) {
      checkIndex(
        operation.pageIndex,
        "/pageIndex",
        context.pageCount - 1,
        issue,
      );
      if (operation.index !== undefined)
        checkIndex(operation.index, "/index", context.pageCount, issue);
    },
    async apply(operation, context) {
      const { model, pkg } = context;
      const source = model.slideAt(operation.pageIndex);
      const index = operation.index ?? operation.pageIndex + 1;
      const number = nextSlideNumber(pkg);
      const part = `/ppt/slides/slide${number}.xml`;
      const transaction = pkg.transaction();
      await clonePart(
        context,
        transaction,
        source.part,
        part,
        SLIDE_CONTENT_TYPE,
        new Map(),
      );
      const sldId = await registerSlide(context, transaction, part);
      transaction.patch(
        model.presentation,
        sldIdPatches(context, index, sldId),
      );
      const elements = await context.elements(operation.pageIndex);
      const change = await transaction.commit();
      return {
        createdIds: elements.records.map(
          (record) =>
            `sld${number}:${record.element.id.slice(source.key.length + 1)}`,
        ),
        changedPages: slideRange(index, model.pageCount),
        warnings: change.warnings,
      };
    },
  };

export const deleteSlideHandler: PptxOperationHandler<PptxDeleteSlideOperation> =
  {
    async validate(operation, context, issue) {
      if (
        !checkIndex(
          operation.pageIndex,
          "/pageIndex",
          context.pageCount - 1,
          issue,
        )
      )
        return;
      if (context.pageCount === 1)
        issue("/pageIndex", "last-slide", "The last slide cannot be deleted");
    },
    async apply(operation, context) {
      const { model, pkg } = context;
      const slide = model.slideAt(operation.pageIndex);
      const elements = await context.elements(operation.pageIndex);
      const transaction = pkg.transaction();
      transaction.patch(model.presentation, [
        patches.removeElement(model.presentation, slide.node),
      ]);
      transaction.removeRelationship(model.presentationPart, slide.rId);
      const relationships = await pkg.relationships(slide.part);
      for (const notes of relationships.byType(RELATIONSHIP_TYPES.notesSlide))
        if (notes.targetPart && pkg.has(notes.targetPart))
          transaction.removePart(notes.targetPart);
      transaction.removePart(slide.part);
      const change = await transaction.commit();
      return {
        createdIds: [],
        removedIds: elements.records.map((record) => record.element.id),
        changedPages: slideRange(operation.pageIndex, model.pageCount - 1),
        warnings: change.warnings,
      };
    },
  };

export const moveSlideHandler: PptxOperationHandler<PptxMoveSlideOperation> = {
  async validate(operation, context, issue) {
    const last = context.pageCount - 1;
    checkIndex(operation.from, "/from", last, issue);
    checkIndex(operation.to, "/to", last, issue);
  },
  async apply(operation, context) {
    const { model, pkg } = context;
    const { from, to } = operation;
    if (from === to) return { createdIds: [], changedPages: [], warnings: [] };
    const presentation = model.presentation;
    const moving = model.slideAt(from);
    const anchor: XmlElement = model.slideAt(to).node;
    const xml = presentation.text.slice(moving.node.start, moving.node.end);
    const transaction = pkg.transaction();
    transaction.patch(presentation, [
      patches.removeElement(presentation, moving.node),
      to > from
        ? patches.insertAfter(presentation, anchor, xml)
        : patches.insertBefore(presentation, anchor, xml),
    ]);
    const change = await transaction.commit();
    const result: PptxOperationResult = {
      createdIds: [],
      changedPages: slideRange(Math.min(from, to), Math.max(from, to)),
      warnings: change.warnings,
    };
    return result;
  },
};
