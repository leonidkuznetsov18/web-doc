import { resolveBinary } from "../assets.js";
import { patches } from "../ooxml/patch.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { ShapeRecord } from "./elements.js";
import { extensionForMime } from "../ooxml/opc.js";
import { relativeTarget } from "../ooxml/transaction.js";
import { pxToEmu } from "./geometry.js";
import { RELATIONSHIP_TYPES } from "./model.js";
import type {
  Issue,
  PptxOperationContext,
  PptxOperationHandler,
  PptxOperationResult,
} from "./operations.js";
import { nextShapeId, paragraphsXml } from "./shape-ops.js";
import { readTextModel } from "./text.js";
import { replacedBodyContent } from "./text-ops.js";
import { textProblem } from "./text-write.js";
import type {
  PptxInsertImageOperation,
  PptxInsertTableOperation,
  PptxSetTableCellOperation,
} from "./types.js";

/*
 * Pictures and tables: a picture is a media part related from the slide
 * and a p:pic that embeds it; a table is a p:graphicFrame holding a:tbl
 * with a grid, rows of equal height and one paragraph per cell; a cell's
 * text is replaced like a shape's.
 */

const MEDIA_FOLDER = "/ppt/media";
const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

function checkSlide(
  pageIndex: number,
  context: PptxOperationContext,
  issue: Issue,
): boolean {
  if (pageIndex < 0 || pageIndex >= context.model.pageCount) {
    issue("/pageIndex", "unknown-target", `No slide ${pageIndex}`);
    return false;
  }
  return true;
}

export const insertImageHandler: PptxOperationHandler<PptxInsertImageOperation> =
  {
    async validate(operation, context, issue) {
      checkSlide(operation.pageIndex, context, issue);
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
      const elements = await context.elements(operation.pageIndex);
      const { part, slide } = elements;
      const tree = part.find("spTree")!;
      const id = nextShapeId(elements);
      const bytes = resolveBinary(operation.data, context.assets);
      const transaction = context.pkg.transaction();
      // The same bytes already in the package are related, not stored again.
      const existing = await existingMedia(context, bytes, operation.mimeType);
      const rId = existing
        ? await transaction.addRelationship(
            slide.part,
            RELATIONSHIP_TYPES.image,
            relativeTarget(slide.part, existing),
          )
        : (
            await transaction.addMedia(
              slide.part,
              MEDIA_FOLDER,
              bytes,
              operation.mimeType,
            )
          ).rId;
      const { rect } = operation;
      const xml =
        `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id - 1}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr><a:xfrm><a:off x="${pxToEmu(rect.x)}" y="${pxToEmu(rect.y)}"/><a:ext cx="${pxToEmu(rect.width)}" cy="${pxToEmu(rect.height)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
      transaction.patch(part, [patches.appendChild(part, tree, xml)]);
      const change = await transaction.commit();
      return {
        createdIds: [`${slide.key}:${id}`],
        changedPages: [operation.pageIndex],
        warnings: change.warnings,
      };
    },
  };

/** A media part of the package with these bytes and this type's extension, if any. */
async function existingMedia(
  context: PptxOperationContext,
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

/** The default table style id the deck names, when it has a style list. */
async function defaultTableStyle(
  context: PptxOperationContext,
): Promise<string | undefined> {
  const name = "/ppt/tableStyles.xml";
  if (!context.pkg.has(name)) return undefined;
  const part = await context.pkg.xml(name);
  return part.attribute(part.root, "def");
}

export const insertTableHandler: PptxOperationHandler<PptxInsertTableOperation> =
  {
    async validate(operation, context, issue) {
      checkSlide(operation.pageIndex, context, issue);
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
          `Expected ${columns} column weights`,
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
      const elements = await context.elements(operation.pageIndex);
      const { part, slide } = elements;
      const tree = part.find("spTree")!;
      const id = nextShapeId(elements);
      const { rect, rows } = operation;
      const columns = rows[0]!.length;
      const weights =
        operation.columnWidths ?? Array.from({ length: columns }, () => 1);
      const total = weights.reduce((sum, weight) => sum + weight, 0);
      const width = pxToEmu(rect.width);
      const gridWidths: number[] = [];
      let used = 0;
      weights.forEach((weight, index) => {
        const value =
          index === weights.length - 1
            ? width - used
            : Math.round((width * weight) / total);
        gridWidths.push(value);
        used += value;
      });
      const height = pxToEmu(rect.height);
      const rowHeights = rows.map((_, index) =>
        index === rows.length - 1
          ? height - Math.round(height / rows.length) * (rows.length - 1)
          : Math.round(height / rows.length),
      );
      const style = operation.style ?? {};
      const styleId = await defaultTableStyle(context);
      const tblPr = `<a:tblPr firstRow="${style.firstRow === false ? 0 : 1}" bandRow="${style.bandRow === false ? 0 : 1}"${
        styleId
          ? `><a:tableStyleId>${styleId}</a:tableStyleId></a:tblPr>`
          : "/>"
      }`;
      const grid = `<a:tblGrid>${gridWidths.map((w) => `<a:gridCol w="${w}"/>`).join("")}</a:tblGrid>`;
      const body = rows
        .map(
          (row, rowIndex) =>
            `<a:tr h="${rowHeights[rowIndex]}">${row
              .map(
                (cell) =>
                  `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml(part, cell, undefined)}</a:txBody><a:tcPr/></a:tc>`,
              )
              .join("")}</a:tr>`,
        )
        .join("");
      const xml =
        `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id - 1}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>` +
        `<p:xfrm><a:off x="${pxToEmu(rect.x)}" y="${pxToEmu(rect.y)}"/><a:ext cx="${width}" cy="${height}"/></p:xfrm>` +
        `<a:graphic><a:graphicData uri="${TABLE_URI}"><a:tbl>${tblPr}${grid}${body}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
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

interface CellTarget {
  readonly record: ShapeRecord;
  readonly part: XmlPart;
  readonly cell: XmlElement;
}

async function cellTarget(
  operation: PptxSetTableCellOperation,
  context: PptxOperationContext,
  issue: Issue,
): Promise<CellTarget | undefined> {
  const record = await context.locate(operation.target);
  if (!record) {
    issue("/target", "unknown-target", `No element ${operation.target}`);
    return undefined;
  }
  const tbl = record.part.find("tbl", record.node);
  if (record.readOnly || record.element.kind !== "table" || !tbl) {
    issue(
      "/target",
      "invalid-target",
      `Element ${operation.target} is not an editable table`,
    );
    return undefined;
  }
  const rows = tbl.children.filter((child) => child.local === "tr");
  const row = rows[operation.row];
  const cell = row?.children.filter((child) => child.local === "tc")[
    operation.column
  ];
  if (!row) {
    issue("/row", "range", `The table has ${rows.length} rows`);
    return undefined;
  }
  if (!cell) {
    issue(
      "/column",
      "range",
      `Row ${operation.row} has ${row.children.filter((child) => child.local === "tc").length} cells`,
    );
    return undefined;
  }
  return { record, part: record.part, cell };
}

export const setTableCellHandler: PptxOperationHandler<PptxSetTableCellOperation> =
  {
    async validate(operation, context, issue) {
      const problem = textProblem(operation.text);
      if (problem) issue("/text", "invalid-text", `The text holds ${problem}`);
      await cellTarget(operation, context, issue);
    },
    async apply(operation, context) {
      const { record, part, cell } = (await cellTarget(
        operation,
        context,
        () => {},
      ))!;
      const txBody = cell.children.find((child) => child.local === "txBody");
      const transaction = context.pkg.transaction();
      if (txBody) {
        const model = readTextModel(part, txBody);
        transaction.patch(part, [
          patches.replaceContent(
            part,
            txBody,
            replacedBodyContent(
              part,
              txBody,
              model,
              operation.text,
              0,
              model.text.length,
              true,
            ),
          ),
        ]);
      } else {
        const xml = `<a:txBody><a:bodyPr/><a:lstStyle/>${paragraphsXml(part, operation.text, undefined)}</a:txBody>`;
        const tcPr = cell.children.find((child) => child.local === "tcPr");
        transaction.patch(part, [
          tcPr
            ? patches.insertBefore(part, tcPr, xml)
            : patches.appendChild(part, cell, xml),
        ]);
      }
      const change = await transaction.commit();
      const result: PptxOperationResult = {
        createdIds: [],
        changedPages: [record.element.pageIndex],
        warnings: change.warnings,
      };
      return result;
    },
  };
