import { validateSchema } from "../../schema.js";
import type { PagePoint, PageRect } from "../../types.js";
import {
  tableMarkSchema,
  tableMemberMarkSchema,
  textBoxMarkSchema,
} from "../schemas.js";
import type { PdfElement, PdfShapeStyle, PdfTextStyle } from "../types.js";
import {
  concat,
  round,
  roundRect,
  unionRects,
  userRectToPage,
  type Matrix,
  type PageGeometry,
} from "./geometry.js";
import type { Pdfium } from "./pdfium.js";
import type { TableSpec } from "./tables.js";

/** FPDFPageObj_GetType values. */
export const OBJECT_TEXT = 1;
export const OBJECT_PATH = 2;
export const OBJECT_IMAGE = 3;
export const OBJECT_FORM = 5;

/** Marked-content tag and parameter that carry web-doc's own elements. */
export const MARK_NAME = "WebDoc";
export const MARK_PARAM = "webdoc";

/** Font descriptor flag bits (PDF 32000-1 table 123). */
const FLAG_ITALIC = 1 << 6;
const FLAG_FORCE_BOLD = 1 << 18;

/** What a `WebDoc` mark says about the objects it tags. */
export interface MarkParams {
  readonly kind: "textBox" | "table";
  readonly id: string;
  readonly [key: string]: unknown;
}

/** One PDFium page object with what the model knows about it. */
export interface ObjectRecord {
  /** Element id of a plain object; objects of a marked group share their group id. */
  readonly id: string;
  readonly type: number;
  readonly mark?: MarkParams;
  /**
   * The id a mark the file still carries names, on an object listed as plain
   * because the mark failed its check: a new id must not take it, or saving
   * would join the new object to this one.
   */
  readonly staleMarkId?: string;
  /** A Form XObject's own objects, in its drawing order. */
  readonly children?: readonly ObjectRecord[];
}

/** What a page scan found: its elements and who draws each character. */
export interface PageScan {
  readonly elements: PdfElement[];
  /** Object handle → element id, in drawing order, text inside forms included. */
  readonly byObject: Map<number, string>;
  /** Text objects inside forms → the matrix that maps their form onto the page. */
  readonly outer: Map<number, Matrix>;
}

/**
 * Reads the objects of a loaded page in drawing order and turns them into
 * elements: plain objects become one element each, objects tagged with the
 * same `WebDoc` mark become one composite element. Text inside Form
 * XObjects becomes elements of its own, listed after its form.
 */
export function scanPage(
  pdfium: Pdfium,
  page: number,
  textPage: number,
  pageIndex: number,
  geometry: PageGeometry,
  records: readonly ObjectRecord[],
): PageScan {
  const { lib } = pdfium;
  const byObject = new Map<number, string>();
  const outer = new Map<number, Matrix>();
  const elements: PdfElement[] = [];
  const groups = new Map<string, PdfElement[]>();
  const count = Math.min(lib.FPDFPage_CountObjects(page), records.length);
  for (let index = 0; index < count; index += 1) {
    const object = lib.FPDFPage_GetObject(page, index);
    const record = records[index]!;
    byObject.set(object, record.id);
    const element = plainElement(
      pdfium,
      object,
      textPage,
      record,
      pageIndex,
      geometry,
    );
    if (!element) continue;
    if (record.mark) {
      const members = groups.get(record.id);
      if (members) members.push(element);
      else {
        groups.set(record.id, [element]);
        elements.push(element); // placeholder keeps the group's z-order
      }
    } else elements.push(element);
    if (record.children)
      formTexts(pdfium, object, record.children, objectMatrix(pdfium, object), {
        textPage,
        pageIndex,
        geometry,
        byObject,
        outer,
        elements,
      });
  }
  return {
    elements: elements.map((element) => {
      const members = groups.get(element.id);
      const record =
        members && records.find((entry) => entry.id === element.id);
      return members && record?.mark
        ? compositeElement(record.mark, members, pageIndex)
        : element;
    }),
    byObject,
    outer,
  };
}

/**
 * Adds the text objects of a form, and of the forms inside it, as text
 * elements; `matrix` maps the form's space onto the page. Other objects in
 * forms stay part of their form's element.
 */
function formTexts(
  pdfium: Pdfium,
  form: number,
  records: readonly ObjectRecord[],
  matrix: Matrix,
  scan: {
    readonly textPage: number;
    readonly pageIndex: number;
    readonly geometry: PageGeometry;
    readonly byObject: Map<number, string>;
    readonly outer: Map<number, Matrix>;
    readonly elements: PdfElement[];
  },
): void {
  const { lib } = pdfium;
  const count = Math.min(lib.FPDFFormObj_CountObjects(form), records.length);
  const geometry = { ...scan.geometry, matrix };
  for (let index = 0; index < count; index += 1) {
    const object = lib.FPDFFormObj_GetObject(form, index);
    const record = records[index]!;
    if (record.type === OBJECT_TEXT) {
      scan.byObject.set(object, record.id);
      scan.outer.set(object, matrix);
      const element = plainElement(
        pdfium,
        object,
        scan.textPage,
        record,
        scan.pageIndex,
        geometry,
      );
      if (element)
        scan.elements.push({ ...element, operations: FORM_TEXT_OPERATIONS });
    } else if (record.children)
      formTexts(
        pdfium,
        object,
        record.children,
        concat(objectMatrix(pdfium, object), matrix),
        scan,
      );
  }
}

/** An object's own matrix, identity when PDFium has none. */
export function objectMatrix(pdfium: Pdfium, object: number): Matrix {
  return (pdfium.readNumbers(6, "float", ([pointer]) =>
    pdfium.lib.FPDFPageObj_GetMatrix(object, pointer!),
  ) ?? [1, 0, 0, 1, 0, 0]) as unknown as Matrix;
}

/** The `WebDoc` mark of an object, if it carries a valid one. */
export function readMark(
  pdfium: Pdfium,
  object: number,
): MarkParams | undefined {
  const { lib } = pdfium;
  for (let index = 0; index < lib.FPDFPageObj_CountMarks(object); index += 1) {
    const mark = lib.FPDFPageObj_GetMark(object, index);
    const name = pdfium.readWideStringOut((buffer, bytes, out) =>
      lib.FPDFPageObjMark_GetName(mark, buffer, bytes, out),
    );
    if (name !== MARK_NAME) continue;
    const raw = pdfium.readWideStringOut((buffer, bytes, out) =>
      lib.FPDFPageObjMark_GetParamStringValue(
        mark,
        MARK_PARAM,
        buffer,
        bytes,
        out,
      ),
    );
    try {
      const params: unknown = JSON.parse(raw);
      // Marks come from files, so they pass the same checks as operations.
      if (params && typeof params === "object" && isValidMark(params))
        return params as MarkParams;
    } catch {
      // A foreign or damaged mark leaves the object a plain element.
    }
  }
  return undefined;
}

function isValidMark(params: object): boolean {
  switch ((params as MarkParams).kind) {
    case "textBox":
      return validateSchema(params, textBoxMarkSchema, 0).length === 0;
    case "table":
      return (
        validateSchema(params, tableMarkSchema, 0).length === 0 ||
        validateSchema(params, tableMemberMarkSchema, 0).length === 0
      );
    default:
      return false;
  }
}

/**
 * Whether a marked group still looks like what its inputs describe. Another
 * tool may have moved, resized or retyped the objects while keeping the mark;
 * rebuilding from stale inputs would then undo that edit silently, so such a
 * group is listed as plain objects instead.
 */
export function markIsFresh(
  pdfium: Pdfium,
  page: number,
  textPage: number,
  geometry: PageGeometry,
  mark: MarkParams,
  indexes: readonly number[],
): boolean {
  const { lib } = pdfium;
  const objects = indexes.map((index) => lib.FPDFPage_GetObject(page, index));
  const bounds = objects
    .map((object) => objectBounds(pdfium, object, geometry))
    .filter((rect): rect is PageRect => rect !== undefined);
  if (bounds.length === 0) return false;
  const union = unionRects(bounds);
  const drawn = normalizeText(
    objects
      .filter((object) => lib.FPDFPageObj_GetType(object) === OBJECT_TEXT)
      .map((object) => textOf(pdfium, object, textPage))
      .join(" "),
  );
  const tolerance = 2;
  if (mark.kind === "textBox") {
    const rect = mark.rect as PageRect;
    const text = mark.text as string;
    const style = mark.style as {
      readonly fontSize: number;
      readonly lineHeight: number;
    };
    return (
      drawn === normalizeText(text) &&
      union.x >= rect.x - tolerance &&
      union.x + union.width <= rect.x + rect.width + tolerance &&
      // The first line's ink starts inside the box's first line band; the
      // bottom is not checked because overflowing text runs past the box.
      union.y >= rect.y - tolerance &&
      union.y <= rect.y + style.fontSize * style.lineHeight + tolerance
    );
  }
  const at = mark.at as PagePoint;
  const widths = mark.columnWidths as readonly number[];
  const style = mark.style as { readonly borderWidth: number };
  const rows = mark.rows as readonly (readonly string[])[];
  const slack = tolerance + style.borderWidth;
  return (
    drawn === normalizeText(rows.flat().join(" ")) &&
    Math.abs(union.x - at.x) <= slack &&
    Math.abs(union.y - at.y) <= slack &&
    Math.abs(union.width - widths.reduce((sum, w) => sum + w, 0)) <= 2 * slack
  );
}

function normalizeText(text: string): string {
  return text.replaceAll(/\s+/g, " ").trim();
}

export function objectBounds(
  pdfium: Pdfium,
  object: number,
  geometry: PageGeometry,
): PageRect | undefined {
  const bounds = pdfium.readNumbers(4, "float", ([left, bottom, right, top]) =>
    pdfium.lib.FPDFPageObj_GetBounds(object, left!, bottom!, right!, top!),
  );
  if (!bounds) return undefined;
  const [left, bottom, right, top] = bounds as [number, number, number, number];
  return roundRect(userRectToPage(geometry, left, bottom, right, top));
}

function plainElement(
  pdfium: Pdfium,
  object: number,
  textPage: number,
  record: ObjectRecord,
  pageIndex: number,
  geometry: PageGeometry,
): PdfElement | undefined {
  const bounds = objectBounds(pdfium, object, geometry);
  if (!bounds) return undefined;
  const rotation = objectRotation(pdfium, object, geometry);
  const base = {
    id: record.id,
    pageIndex,
    bounds,
    ...(rotation === 0 ? {} : { rotation }),
  };
  switch (record.type) {
    case OBJECT_TEXT:
      return {
        ...base,
        kind: "text",
        text: textOf(pdfium, object, textPage),
        textStyle: textStyle(pdfium, object, geometry.matrix),
        operations: TEXT_OPERATIONS,
      };
    case OBJECT_IMAGE:
      return { ...base, kind: "image", operations: TRANSFORM_OPERATIONS };
    case OBJECT_PATH:
      return {
        ...base,
        kind: "shape",
        shapeStyle: shapeStyle(pdfium, object),
        operations: SHAPE_OPERATIONS,
      };
    default:
      return { ...base, kind: "other", operations: TRANSFORM_OPERATIONS };
  }
}

function compositeElement(
  mark: MarkParams,
  members: readonly PdfElement[],
  pageIndex: number,
): PdfElement {
  const bounds = roundRect(unionRects(members.map((member) => member.bounds)));
  const text = members
    .filter((member) => member.kind === "text" && member.text)
    .map((member) => member.text)
    .join("\n");
  if (mark.kind === "table") {
    // The head mark comes first in drawing order and carries the inputs.
    const spec = mark as unknown as TableSpec;
    const cells = spec.rows.map((row) => row.join("\t")).join("\n");
    return {
      id: mark.id,
      kind: "table",
      pageIndex,
      bounds,
      ...(cells.trim() ? { text: cells } : {}),
      shapeStyle: {
        stroke: {
          color: spec.style.borderColor,
          width: spec.style.borderWidth,
        },
        ...(spec.style.headerFill
          ? { fill: { color: spec.style.headerFill } }
          : {}),
      },
      table: { rows: spec.rows },
      operations: TABLE_OPERATIONS,
    };
  }
  const style = members.find((member) => member.textStyle)?.textStyle;
  return {
    id: mark.id,
    kind: "textBox",
    pageIndex,
    bounds,
    text: typeof mark.text === "string" ? mark.text : text,
    ...(style ? { textStyle: style } : {}),
    operations: TEXT_BOX_OPERATIONS,
  };
}

const TRANSFORM_OPERATIONS = Object.freeze([
  "moveElement",
  "resizeElement",
  "deleteElement",
]);
const TEXT_OPERATIONS = Object.freeze([
  "replaceText",
  "setTextStyle",
  ...TRANSFORM_OPERATIONS,
]);
const TEXT_BOX_OPERATIONS = TEXT_OPERATIONS;
const SHAPE_OPERATIONS = Object.freeze([
  "setShapeStyle",
  ...TRANSFORM_OPERATIONS,
]);
/** Text inside a form is rewritten with its form; stretching it is not offered. */
const FORM_TEXT_OPERATIONS = Object.freeze([
  "replaceText",
  "setTextStyle",
  "moveElement",
  "deleteElement",
]);
const TABLE_OPERATIONS = Object.freeze([
  "setTableCell",
  "moveElement",
  "deleteElement",
]);

function textOf(pdfium: Pdfium, object: number, textPage: number): string {
  return pdfium.readWideString((buffer, bytes) =>
    pdfium.lib.FPDFTextObj_GetText(object, textPage, buffer, bytes),
  );
}

/**
 * How much a text object's matrix scales its glyphs: the length of its
 * vertical axis. Producers often write `1 Tf` and carry the size in the
 * matrix, so the point size the text shows is its font size times this; a
 * turn or a horizontal squeeze leaves it alone.
 */
export function textScale(matrix: readonly number[] | undefined): number {
  const scale = matrix ? Math.hypot(matrix[2]!, matrix[3]!) : 1;
  return Number.isFinite(scale) && scale > 0 ? scale : 1;
}

/** The style a text object shows; `outer` maps the form it sits in onto the page. */
export function textStyle(
  pdfium: Pdfium,
  object: number,
  outer?: Matrix,
): PdfTextStyle {
  const { lib } = pdfium;
  const font = lib.FPDFTextObj_GetFont(object);
  const baseName = pdfium.readUtf8String((buffer, bytes) =>
    lib.FPDFFont_GetBaseFontName(font, buffer, bytes),
  );
  // PDFium reports the family of whatever font it loaded, which for a
  // non-embedded font is its own substitute; the document's name is the
  // one a host wants to show.
  const family = lib.FPDFFont_GetIsEmbedded(font)
    ? pdfium.readUtf8String((buffer, bytes) =>
        lib.FPDFFont_GetFamilyName(font, buffer, bytes),
      ) || declaredFamily(baseName)
    : declaredFamily(baseName);
  const flags = lib.FPDFFont_GetFlags(font);
  const weight = lib.FPDFFont_GetWeight(font);
  const size =
    pdfium.readNumbers(1, "float", ([pointer]) =>
      lib.FPDFTextObj_GetFontSize(object, pointer!),
    )?.[0] ?? 0;
  const matrix = objectMatrix(pdfium, object);
  return {
    fontFamily: family,
    fontSize: round(size * textScale(outer ? concat(matrix, outer) : matrix)),
    bold:
      weight >= 600 ||
      (flags & FLAG_FORCE_BOLD) !== 0 ||
      /bold|black|heavy/i.test(baseName),
    italic: (flags & FLAG_ITALIC) !== 0 || /italic|oblique/i.test(baseName),
    color: fillColor(pdfium, object) ?? "#000000",
  };
}

/** "ABCDEF+Helvetica-BoldOblique" → "Helvetica"; "Arial,Bold" → "Arial". */
function declaredFamily(baseName: string): string {
  const name = baseName.replace(/^[A-Z]{6}\+/, "");
  const separator = name.search(/[-,]/);
  if (separator < 0) return name;
  const suffix = name.slice(separator + 1);
  return /^(bold|italic|oblique|regular|roman|light|medium|black|heavy|mt|ps|,)*$/i.test(
    suffix.replaceAll(/[-,]/g, ""),
  )
    ? name.slice(0, separator)
    : name;
}

function shapeStyle(pdfium: Pdfium, object: number): PdfShapeStyle {
  const { lib } = pdfium;
  const mode = pdfium.readNumbers(2, "i32", ([fill, stroke]) =>
    lib.FPDFPath_GetDrawMode(object, fill!, stroke!),
  );
  const [fillMode, stroked] = mode ?? [0, 0];
  const width =
    pdfium.readNumbers(1, "float", ([pointer]) =>
      lib.FPDFPageObj_GetStrokeWidth(object, pointer!),
    )?.[0] ?? 1;
  return {
    ...(stroked
      ? {
          stroke: {
            color: strokeColor(pdfium, object) ?? "#000000",
            width: round(width),
          },
        }
      : {}),
    ...(fillMode
      ? { fill: { color: fillColor(pdfium, object) ?? "#000000" } }
      : {}),
  };
}

function fillColor(pdfium: Pdfium, object: number): string | undefined {
  const rgba = pdfium.readNumbers(4, "i32", ([r, g, b, a]) =>
    pdfium.lib.FPDFPageObj_GetFillColor(object, r!, g!, b!, a!),
  );
  return rgba ? hex(rgba) : undefined;
}

function strokeColor(pdfium: Pdfium, object: number): string | undefined {
  const rgba = pdfium.readNumbers(4, "i32", ([r, g, b, a]) =>
    pdfium.lib.FPDFPageObj_GetStrokeColor(object, r!, g!, b!, a!),
  );
  return rgba ? hex(rgba) : undefined;
}

export function hex(rgba: readonly number[]): string {
  return `#${rgba
    .slice(0, 3)
    .map((channel) =>
      Math.max(0, Math.min(255, channel)).toString(16).padStart(2, "0"),
    )
    .join("")}`;
}

/** Clockwise rotation in page space, in whole degrees, 0 when upright. */
function objectRotation(
  pdfium: Pdfium,
  object: number,
  geometry: PageGeometry,
): number {
  const own = objectMatrix(pdfium, object);
  const [a, b] = geometry.matrix ? concat(own, geometry.matrix) : own;
  // User space is y-up, so a positive angle there is counter-clockwise on
  // the displayed page; the page's own quarter turns add to it.
  const degrees = (-Math.atan2(b, a) * 180) / Math.PI + geometry.rotation * 90;
  const normalized = ((Math.round(degrees) % 360) + 360) % 360;
  return normalized;
}
