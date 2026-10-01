import type { PageRect } from "../../types.js";
import type {
  InsertTextBoxOperation,
  PdfTextAlign,
  PdfTextBoxStyle,
  ReplaceTextOperation,
  SetTextStyleOperation,
} from "../types.js";
import { MARK_NAME, MARK_PARAM, OBJECT_TEXT } from "./elements.js";
import {
  firstNonWinAnsi,
  isStandardFamily,
  standardFontName,
} from "./fonts.js";
import { displayedSize, pageToUser, type PageGeometry } from "./geometry.js";
import { layoutText } from "./text-layout.js";
import type {
  ElementLocation,
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";
import type { Pdfium } from "./pdfium.js";

/** A text box's inputs with every default filled in; stored in its mark. */
export type TextBoxSpec = {
  readonly kind: "textBox";
  readonly id: string;
  readonly rect: PageRect;
  readonly text: string;
  readonly style: {
    readonly fontFamily: string;
    readonly fontSize: number;
    readonly bold: boolean;
    readonly italic: boolean;
    readonly color: string;
    readonly align: PdfTextAlign;
    readonly lineHeight: number;
  };
};

export const insertTextBox: OperationHandler<InsertTextBoxOperation> = {
  validate(operation, context, issue) {
    if (operation.pageIndex >= context.pageCount) {
      issue("/pageIndex", "unknown-target", `No page ${operation.pageIndex}`);
      return;
    }
    validateRect(operation.rect, context.geometry(operation.pageIndex), issue);
    validateTextStyle(operation.style ?? {}, issue);
    const bad = firstNonWinAnsi(operation.text);
    if (bad !== undefined) issue("/text", "font-unavailable", unencodable(bad));
  },

  apply(operation, context) {
    const id = context.newId(operation.pageIndex);
    const spec: TextBoxSpec = {
      kind: "textBox",
      id,
      rect: operation.rect,
      text: operation.text,
      style: resolveStyle(operation.style ?? {}),
    };
    const overflow = drawTextBox(context, operation.pageIndex, spec);
    return {
      createdIds: [id],
      changedPages: [operation.pageIndex],
      warnings: overflowWarnings(overflow, id),
    };
  },
};

export const replaceText: OperationHandler<ReplaceTextOperation> = {
  validate(operation, context, issue) {
    if (!textBoxTarget(operation.target, context, issue)) return;
    const bad = firstNonWinAnsi(operation.text);
    if (bad !== undefined) issue("/text", "font-unavailable", unencodable(bad));
  },
  apply(operation, context) {
    const { location, spec } = textBoxTarget(operation.target, context)!;
    return changed(
      location,
      rebuildTextBox(context, location, { ...spec, text: operation.text }),
    );
  },
};

export const setTextStyle: OperationHandler<SetTextStyleOperation> = {
  validate(operation, context, issue) {
    if (!textBoxTarget(operation.target, context, issue)) return;
    validateTextStyle(operation.style, issue);
  },
  apply(operation, context) {
    const { location, spec } = textBoxTarget(operation.target, context)!;
    return changed(
      location,
      rebuildTextBox(context, location, {
        ...spec,
        style: { ...spec.style, ...definedFields(operation.style) },
      }),
    );
  },
};

/**
 * Lays the box out and inserts one text object per line, appended to the
 * page or at `insertAt` in drawing order. Returns whether it overflowed.
 */
export function drawTextBox(
  context: OperationContext,
  pageIndex: number,
  spec: TextBoxSpec,
  insertAt?: number,
): boolean {
  const { pdfium, measurer } = context;
  const { lib } = pdfium;
  const fontName = standardFontName(
    spec.style.fontFamily,
    spec.style.bold,
    spec.style.italic,
  )!;
  const font = lib.FPDFText_LoadStandardFont(context.document, fontName);
  const { ascent } = measurer.metrics(font, spec.style.fontSize);
  const layout = layoutText({
    text: spec.text,
    width: spec.rect.width,
    height: spec.rect.height,
    fontSize: spec.style.fontSize,
    lineHeight: spec.style.lineHeight,
    align: spec.style.align,
    ascent,
    advance: (text) => measurer.advance(font, spec.style.fontSize, text),
  });
  const geometry = context.geometry(pageIndex);
  const [r, g, b] = parseColor(spec.style.color);
  const params = JSON.stringify(spec);
  context.withPage(pageIndex, (page) => {
    const records = [];
    let index = insertAt ?? -1;
    for (const line of layout.lines) {
      const object = lib.FPDFPageObj_CreateTextObj(
        context.document,
        font,
        spec.style.fontSize,
      );
      setText(pdfium, object, line.text);
      lib.FPDFPageObj_SetFillColor(object, r, g, b, 255);
      placeUpright(
        pdfium,
        object,
        geometry,
        spec.rect.x + line.x,
        spec.rect.y + line.baseline,
      );
      const mark = lib.FPDFPageObj_AddMark(object, MARK_NAME);
      lib.FPDFPageObjMark_SetStringParam(
        context.document,
        object,
        mark,
        MARK_PARAM,
        params,
      );
      if (insertAt === undefined) lib.FPDFPage_InsertObject(page, object);
      else lib.FPDFPage_InsertObjectAtIndex(page, object, index++);
      records.push({ id: spec.id, type: OBJECT_TEXT, mark: spec });
    }
    if (insertAt === undefined) context.appendObjects(pageIndex, records);
    else context.spliceObjects(pageIndex, insertAt, 0, records);
  });
  return layout.overflow;
}

/**
 * Replaces a text box's objects with ones drawn from `spec`, at the same
 * place in the drawing order so nothing else changes its stacking.
 */
export function rebuildTextBox(
  context: OperationContext,
  location: ElementLocation,
  spec: TextBoxSpec,
): boolean {
  const { lib } = context.pdfium;
  const first = location.indexes[0]!;
  context.withPage(location.pageIndex, (page) => {
    // Highest first, so earlier indexes stay valid while removing.
    for (const index of [...location.indexes].reverse()) {
      const object = lib.FPDFPage_GetObject(page, index);
      lib.FPDFPage_RemoveObject(page, object);
      lib.FPDFPageObj_Destroy(object);
    }
  });
  context.spliceObjects(location.pageIndex, first, location.indexes.length, []);
  return drawTextBox(context, location.pageIndex, spec, first);
}

/** The location and stored inputs of a text box, or the issue that stops the edit. */
export function textBoxTarget(
  target: string,
  context: OperationContext,
  issue: Issue = () => {},
):
  | { readonly location: ElementLocation; readonly spec: TextBoxSpec }
  | undefined {
  const location = context.locate(target);
  if (!location) {
    issue("/target", "unknown-target", `No element ${target}`);
    return undefined;
  }
  if (location.record.mark?.kind !== "textBox") {
    issue(
      "/target",
      "unsupported-target",
      "Only text boxes created by web-doc can be edited yet",
    );
    return undefined;
  }
  return { location, spec: location.record.mark as unknown as TextBoxSpec };
}

export function changed(
  location: ElementLocation,
  overflow: boolean,
): OperationResult {
  return {
    createdIds: [],
    changedPages: [location.pageIndex],
    warnings: overflowWarnings(overflow, location.record.id),
  };
}

function overflowWarnings(
  overflow: boolean,
  elementId: string,
): OperationResult["warnings"] {
  return overflow
    ? [
        {
          code: "fidelity-degraded",
          message: "The text does not fit the box's height and runs past it",
          details: { elementId },
        },
      ]
    : [];
}

/**
 * Positions an object so it reads upright on the displayed page: the object's
 * origin goes to the page-space point and it turns against the page's /Rotate.
 */
export function placeUpright(
  pdfium: Pdfium,
  object: number,
  geometry: PageGeometry,
  pageX: number,
  pageY: number,
): void {
  const [a, b, c, d] = ROTATIONS[geometry.rotation % 4]!;
  const origin = pageToUser(geometry, pageX, pageY);
  pdfium.lib.FPDFPageObj_Transform(object, a, b, c, d, origin.x, origin.y);
}

/** Counter-clockwise quarter turns in user space, which cancel the page's clockwise /Rotate. */
const ROTATIONS: readonly (readonly [number, number, number, number])[] = [
  [1, 0, 0, 1],
  [0, 1, -1, 0],
  [-1, 0, 0, -1],
  [0, -1, 1, 0],
];

export function setText(pdfium: Pdfium, object: number, text: string): void {
  const wide = pdfium.writeWideString(text);
  try {
    pdfium.lib.FPDFText_SetText(object, wide);
  } finally {
    pdfium.free(wide);
  }
}

export function parseColor(color: string): [number, number, number] {
  return [1, 3, 5].map((offset) =>
    Number.parseInt(color.slice(offset, offset + 2), 16),
  ) as [number, number, number];
}

export function resolveStyle(style: PdfTextBoxStyle): TextBoxSpec["style"] {
  return {
    fontFamily: style.fontFamily ?? "Helvetica",
    fontSize: style.fontSize ?? 12,
    bold: style.bold ?? false,
    italic: style.italic ?? false,
    color: style.color ?? "#000000",
    align: style.align ?? "left",
    lineHeight: style.lineHeight ?? 1.2,
  };
}

function definedFields(style: PdfTextBoxStyle): Partial<TextBoxSpec["style"]> {
  return Object.fromEntries(
    Object.entries(style).filter(([, value]) => value !== undefined),
  ) as Partial<TextBoxSpec["style"]>;
}

export function validateTextStyle(style: PdfTextBoxStyle, issue: Issue): void {
  if (style.fontFamily !== undefined && !isStandardFamily(style.fontFamily))
    issue(
      "/style/fontFamily",
      "unknown-font",
      `Unknown font family ${style.fontFamily}; Helvetica, Times and Courier are available`,
    );
}

export function validateRect(
  rect: PageRect,
  geometry: PageGeometry,
  issue: Issue,
  path = "/rect",
): void {
  const size = displayedSize(geometry);
  const tolerance = 0.01;
  if (
    rect.x < -tolerance ||
    rect.y < -tolerance ||
    rect.x + rect.width > size.width + tolerance ||
    rect.y + rect.height > size.height + tolerance
  )
    issue(
      path,
      "range",
      `The rectangle must lie within the ${size.width}×${size.height} pt page`,
    );
}

function unencodable(character: string): string {
  const code = character.codePointAt(0)!.toString(16).toUpperCase();
  return `No available font can draw "${character}" (U+${code.padStart(4, "0")})`;
}
