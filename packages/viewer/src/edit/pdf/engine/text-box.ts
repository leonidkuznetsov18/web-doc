import type { EditOperation, PageRect } from "../../types.js";
import type {
  InsertTextBoxOperation,
  PdfElement,
  PdfOperation,
  PdfTextAlign,
  PdfTextBoxStyle,
  ReplaceTextOperation,
  SetTextStyleOperation,
} from "../types.js";
import {
  MARK_NAME,
  MARK_PARAM,
  OBJECT_TEXT,
  type MarkParams,
} from "./elements.js";
import type { FontRequest } from "./fonts.js";
import { displayedSize, pageToUser, type PageGeometry } from "./geometry.js";
import { layoutText } from "./text-layout.js";
import { tableSpecOf, tableText, type TableSpec } from "./tables.js";
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
    validateFont(
      fontRequest(resolveStyle(operation.style ?? {}), operation.text),
      context,
      issue,
    );
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
    const drawn = drawTextBox(context, operation.pageIndex, spec);
    return {
      createdIds: [id],
      changedPages: [operation.pageIndex],
      warnings: warningsFor(drawn, id),
    };
  },
};

/** `replaceText` for a text box: a rebuild from its inputs with new text. */
export const textBoxReplaceText: OperationHandler<ReplaceTextOperation> = {
  validate(operation, context, issue) {
    const target = textBoxTarget(operation.target, context, issue);
    if (!target) return;
    validateFont(
      fontRequest(target.spec.style, operation.text),
      context,
      issue,
    );
  },
  apply(operation, context) {
    const { location, spec } = textBoxTarget(operation.target, context)!;
    return changed(
      location,
      rebuildTextBox(context, location, { ...spec, text: operation.text }),
    );
  },
};

/** `setTextStyle` for a text box: a rebuild with the merged style. */
export const textBoxSetTextStyle: OperationHandler<SetTextStyleOperation> = {
  validate(operation, context, issue) {
    const target = textBoxTarget(operation.target, context, issue);
    if (!target) return;
    const style = { ...target.spec.style, ...definedFields(operation.style) };
    validateFont(fontRequest(style, target.spec.text), context, issue);
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

export interface DrawnTextBox {
  readonly overflow: boolean;
  /** Set when the text was drawn with another font than asked for. */
  readonly substitution?: string;
}

/**
 * Lays the box out and inserts one text object per line, appended to the
 * page or at `insertAt` in drawing order.
 */
export function drawTextBox(
  context: OperationContext,
  pageIndex: number,
  spec: TextBoxSpec,
  insertAt?: number,
): DrawnTextBox {
  const { pdfium, measurer } = context;
  const { lib } = pdfium;
  const font = context.fonts.resolve(
    pdfium,
    context.document,
    fontRequest(spec.style, spec.text),
  );
  const { ascent } = measurer.metrics(font.handle, spec.style.fontSize);
  const layout = layoutText({
    text: spec.text,
    width: spec.rect.width,
    height: spec.rect.height,
    fontSize: spec.style.fontSize,
    lineHeight: spec.style.lineHeight,
    align: spec.style.align,
    ascent,
    advance: (text) => measurer.advance(font.handle, spec.style.fontSize, text),
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
        font.handle,
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
  return {
    overflow: layout.overflow,
    ...(font.substitution ? { substitution: font.substitution } : {}),
  };
}

/**
 * Replaces a text box's objects with ones drawn from `spec`, at the same
 * place in the drawing order so nothing else changes its stacking.
 */
export function rebuildTextBox(
  context: OperationContext,
  location: ElementLocation,
  spec: TextBoxSpec,
): DrawnTextBox {
  const first = removeObjects(context, location);
  return drawTextBox(context, location.pageIndex, spec, first);
}

/** Removes an element's objects from the page; returns where they started. */
export function removeObjects(
  context: OperationContext,
  location: ElementLocation,
): number {
  const { lib } = context.pdfium;
  const first = location.indexes[0]!;
  context.withHolder(location, (holder) => {
    // Highest first, so earlier indexes stay valid while removing.
    for (const index of [...location.indexes].reverse()) {
      const object = lib.FPDFPage_GetObject(holder, index);
      lib.FPDFPage_RemoveObject(holder, object);
      lib.FPDFPageObj_Destroy(object);
    }
  });
  context.spliceObjects(
    location.pageIndex,
    first,
    location.indexes.length,
    [],
    location.forms,
  );
  return first;
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
  drawn: DrawnTextBox,
): OperationResult {
  return {
    createdIds: [],
    changedPages: [location.pageIndex],
    warnings: warningsFor(drawn, location.record.id),
  };
}

function warningsFor(
  drawn: DrawnTextBox,
  elementId: string,
): OperationResult["warnings"] {
  return [
    ...(drawn.overflow
      ? [
          {
            code: "fidelity-degraded" as const,
            message: "The text does not fit the box's height and runs past it",
            details: { elementId },
          },
        ]
      : []),
    ...(drawn.substitution
      ? [
          {
            code: "font-substitution" as const,
            message: drawn.substitution,
            details: { elementId },
          },
        ]
      : []),
  ];
}

/**
 * The fonts a batch will need: what each operation draws, read from the
 * operation itself, from the text box it rebuilds, or from the text object
 * it may have to replace.
 */
export function fontRequestsOf(
  operations: readonly EditOperation[],
  markOf: (id: string) => MarkParams | undefined,
  elementOf: (id: string) => PdfElement | undefined,
): { readonly family: string; readonly text: string }[] {
  const requests: { family: string; text: string }[] = [];
  const spec = (id: string): TextBoxSpec | undefined => {
    const mark = markOf(id);
    return mark?.kind === "textBox"
      ? (mark as unknown as TextBoxSpec)
      : undefined;
  };
  const table = (id: string): TableSpec | undefined => tableSpecOf(markOf(id));
  for (const raw of operations) {
    const operation = raw as PdfOperation;
    switch (operation.op) {
      case "insertTextBox":
        requests.push({
          family: operation.style?.fontFamily ?? "Helvetica",
          text: operation.text,
        });
        break;
      case "replaceText": {
        const box = spec(operation.target);
        if (box) {
          requests.push({ family: box.style.fontFamily, text: operation.text });
          break;
        }
        const element = elementOf(operation.target);
        if (element?.kind === "text")
          requests.push({
            family: element.textStyle?.fontFamily ?? "Helvetica",
            text: operation.text,
          });
        break;
      }
      case "setTextStyle": {
        const box = spec(operation.target);
        if (box)
          requests.push({
            family: operation.style.fontFamily ?? box.style.fontFamily,
            text: box.text,
          });
        break;
      }
      case "resizeElement":
      case "moveElement": {
        const box = spec(operation.target);
        if (box)
          requests.push({ family: box.style.fontFamily, text: box.text });
        const grid = table(operation.target);
        if (grid)
          requests.push({
            family: grid.style.fontFamily,
            text: tableText(grid.rows),
          });
        break;
      }
      case "insertTable":
        requests.push({
          family: operation.style?.fontFamily ?? "Helvetica",
          text: tableText(operation.rows),
        });
        break;
      case "setTableCell": {
        const grid = table(operation.target);
        if (grid)
          requests.push({
            family: grid.style.fontFamily,
            text: `${tableText(grid.rows)}\n${operation.text}`,
          });
        break;
      }
      default:
        break;
    }
  }
  return requests;
}

export function fontRequest(
  style: TextBoxSpec["style"],
  text: string,
): FontRequest {
  return {
    family: style.fontFamily,
    bold: style.bold,
    italic: style.italic,
    text,
  };
}

export function validateFont(
  request: FontRequest,
  context: OperationContext,
  issue: Issue,
): void {
  const problem = context.fonts.problem(request);
  if (problem) issue(problem.path, problem.code, problem.message);
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
