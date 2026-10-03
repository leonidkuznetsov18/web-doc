import { ViewerError } from "../../../errors.js";
import type { TextRange } from "../../types.js";
import type {
  PdfElement,
  PdfTextStyle,
  ReplaceTextOperation,
  SetTextStyleOperation,
} from "../types.js";
import {
  paragraphSetTextStyle,
  replaceParagraphText,
} from "./paragraph-edit.js";
import {
  objectBounds,
  OBJECT_TEXT,
  OBJECT_PATH,
  textScale,
  clearWebDocMark,
  writeWebDocMark,
  type MarkParams,
} from "./elements.js";
import { fontCanRewrite, validateScript, type ResolvedFont } from "./fonts.js";
import type {
  ElementLocation,
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";
import { pageToUser } from "./geometry.js";
import { createUnderline } from "./text-decoration.js";
import type { Pdfium } from "./pdfium.js";
import {
  parseColor,
  setText,
  textBoxReplaceText,
  textBoxSetTextStyle,
  textBoxTarget,
} from "./text-box.js";

/*
 * Edits to text objects that already exist in the file. The object is changed
 * in place only when its font provably covers the new text; otherwise it is
 * replaced by a new object in a font that does, at the same place, size and
 * colour, and the receipt says so.
 */

export const replaceText: OperationHandler<ReplaceTextOperation> = {
  validate(operation, context, issue) {
    if (context.paragraph(operation.target)?.paragraph.id === operation.target)
      return replaceParagraphText.validate(
        { ...operation, op: "replaceParagraphText" },
        context,
        issue,
      );
    const box = textBoxTarget(operation.target, context);
    if (box) {
      const spliced = splice(box.spec.text, operation, issue);
      if (spliced === undefined) return;
      return textBoxReplaceText.validate(
        wholeTextOperation(operation, spliced),
        context,
        issue,
      );
    }
    const target = textTarget(operation.target, context, issue);
    if (!target) return;
    const whole = splice(target.element.text ?? "", operation, issue);
    if (whole === undefined) return;
    if (!validateScript(operation.text, issue)) return;
    if (canKeepFont(context, target, whole)) return;
    const request = {
      family: target.element.textStyle?.fontFamily ?? "Helvetica",
      bold: target.element.textStyle?.bold ?? false,
      italic: target.element.textStyle?.italic ?? false,
      text: operation.text,
    };
    if (request.bold || request.italic) {
      const font = context.fonts.resolveStyle(
        context.pdfium,
        context.document,
        request,
      );
      if ("code" in font) issue(font.path, font.code, font.message);
      return;
    }
    // The file's own family is only a hint: what matters is whether some
    // available font can draw the new text.
    let problem = context.fonts.problem(request);
    if (problem?.code === "unknown-font")
      problem = context.fonts.problem({ ...request, family: "Helvetica" });
    if (problem) issue(problem.path, problem.code, problem.message);
  },
  apply(operation, context) {
    if (context.paragraph(operation.target)?.paragraph.id === operation.target)
      return replaceParagraphText.apply(
        { ...operation, op: "replaceParagraphText" },
        context,
      );
    const box = textBoxTarget(operation.target, context);
    if (box)
      return textBoxReplaceText.apply(
        wholeTextOperation(operation, splice(box.spec.text, operation)!),
        context,
      );
    const target = textTarget(operation.target, context)!;
    const underlined = target.element.textStyle?.underline === true;
    if (underlined) removeUnderline(context, target);
    const changed = replaceNativeText(operation, context, target);
    if (underlined)
      for (const id of [operation.target, ...changed.createdIds])
        addUnderline(context, id);
    return changed;
  },
};

function replaceNativeText(
  operation: ReplaceTextOperation,
  context: OperationContext,
  target: TextTarget,
): OperationResult {
  const previous = target.element.text ?? "";
  const whole = splice(previous, operation)!;
  if (canKeepFont(context, target, whole)) {
    const kept = replaceInPlace(context, target, whole);
    if (kept) return result(target.location);
  }
  const span = spanOf(operation, previous);
  const before = previous.slice(0, span.start);
  const after = previous.slice(span.end);
  if (!before && !after)
    return replaceWithFallback(context, target, operation.text);
  return (
    splitAround(context, target, before, operation.text, after) ??
    replaceWithFallback(context, target, whole)
  );
}

/** The same operation as a whole-text replacement with `text`. */
function wholeTextOperation(
  operation: ReplaceTextOperation,
  text: string,
): ReplaceTextOperation {
  const { range: _range, ...rest } = operation;
  return { ...rest, text };
}

/** The target's text with the operation's range replaced; the whole text without one. */
function splice(
  previous: string,
  operation: ReplaceTextOperation,
  issue: Issue = () => {},
): string | undefined {
  const { range } = operation;
  if (!range) return operation.text;
  const onTarget = (position: TextRange["start"]): boolean =>
    position.elementId === operation.target &&
    Number.isInteger(position.offset) &&
    position.offset >= 0 &&
    position.offset <= previous.length;
  if (
    !onTarget(range.start) ||
    !onTarget(range.end) ||
    range.start.offset > range.end.offset
  ) {
    issue(
      "/range",
      "invalid-range",
      `The range must lie inside the target's text (0–${previous.length})`,
    );
    return undefined;
  }
  const { start, end } = spanOf(operation, previous);
  return previous.slice(0, start) + operation.text + previous.slice(end);
}

function spanOf(
  operation: ReplaceTextOperation,
  previous: string,
): { readonly start: number; readonly end: number } {
  if (!operation.range) return { start: 0, end: previous.length };
  return {
    start: operation.range.start.offset,
    end: operation.range.end.offset,
  };
}

export const setTextStyle: OperationHandler<SetTextStyleOperation> = {
  validate(operation, context, issue) {
    if (context.paragraph(operation.target)?.paragraph.id === operation.target)
      return paragraphSetTextStyle.validate(operation, context, issue);
    if (textBoxTarget(operation.target, context))
      return textBoxSetTextStyle.validate(operation, context, issue);
    const target = textTarget(operation.target, context, issue);
    if (!target) return;
    for (const field of ["fontFamily", "align", "lineHeight"] as const)
      if (operation.style[field] !== undefined)
        issue(
          `/style/${field}`,
          "unsupported-style",
          `Existing text accepts color, fontSize, bold, italic and underline; ${field} cannot`,
        );
    const style = target.element.textStyle;
    if (
      style &&
      (faceChanges(style, operation) || operation.style.underline === true)
    ) {
      const filled = context.readPage(target.location.pageIndex, (page) => {
        const index = target.location.indexes[0];
        return (
          index !== undefined &&
          context.pdfium.lib.FPDFTextObj_GetTextRenderMode(
            context.pdfium.lib.FPDFPage_GetObject(page, index),
          ) === 0
        );
      });
      if (!filled) {
        issue(
          "/style",
          "unsupported-style",
          "Bold, italic and underline require filled text",
        );
        return;
      }
    }
    if (style && faceChanges(style, operation)) {
      const font = context.fonts.resolveStyle(
        context.pdfium,
        context.document,
        {
          family: style.fontFamily,
          text: target.element.text ?? "",
          bold: operation.style.bold ?? style.bold,
          italic: operation.style.italic ?? style.italic,
        },
      );
      if ("code" in font) issue(font.path, font.code, font.message);
    }
  },
  apply(operation, context) {
    if (context.paragraph(operation.target)?.paragraph.id === operation.target)
      return paragraphSetTextStyle.apply(operation, context);
    if (textBoxTarget(operation.target, context))
      return textBoxSetTextStyle.apply(operation, context);
    const target = textTarget(operation.target, context)!;
    const { lib } = context.pdfium;
    const { location } = target;
    const underlined =
      operation.style.underline ?? target.element.textStyle?.underline ?? false;
    if (target.element.textStyle?.underline) removeUnderline(context, target);
    let changedFont: OperationResult | undefined;
    const style = target.element.textStyle;
    if (style && faceChanges(style, operation)) {
      const next = {
        ...style,
        bold: operation.style.bold ?? style.bold,
        italic: operation.style.italic ?? style.italic,
      };
      const font = context.fonts.resolveStyle(
        context.pdfium,
        context.document,
        {
          family: next.fontFamily,
          text: target.element.text ?? "",
          bold: next.bold,
          italic: next.italic,
        },
      );
      if ("code" in font)
        throw new ViewerError("invalid-operation", font.message);
      changedFont = replaceWithFallback(
        context,
        target,
        target.element.text ?? "",
        { font, style: next },
      );
    }
    if (operation.style.color !== undefined) {
      const [r, g, b] = parseColor(operation.style.color);
      context.withPage(location.pageIndex, (page) => {
        lib.FPDFPageObj_SetFillColor(
          lib.FPDFPage_GetObject(page, location.indexes[0]!),
          r,
          g,
          b,
          255,
        );
      });
    }
    if (operation.style.fontSize !== undefined)
      resize(context, target, operation.style.fontSize);
    if (underlined) addUnderline(context, operation.target);
    return changedFont ?? result(location);
  },
};

function removeUnderline(context: OperationContext, target: TextTarget): void {
  if (target.location.record.mark?.kind !== "text") return;
  const { pdfium } = context;
  const { lib } = pdfium;
  const { pageIndex, indexes } = target.location;
  const [first] = indexes;
  if (first === undefined) return;
  context.withPage(pageIndex, (page) => {
    for (const index of indexes.slice(1).reverse()) {
      const path = lib.FPDFPage_GetObject(page, index);
      lib.FPDFPage_RemoveObject(page, path);
      lib.FPDFPageObj_Destroy(path);
      context.spliceObjects(pageIndex, index, 1, []);
    }
    clearWebDocMark(pdfium, lib.FPDFPage_GetObject(page, first));
    context.spliceObjects(pageIndex, first, 1, [
      { id: target.element.id, type: OBJECT_TEXT },
    ]);
  });
}

function addUnderline(context: OperationContext, id: string): void {
  const location = context.locate(id);
  const first = location?.indexes[0];
  if (!location || first === undefined)
    throw new ViewerError(
      "edit-failed",
      "The text to underline no longer exists",
    );
  const { pdfium } = context;
  const { lib } = pdfium;
  context.withPage(location.pageIndex, (page) => {
    const object = lib.FPDFPage_GetObject(page, first);
    const textPage = lib.FPDFText_LoadPage(page);
    let path: number | undefined;
    try {
      path = createUnderline(pdfium, object, textPage);
    } finally {
      lib.FPDFText_ClosePage(textPage);
    }
    if (!path)
      throw new ViewerError(
        "invalid-operation",
        "This text has no drawable underline geometry",
      );
    const mark: MarkParams = { kind: "text", id, underline: true };
    writeWebDocMark(pdfium, context.document, object, mark);
    writeWebDocMark(pdfium, context.document, path, mark);
    lib.FPDFPage_InsertObjectAtIndex(page, path, first + 1);
    context.spliceObjects(location.pageIndex, first, 1, [
      { id, type: OBJECT_TEXT, mark },
      { id, type: OBJECT_PATH, mark },
    ]);
  });
}

function faceChanges(
  style: PdfTextStyle,
  operation: SetTextStyleOperation,
): boolean {
  return (
    (operation.style.bold !== undefined &&
      operation.style.bold !== style.bold) ||
    (operation.style.italic !== undefined &&
      operation.style.italic !== style.italic)
  );
}

interface TextTarget {
  readonly location: ElementLocation;
  readonly element: PdfElement;
}

function textTarget(
  target: string,
  context: OperationContext,
  issue: Issue = () => {},
): TextTarget | undefined {
  const location = context.locate(target);
  const element = location && context.element(target);
  if (!location || !element) {
    issue("/target", "unknown-target", `No element ${target}`);
    return undefined;
  }
  if (element.kind !== "text" || location.record.type !== OBJECT_TEXT) {
    issue(
      "/target",
      "unsupported-target",
      "Only text objects and text boxes carry text",
    );
    return undefined;
  }
  return { location, element };
}

/**
 * Whether the object's own font can draw `text`: a standard font for WinAnsi
 * text, or an embedded TrueType or bare CFF program with a glyph for every
 * character (which also answers for subsets, since they hold only what they
 * kept), each with a width in the PDF.
 */
function canKeepFont(
  context: OperationContext,
  target: TextTarget,
  text: string,
): boolean {
  return context.readPage(target.location.pageIndex, (page) => {
    const object = context.pdfium.lib.FPDFPage_GetObject(
      page,
      target.location.indexes[0]!,
    );
    return fontCanRewrite(
      context.pdfium,
      context.pdfium.lib.FPDFTextObj_GetFont(object),
      text,
    );
  });
}

/** Sets the text and confirms PDFium reads it back; false means it must be replaced. */
function replaceInPlace(
  context: OperationContext,
  target: TextTarget,
  text: string,
): boolean {
  const { pdfium } = context;
  const { lib } = pdfium;
  const index = target.location.indexes[0]!;
  const previous = target.element.text ?? "";
  context.withPage(target.location.pageIndex, (page) => {
    setText(pdfium, lib.FPDFPage_GetObject(page, index), text);
  });
  const faithful = context.withPage(target.location.pageIndex, (page) => {
    const textPage = lib.FPDFText_LoadPage(page);
    try {
      const object = lib.FPDFPage_GetObject(page, index);
      const readBack = pdfium.readWideString((buffer, bytes) =>
        lib.FPDFTextObj_GetText(object, textPage, buffer, bytes),
      );
      if (readBack === text) return true;
      // Reading-order extraction can append a generated gap before the next
      // object. Verify the object's own characters instead, retaining every
      // authored space: trimming would also hide a lost deliberate space.
      let authored = "";
      const count = lib.FPDFText_CountChars(textPage);
      for (let at = 0; at < count; at += 1) {
        if (
          lib.FPDFText_GetTextObject(textPage, at) === object &&
          lib.FPDFText_IsGenerated(textPage, at) === 0
        )
          authored += String.fromCodePoint(
            lib.FPDFText_GetUnicode(textPage, at),
          );
      }
      return authored === text;
    } finally {
      lib.FPDFText_ClosePage(textPage);
    }
  });
  if (faithful) return true;
  // Put the old text back so the fallback path starts from a known state.
  context.withPage(target.location.pageIndex, (page) => {
    setText(pdfium, lib.FPDFPage_GetObject(page, index), previous);
  });
  return false;
}

/**
 * Splits the object around a replaced span: the parts before and after keep
 * the object's font, size, colour and baseline, the middle part is drawn in
 * a font that covers it, and the parts follow each other by their advances.
 * The first part keeps the element's id. Returns nothing when a part cannot
 * be written back in the original font, leaving the page as it was.
 */
function splitAround(
  context: OperationContext,
  target: TextTarget,
  before: string,
  middle: string,
  after: string,
): OperationResult | undefined {
  const { pdfium, measurer } = context;
  const { lib } = pdfium;
  const { location, element } = target;
  const style = element.textStyle!;
  const index = location.indexes[0]!;
  const fallback = replacementFont(context, style, middle);
  const createdIds: string[] = [];
  const written = context.withPage(location.pageIndex, (page) => {
    const old = lib.FPDFPage_GetObject(page, index);
    const oldFont = lib.FPDFTextObj_GetFont(old);
    const matrix = pdfium.readNumbers(6, "float", ([pointer]) =>
      lib.FPDFPageObj_GetMatrix(old, pointer!),
    ) ?? [1, 0, 0, 1, 0, 0];
    const size =
      pdfium.readNumbers(1, "float", ([pointer]) =>
        lib.FPDFTextObj_GetFontSize(old, pointer!),
      )?.[0] ?? style.fontSize / textScale(matrix);
    const [a, b, c, d, e, f] = matrix as [
      number,
      number,
      number,
      number,
      number,
      number,
    ];
    const [r, g, bl] = parseColor(style.color);
    const parts = [
      { text: before, font: oldFont },
      { text: middle, font: fallback.handle },
      { text: after, font: oldFont },
    ].filter((part) => part.text.length > 0);
    const objects: number[] = [];
    let cursor = 0;
    for (const part of parts) {
      const object = lib.FPDFPageObj_CreateTextObj(
        context.document,
        part.font,
        size,
      );
      if (object) {
        setText(pdfium, object, part.text);
        lib.FPDFPageObj_SetFillColor(object, r, g, bl, 255);
        lib.FPDFPageObj_Transform(
          object,
          a,
          b,
          c,
          d,
          e + cursor * a,
          f + cursor * b,
        );
      }
      objects.push(object);
      cursor += measurer.advance(part.font, size, part.text);
    }
    // Every part must read back as written, or the original font cannot
    // encode its share and the whole object goes to the fallback instead.
    if (objects.some((object) => object === 0)) {
      for (const object of objects) if (object) lib.FPDFPageObj_Destroy(object);
      return false;
    }
    lib.FPDFPage_RemoveObject(page, old);
    objects.forEach((object, at) =>
      lib.FPDFPage_InsertObjectAtIndex(page, object, index + at),
    );
    const textPage = lib.FPDFText_LoadPage(page);
    let faithful = true;
    try {
      objects.forEach((object, at) => {
        const read = pdfium.readWideString((buffer, bytes) =>
          lib.FPDFTextObj_GetText(object, textPage, buffer, bytes),
        );
        // PDFium appends a generated space to an object a gap follows.
        if (read.trimEnd() !== parts[at]!.text.trimEnd()) faithful = false;
      });
    } finally {
      lib.FPDFText_ClosePage(textPage);
    }
    if (!faithful) {
      for (const object of objects) {
        lib.FPDFPage_RemoveObject(page, object);
        lib.FPDFPageObj_Destroy(object);
      }
      lib.FPDFPage_InsertObjectAtIndex(page, old, index);
      return false;
    }
    lib.FPDFPageObj_Destroy(old);
    return true;
  });
  if (!written) return undefined;
  const partCount = [before, middle, after].filter(Boolean).length;
  const records = Array.from({ length: partCount }, (_, at) => {
    if (at === 0) return { id: location.record.id, type: OBJECT_TEXT };
    const id = context.newId(location.pageIndex);
    createdIds.push(id);
    return { id, type: OBJECT_TEXT };
  });
  context.spliceObjects(location.pageIndex, index, 1, records);
  return {
    createdIds,
    changedPages: [location.pageIndex],
    warnings: [
      {
        code: "font-substitution",
        message:
          fallback.substitution ??
          `${style.fontFamily} cannot draw the new text; ${fallback.family} is used for it`,
        details: { elementId: location.record.id },
      },
    ],
  };
}

function replacementFont(
  context: OperationContext,
  style: PdfTextStyle,
  text: string,
): ResolvedFont {
  const request = {
    family: style.fontFamily,
    bold: style.bold,
    italic: style.italic,
    text,
  };
  if (!style.bold && !style.italic)
    return context.fonts.resolve(context.pdfium, context.document, request);
  const font = context.fonts.resolveStyle(
    context.pdfium,
    context.document,
    request,
  );
  if ("code" in font) throw new ViewerError("invalid-operation", font.message);
  return font;
}

/** Replaces the object with one in a covering font at the same place, size and colour. */
function replaceWithFallback(
  context: OperationContext,
  target: TextTarget,
  text: string,
  change?: { readonly font: ResolvedFont; readonly style: PdfTextStyle },
): OperationResult {
  const { pdfium } = context;
  const { lib } = pdfium;
  const { location, element } = target;
  const style = change?.style ?? element.textStyle!;
  const index = location.indexes[0]!;
  const font = change?.font ?? replacementFont(context, style, text);
  context.withPage(location.pageIndex, (page) => {
    const old = lib.FPDFPage_GetObject(page, index);
    const matrix = pdfium.readNumbers(6, "float", ([pointer]) =>
      lib.FPDFPageObj_GetMatrix(old, pointer!),
    ) ?? [1, 0, 0, 1, 0, 0];
    const size =
      pdfium.readNumbers(1, "float", ([pointer]) =>
        lib.FPDFTextObj_GetFontSize(old, pointer!),
      )?.[0] ?? style.fontSize / textScale(matrix);
    const object = lib.FPDFPageObj_CreateTextObj(
      context.document,
      font.handle,
      size,
    );
    setText(pdfium, object, text);
    const [r, g, b] = parseColor(style.color);
    const alpha =
      pdfium.readNumbers(
        4,
        "i32",
        ([red, green, blue, opacity]) =>
          red !== undefined &&
          green !== undefined &&
          blue !== undefined &&
          opacity !== undefined &&
          lib.FPDFPageObj_GetFillColor(old, red, green, blue, opacity),
      )?.[3] ?? 255;
    lib.FPDFPageObj_SetFillColor(object, r, g, b, alpha);
    const [a, bb, c, d, e, f] = matrix as [
      number,
      number,
      number,
      number,
      number,
      number,
    ];
    lib.FPDFPageObj_Transform(object, a, bb, c, d, e, f);
    lib.FPDFPage_RemoveObject(page, old);
    lib.FPDFPageObj_Destroy(old);
    lib.FPDFPage_InsertObjectAtIndex(page, object, index);
  });
  context.spliceObjects(location.pageIndex, index, 1, [
    { id: location.record.id, type: OBJECT_TEXT },
  ]);
  return {
    createdIds: [],
    changedPages: [location.pageIndex],
    warnings:
      change && !font.substitution
        ? []
        : [
            {
              code: "font-substitution",
              message:
                font.substitution ??
                `${style.fontFamily} cannot draw the new text; ${font.family} is used`,
              details: { elementId: location.record.id },
            },
          ],
  };
}

/**
 * Gives the object a new font size by rebuilding it with the same font, text,
 * colour and matrix, then moving it so its top-left corner stays put.
 */
function resize(
  context: OperationContext,
  target: TextTarget,
  fontSize: number,
): void {
  const { pdfium } = context;
  const { lib } = pdfium;
  const { location } = target;
  const index = location.indexes[0]!;
  const geometry = context.geometry(location.pageIndex);
  const before = target.element.bounds;
  context.withPage(location.pageIndex, (page) => {
    const old = lib.FPDFPage_GetObject(page, index);
    const textPage = lib.FPDFText_LoadPage(page);
    let text: string;
    try {
      text = pdfium.readWideString((buffer, bytes) =>
        lib.FPDFTextObj_GetText(old, textPage, buffer, bytes),
      );
    } finally {
      lib.FPDFText_ClosePage(textPage);
    }
    const matrix = pdfium.readNumbers(6, "float", ([pointer]) =>
      lib.FPDFPageObj_GetMatrix(old, pointer!),
    ) ?? [1, 0, 0, 1, 0, 0];
    const color = pdfium.readNumbers(4, "i32", ([r, g, b, a]) =>
      lib.FPDFPageObj_GetFillColor(old, r!, g!, b!, a!),
    ) ?? [0, 0, 0, 255];
    // The size is the one the text shows: the old matrix, applied below,
    // scales it again, so the font gets the size undone by that scale.
    const object = lib.FPDFPageObj_CreateTextObj(
      context.document,
      lib.FPDFTextObj_GetFont(old),
      fontSize / textScale(matrix),
    );
    setText(pdfium, object, text);
    lib.FPDFPageObj_SetFillColor(
      object,
      color[0]!,
      color[1]!,
      color[2]!,
      color[3]!,
    );
    const [a, b, c, d, e, f] = matrix as [
      number,
      number,
      number,
      number,
      number,
      number,
    ];
    lib.FPDFPageObj_Transform(object, a, b, c, d, e, f);
    lib.FPDFPage_RemoveObject(page, old);
    lib.FPDFPageObj_Destroy(old);
    lib.FPDFPage_InsertObjectAtIndex(page, object, index);
    // Anchor the top-left corner where it was.
    const after = objectBounds(pdfium, object, geometry);
    if (after) {
      const from = pageToUser(geometry, 0, 0);
      const to = pageToUser(geometry, before.x - after.x, before.y - after.y);
      lib.FPDFPageObj_Transform(
        object,
        1,
        0,
        0,
        1,
        to.x - from.x,
        to.y - from.y,
      );
    }
  });
  context.spliceObjects(location.pageIndex, index, 1, [
    { id: location.record.id, type: OBJECT_TEXT },
  ]);
}

function result(location: ElementLocation): OperationResult {
  return { createdIds: [], changedPages: [location.pageIndex], warnings: [] };
}
