import type {
  PdfElement,
  ReplaceTextOperation,
  SetTextStyleOperation,
} from "../types.js";
import { objectBounds, OBJECT_TEXT } from "./elements.js";
import {
  firstNonWinAnsi,
  isStandardFamily,
  parseCmap,
  type CmapCoverage,
} from "./fonts.js";
import type {
  ElementLocation,
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";
import { pageToUser } from "./geometry.js";
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
    if (textBoxTarget(operation.target, context))
      return textBoxReplaceText.validate(operation, context, issue);
    const target = textTarget(operation.target, context, issue);
    if (!target) return;
    if (!validateScript(operation.text, issue)) return;
    if (canKeepFont(context, target, operation.text)) return;
    const request = {
      family: target.element.textStyle?.fontFamily ?? "Helvetica",
      bold: target.element.textStyle?.bold ?? false,
      italic: target.element.textStyle?.italic ?? false,
      text: operation.text,
    };
    // The file's own family is only a hint: what matters is whether some
    // available font can draw the new text.
    let problem = context.fonts.problem(request);
    if (problem?.code === "unknown-font")
      problem = context.fonts.problem({ ...request, family: "Helvetica" });
    if (problem) issue(problem.path, problem.code, problem.message);
  },
  apply(operation, context) {
    if (textBoxTarget(operation.target, context))
      return textBoxReplaceText.apply(operation, context);
    const target = textTarget(operation.target, context)!;
    if (canKeepFont(context, target, operation.text)) {
      const kept = replaceInPlace(context, target, operation.text);
      if (kept) return result(target.location);
    }
    return replaceWithFallback(context, target, operation.text);
  },
};

export const setTextStyle: OperationHandler<SetTextStyleOperation> = {
  validate(operation, context, issue) {
    if (textBoxTarget(operation.target, context))
      return textBoxSetTextStyle.validate(operation, context, issue);
    if (!textTarget(operation.target, context, issue)) return;
    for (const field of [
      "fontFamily",
      "bold",
      "italic",
      "align",
      "lineHeight",
    ] as const)
      if (operation.style[field] !== undefined)
        issue(
          `/style/${field}`,
          "unsupported-style",
          `Only color and fontSize can change on existing text; ${field} cannot`,
        );
  },
  apply(operation, context) {
    if (textBoxTarget(operation.target, context))
      return textBoxSetTextStyle.apply(operation, context);
    const target = textTarget(operation.target, context)!;
    const { lib } = context.pdfium;
    const { location } = target;
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
    return result(location);
  },
};

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

/** Right-to-left scripts need shaping the MVP does not do. */
function validateScript(text: string, issue: Issue): boolean {
  for (const character of text) {
    const code = character.codePointAt(0)!;
    if (
      (code >= 0x0590 && code <= 0x08ff) ||
      (code >= 0xfb1d && code <= 0xfdff) ||
      (code >= 0xfe70 && code <= 0xfeff)
    ) {
      issue(
        "/text",
        "unsupported-script",
        `Right-to-left text such as "${character}" cannot be edited yet`,
      );
      return false;
    }
  }
  return true;
}

/**
 * Whether the object's own font can draw `text`: a standard font for WinAnsi
 * text, or an embedded TrueType program whose cmap covers every character
 * (which also answers for subsets, since their cmaps hold only what they
 * kept).
 */
function canKeepFont(
  context: OperationContext,
  target: TextTarget,
  text: string,
): boolean {
  const { pdfium } = context;
  const { lib } = pdfium;
  return context.withPage(target.location.pageIndex, (page) => {
    const object = lib.FPDFPage_GetObject(page, target.location.indexes[0]!);
    const font = lib.FPDFTextObj_GetFont(object);
    if (!lib.FPDFFont_GetIsEmbedded(font)) {
      const base = pdfium.readUtf8String((buffer, bytes) =>
        lib.FPDFFont_GetBaseFontName(font, buffer, bytes),
      );
      return (
        isStandardFamily(base.split(/[-,]/)[0] ?? "") &&
        firstNonWinAnsi(text) === undefined
      );
    }
    const coverage = fontCoverage(pdfium, font);
    if (!coverage) return false;
    for (const character of text) {
      const code = character.codePointAt(0)!;
      if (code !== 0x0a && code !== 0x20 && !coverage.has(code)) return false;
    }
    return true;
  });
}

function fontCoverage(pdfium: Pdfium, font: number): CmapCoverage | undefined {
  const { lib } = pdfium;
  const size = pdfium.readNumbers(1, "i32", ([out]) =>
    lib.FPDFFont_GetFontData(font, 0, 0, out!),
  )?.[0];
  if (!size) return undefined;
  const buffer = pdfium.malloc(size);
  try {
    const out = pdfium.malloc(4);
    try {
      if (!lib.FPDFFont_GetFontData(font, buffer, size, out)) return undefined;
    } finally {
      pdfium.free(out);
    }
    const data = pdfium.readBytes(buffer, size);
    const tag = String.fromCharCode(data[0]!, data[1]!, data[2]!, data[3]!);
    const truetype =
      (data[0] === 0 && data[1] === 1 && data[2] === 0 && data[3] === 0) ||
      tag === "true";
    return truetype ? parseCmap(data) : undefined;
  } finally {
    pdfium.free(buffer);
  }
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
  const readBack = context.withPage(target.location.pageIndex, (page) => {
    const textPage = lib.FPDFText_LoadPage(page);
    try {
      return pdfium.readWideString((buffer, bytes) =>
        lib.FPDFTextObj_GetText(
          lib.FPDFPage_GetObject(page, index),
          textPage,
          buffer,
          bytes,
        ),
      );
    } finally {
      lib.FPDFText_ClosePage(textPage);
    }
  });
  if (readBack === text) return true;
  // Put the old text back so the fallback path starts from a known state.
  context.withPage(target.location.pageIndex, (page) => {
    setText(pdfium, lib.FPDFPage_GetObject(page, index), previous);
  });
  return false;
}

/** Replaces the object with one in a covering font at the same place, size and colour. */
function replaceWithFallback(
  context: OperationContext,
  target: TextTarget,
  text: string,
): OperationResult {
  const { pdfium } = context;
  const { lib } = pdfium;
  const { location, element } = target;
  const style = element.textStyle!;
  const index = location.indexes[0]!;
  const font = context.fonts.resolve(pdfium, context.document, {
    family: style.fontFamily,
    bold: style.bold,
    italic: style.italic,
    text,
  });
  context.withPage(location.pageIndex, (page) => {
    const old = lib.FPDFPage_GetObject(page, index);
    const matrix = pdfium.readNumbers(6, "float", ([pointer]) =>
      lib.FPDFPageObj_GetMatrix(old, pointer!),
    ) ?? [1, 0, 0, 1, 0, 0];
    const size =
      pdfium.readNumbers(1, "float", ([pointer]) =>
        lib.FPDFTextObj_GetFontSize(old, pointer!),
      )?.[0] ?? style.fontSize;
    const object = lib.FPDFPageObj_CreateTextObj(
      context.document,
      font.handle,
      size,
    );
    setText(pdfium, object, text);
    const [r, g, b] = parseColor(style.color);
    lib.FPDFPageObj_SetFillColor(object, r, g, b, 255);
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
    warnings: [
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
    const object = lib.FPDFPageObj_CreateTextObj(
      context.document,
      lib.FPDFTextObj_GetFont(old),
      fontSize,
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
