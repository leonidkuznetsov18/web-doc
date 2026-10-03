import { ViewerError } from "../../../errors.js";
import type { ViewerWarning } from "../../../contracts.js";
import type {
  ReplaceParagraphTextOperation,
  SetTextStyleOperation,
} from "../types.js";
import { MARK_NAME, MARK_PARAM, OBJECT_PATH, OBJECT_TEXT } from "./elements.js";
import { fontCanDraw, isStandardFamily, validateScript } from "./fonts.js";
import { pageToUser } from "./geometry.js";
import { layoutText, type Layout } from "./text-layout.js";
import { parseColor, placeUpright, setText } from "./text-box.js";
import {
  intersects,
  type ParagraphSpec,
  type ParagraphTarget,
} from "./paragraph.js";
import type {
  Issue,
  OperationContext,
  OperationHandler,
  OperationResult,
} from "./operations.js";

export const replaceParagraphText: OperationHandler<ReplaceParagraphTextOperation> =
  {
    validate(operation, context, issue) {
      const target = targetOf(operation.target, context, issue);
      if (!target) return;
      const text = replacement(operation, target.spec.text, issue);
      if (text === undefined || !validateScript(text, issue)) return;
      if (text.length > 20000) {
        issue(
          "/text",
          "max-length",
          "The complete paragraph cannot exceed 20000 characters",
        );
        return;
      }
      prepare(context, target, { ...target.spec, text }, issue);
    },
    apply(operation, context) {
      const target = requireTarget(operation.target, context);
      const text = replacement(operation, target.spec.text);
      if (text === undefined)
        throw new ViewerError("invalid-operation", "Invalid paragraph range");
      return rebuild(context, target, { ...target.spec, text });
    },
  };

export const paragraphSetTextStyle: OperationHandler<SetTextStyleOperation> = {
  validate(operation, context, issue) {
    const target = targetOf(operation.target, context, issue);
    if (!target) return;
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
          `Only color and fontSize can change on imported paragraphs; ${field} cannot`,
        );
    if (
      operation.style.fontSize !== undefined &&
      operation.style.fontSize !== target.spec.style.fontSize
    )
      prepare(context, target, styled(target.spec, operation), issue);
  },
  apply(operation, context) {
    const target = requireTarget(operation.target, context);
    if (
      operation.style.fontSize === undefined ||
      operation.style.fontSize === target.spec.style.fontSize
    )
      return colorInPlace(context, target, styled(target.spec, operation));
    return rebuild(context, target, styled(target.spec, operation));
  },
};

/** Color never rewrites glyphs or changes an imported font resource. */
function colorInPlace(
  context: OperationContext,
  target: ParagraphTarget,
  spec: ParagraphSpec,
): OperationResult {
  const { pdfium } = context;
  const { lib } = pdfium;
  const color = parseColor(spec.style.color);
  const persisted = target.paragraph.memberIds.includes(spec.id);
  context.withPage(target.paragraph.pageIndex, (page) => {
    for (const index of target.indexes) {
      const object = lib.FPDFPage_GetObject(page, index);
      if (lib.FPDFPageObj_GetType(object) === OBJECT_TEXT)
        lib.FPDFPageObj_SetFillColor(object, ...color, 255);
      if (!persisted) continue;
      for (let at = 0; at < lib.FPDFPageObj_CountMarks(object); at += 1) {
        const mark = lib.FPDFPageObj_GetMark(object, at);
        const name = pdfium.readWideStringOut((buffer, bytes, out) =>
          lib.FPDFPageObjMark_GetName(mark, buffer, bytes, out),
        );
        if (name === MARK_NAME)
          lib.FPDFPageObjMark_SetStringParam(
            context.document,
            object,
            mark,
            MARK_PARAM,
            JSON.stringify(spec),
          );
      }
      context.spliceObjects(target.paragraph.pageIndex, index, 1, [
        { id: spec.id, type: lib.FPDFPageObj_GetType(object), mark: spec },
      ]);
    }
  });
  return {
    createdIds: [],
    changedPages: [target.paragraph.pageIndex],
    warnings: [],
  };
}

function styled(
  spec: ParagraphSpec,
  operation: SetTextStyleOperation,
): ParagraphSpec {
  return {
    ...spec,
    style: {
      ...spec.style,
      ...(operation.style.fontSize === undefined
        ? {}
        : { fontSize: operation.style.fontSize }),
      ...(operation.style.color === undefined
        ? {}
        : { color: operation.style.color }),
    },
  };
}

function targetOf(
  id: string,
  context: OperationContext,
  issue: Issue,
): ParagraphTarget | undefined {
  const target = context.paragraph(id);
  if (!target || target.paragraph.id !== id) {
    issue(
      "/target",
      "unsupported-target",
      "The target must be a resolved native paragraph id",
    );
    return undefined;
  }
  return target;
}

function requireTarget(id: string, context: OperationContext): ParagraphTarget {
  const target = targetOf(id, context, () => {});
  if (!target)
    throw new ViewerError(
      "invalid-operation",
      "The paragraph no longer exists",
    );
  return target;
}

function replacement(
  operation: ReplaceParagraphTextOperation,
  previous: string,
  issue: Issue = () => {},
): string | undefined {
  const range = operation.range;
  if (!range) return operation.text;
  const valid = [range.start, range.end].every(
    (position) =>
      position.elementId === operation.target &&
      Number.isInteger(position.offset) &&
      position.offset >= 0 &&
      position.offset <= previous.length,
  );
  if (!valid || range.start.offset > range.end.offset) {
    issue(
      "/range",
      "invalid-range",
      "The range must lie inside the paragraph's logical text",
    );
    return undefined;
  }
  return (
    previous.slice(0, range.start.offset) +
    operation.text +
    previous.slice(range.end.offset)
  );
}

interface Prepared {
  readonly spec: ParagraphSpec;
  readonly layout: Layout;
  readonly warnings: readonly ViewerWarning[];
}

function fontFor(
  context: OperationContext,
  target: ParagraphTarget,
  spec: ParagraphSpec,
  page: number,
) {
  const { lib } = context.pdfium;
  const original = target.indexes
    .map((index) => lib.FPDFPage_GetObject(page, index))
    .find((object) => lib.FPDFPageObj_GetType(object) === OBJECT_TEXT);
  const native =
    original === undefined ? undefined : lib.FPDFTextObj_GetFont(original);
  if (native !== undefined && fontCanDraw(context.pdfium, native, spec.text))
    return {
      handle: native,
      family: spec.style.fontFamily,
      substituted: false,
    };
  const family =
    isStandardFamily(spec.style.fontFamily) ||
    context.fonts.hasFamily(spec.style.fontFamily)
      ? spec.style.fontFamily
      : "Helvetica";
  const request = {
    family,
    bold: spec.style.bold,
    italic: spec.style.italic,
    text: spec.text,
  };
  const problem = context.fonts.problem(request);
  if (problem) return { problem };
  const font = context.fonts.resolve(context.pdfium, context.document, request);
  return {
    handle: font.handle,
    family: font.family,
    substituted:
      original !== undefined || !isStandardFamily(spec.style.fontFamily),
    ...(font.substitution ? { substitution: font.substitution } : {}),
  };
}

/** Validate the actual final baseline/descent, not line-count × leading. */
function prepare(
  context: OperationContext,
  target: ParagraphTarget,
  spec: ParagraphSpec,
  issue: Issue,
): Prepared | undefined {
  return context.readPage(target.paragraph.pageIndex, (page) => {
    const font = fontFor(context, target, spec, page);
    if ("problem" in font) {
      issue(font.problem.path, font.problem.code, font.problem.message);
      return undefined;
    }
    const metrics = context.measurer.metrics(font.handle, spec.style.fontSize);
    const layout = layoutText({
      text: spec.text,
      width: spec.rect.width,
      height: Number.POSITIVE_INFINITY,
      fontSize: spec.style.fontSize,
      lineHeight: spec.style.lineHeight,
      align: "left",
      ascent: spec.baselineOffset * spec.style.fontSize,
      advance: (text) =>
        context.measurer.advance(font.handle, spec.style.fontSize, text),
    });
    const last = layout.lines.at(-1);
    const height =
      spec.text.trim() && last
        ? last.baseline + metrics.descent
        : spec.rect.height;
    const rect = { ...spec.rect, height: Math.max(height, 0.001) };
    const size = context.pageSize(target.paragraph.pageIndex);
    const members = new Set(target.paragraph.memberIds);
    const collision = context
      .pageElements(target.paragraph.pageIndex)
      .some((element) => {
        if (members.has(element.id) || !intersects(rect, element.bounds))
          return false;
        // An already-present background stays in its own drawing position.
        const old = target.spec.rect;
        const b = element.bounds;
        const location = context.locate(element.id);
        const behind =
          location &&
          location.indexes.every((index) => index < target.indexes[0]!);
        return !(
          element.text === undefined &&
          behind &&
          b.x <= old.x &&
          b.y <= old.y &&
          b.x + b.width >= rect.x + rect.width &&
          b.y + b.height >= rect.y + rect.height
        );
      });
    if (
      collision ||
      rect.x < 0 ||
      rect.y < 0 ||
      rect.x + rect.width > size.width + 0.01 ||
      rect.y + rect.height > size.height + 0.01
    ) {
      issue(
        "/text",
        "paragraph-overflow",
        "The paragraph would overlap neighboring content or leave the page",
      );
      return undefined;
    }
    const warnings: ViewerWarning[] = font.substituted
      ? [
          {
            code: "font-substitution",
            message:
              "substitution" in font && font.substitution
                ? font.substitution
                : `${spec.style.fontFamily} cannot draw the new paragraph; ${font.family} is used`,
            details: { elementId: spec.id },
          },
        ]
      : [];
    return {
      spec: {
        ...spec,
        rect,
        lines: layout.lines.map((line) => line.text).filter(Boolean),
        style: {
          ...spec.style,
          fontFamily: font.substituted ? font.family : spec.style.fontFamily,
        },
      },
      layout,
      warnings,
    };
  });
}

function rebuild(
  context: OperationContext,
  target: ParagraphTarget,
  spec: ParagraphSpec,
): OperationResult {
  const issues: { path: string; code: string; message: string }[] = [];
  const prepared = prepare(context, target, spec, (path, code, message) =>
    issues.push({ path, code, message }),
  );
  if (!prepared)
    throw new ViewerError(
      "invalid-operation",
      "The paragraph cannot be changed",
      { details: { issues } },
    );
  const { pdfium } = context;
  const { lib } = pdfium;
  const pageIndex = target.paragraph.pageIndex;
  const geometry = context.geometry(pageIndex);
  const first = target.indexes[0]!;
  context.withPage(pageIndex, (page) => {
    const font = fontFor(context, target, spec, page);
    if ("problem" in font)
      throw new ViewerError("invalid-operation", font.problem.message);
    const objects: { object: number; type: number }[] = [];
    const [r, g, b] = parseColor(spec.style.color);
    // Construct before removing borrowed font owners. Empty drafts use a
    // non-painting native path, so no placeholder character leaks into text.
    if (!spec.text.trim()) {
      const rect = prepared.spec.rect;
      const origin = pageToUser(geometry, rect.x, rect.y + rect.height);
      const object = lib.FPDFPageObj_CreateNewRect(
        origin.x,
        origin.y,
        rect.width,
        rect.height,
      );
      // PDFium discards a path ending in `n` when reopening. A zero-alpha
      // fill preserves the native marked object without painting any pixels.
      lib.FPDFPageObj_SetFillColor(object, 0, 0, 0, 0);
      lib.FPDFPath_SetDrawMode(object, 1, false);
      objects.push({ object, type: OBJECT_PATH });
    } else
      for (const line of prepared.layout.lines) {
        if (!line.text) continue;
        const object = lib.FPDFPageObj_CreateTextObj(
          context.document,
          font.handle,
          spec.style.fontSize,
        );
        if (!object) {
          for (const made of objects) lib.FPDFPageObj_Destroy(made.object);
          throw new ViewerError(
            "edit-failed",
            "PDFium could not create paragraph text",
          );
        }
        setText(pdfium, object, line.text);
        lib.FPDFPageObj_SetFillColor(object, r, g, b, 255);
        placeUpright(
          pdfium,
          object,
          geometry,
          spec.rect.x + line.x,
          spec.rect.y + line.baseline,
        );
        objects.push({ object, type: OBJECT_TEXT });
      }
    const params = JSON.stringify(prepared.spec);
    for (const { object } of objects) {
      const mark = lib.FPDFPageObj_AddMark(object, MARK_NAME);
      lib.FPDFPageObjMark_SetStringParam(
        context.document,
        object,
        mark,
        MARK_PARAM,
        params,
      );
    }
    for (const index of [...target.indexes].reverse()) {
      const object = lib.FPDFPage_GetObject(page, index);
      lib.FPDFPage_RemoveObject(page, object);
      lib.FPDFPageObj_Destroy(object);
      context.spliceObjects(pageIndex, index, 1, []);
    }
    objects.forEach(({ object }, at) =>
      lib.FPDFPage_InsertObjectAtIndex(page, object, first + at),
    );
    const textPage = lib.FPDFText_LoadPage(page);
    try {
      const drawn = objects
        .filter(({ type }) => type === OBJECT_TEXT)
        .map(({ object }) =>
          pdfium
            .readWideString((buffer, bytes) =>
              lib.FPDFTextObj_GetText(object, textPage, buffer, bytes),
            )
            .trimEnd(),
        );
      if (
        drawn.length !== prepared.spec.lines.length ||
        drawn.some((line, index) => line !== prepared.spec.lines[index])
      )
        throw new ViewerError(
          "edit-failed",
          "The native font could not encode the paragraph faithfully",
        );
    } finally {
      lib.FPDFText_ClosePage(textPage);
    }
    context.spliceObjects(
      pageIndex,
      first,
      0,
      objects.map(({ type }) => ({ id: spec.id, type, mark: prepared.spec })),
    );
  });
  const promoted = !target.paragraph.memberIds.includes(spec.id);
  // The logical paragraph already existed before promotion. Reporting it as
  // newly created would incorrectly destroy its ranges when Undo restores rows.
  return {
    createdIds: [],
    removedIds: promoted ? target.paragraph.memberIds : [],
    changedPages: [pageIndex],
    warnings: prepared.warnings,
  };
}
