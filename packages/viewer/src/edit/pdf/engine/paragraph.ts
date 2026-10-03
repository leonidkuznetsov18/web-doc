import type { PageRect } from "../../types.js";
import type { PdfElement, PdfTextParagraph } from "../types.js";
import {
  OBJECT_TEXT,
  paragraphSpecOf,
  type MarkParams,
  type ObjectRecord,
} from "./elements.js";
import {
  roundRect,
  round,
  unionRects,
  userToPage,
  type PageGeometry,
} from "./geometry.js";
import { fontBytes, fontCanDraw, type TextMeasurer } from "./fonts.js";
import type { Pdfium } from "./pdfium.js";

/** Persisted inputs, validated at the PDF marked-content boundary. */
export interface ParagraphSpec extends MarkParams {
  readonly kind: "paragraph";
  readonly rect: PageRect;
  readonly text: string;
  /** Actual native line strings, retained to detect edits made by another tool. */
  readonly lines: readonly string[];
  readonly style: PdfTextParagraph["textStyle"];
  /** First baseline relative to rect.y, in multiples of the current style font size. */
  readonly baselineOffset: number;
}

export interface ParagraphTarget {
  readonly paragraph: PdfTextParagraph;
  readonly spec: ParagraphSpec;
  /** Native drawing-order indexes, not logical reading order. */
  readonly indexes: readonly number[];
}

export function paragraphElement(paragraph: PdfTextParagraph): PdfElement {
  return {
    id: paragraph.id,
    kind: "paragraph",
    pageIndex: paragraph.pageIndex,
    text: paragraph.text,
    bounds: paragraph.bounds,
    textStyle: paragraph.textStyle,
    operations: [
      "replaceParagraphText",
      "replaceText",
      "setTextStyle",
      "deleteElement",
    ],
  };
}

interface Row {
  readonly element: PdfElement;
  readonly index: number;
  readonly font: string;
  readonly x: number;
  readonly baseline: number;
  readonly width: number;
}

const MIN_CONTINUATION_FILL = 0.9;
const MAX_CONNECTED_EDGE_GAP = 0.65;
// Public PDF geometry is canonicalized to thousandths of a point. This only
// absorbs native floating-point rounding, not visible tracking or kerning.
const GLYPH_POSITION_PRECISION = 0.001;

interface NativeGlyph {
  readonly text: string;
  readonly x: number;
  readonly y: number;
}

/**
 * Conservative import recognition. Only whole, equally styled native lines
 * with uniform upright matrices and regular left-aligned leading participate.
 * Content-stream order is deliberately irrelevant to reading order.
 */
export function discoverParagraphs(
  pdfium: Pdfium,
  page: number,
  pageIndex: number,
  geometry: PageGeometry,
  records: readonly ObjectRecord[],
  elements: readonly PdfElement[],
  textPage: number,
  measurer: TextMeasurer,
): ReadonlyMap<string, ParagraphTarget> {
  const result = new Map<string, ParagraphTarget>();
  if (geometry.rotation !== 0) return result;
  for (const element of elements) {
    const indexes = records.flatMap((record, index) =>
      record.id === element.id ? [index] : [],
    );
    const spec = paragraphSpecOf(records[indexes[0] ?? -1]?.mark);
    if (!spec) continue;
    const paragraph: PdfTextParagraph = {
      id: element.id,
      pageIndex,
      text: spec.text,
      bounds: spec.rect,
      textStyle: spec.style,
      memberIds: [element.id],
      members: [{ elementId: element.id, start: 0, end: spec.text.length }],
    };
    result.set(element.id, { paragraph, spec, indexes });
  }
  // Arbitrarily rotated pages/objects require transformed layout and are left
  // independently editable until that native layout contract exists.
  const rows: Row[] = [];
  const fontKeys = new Map<number, string>();
  const programs: Uint8Array[] = [];
  const fontKey = (font: number): string => {
    const cached = fontKeys.get(font);
    if (cached) return cached;
    let key: string;
    if (!pdfium.lib.FPDFFont_GetIsEmbedded(font))
      key = `standard:${pdfium.readUtf8String((buffer, bytes) => pdfium.lib.FPDFFont_GetBaseFontName(font, buffer, bytes))}`;
    else {
      const bytes = fontBytes(pdfium, font);
      let at = bytes
        ? programs.findIndex(
            (program) =>
              program.length === bytes.length &&
              program.every((byte, index) => byte === bytes[index]),
          )
        : -1;
      if (at < 0 && bytes) {
        at = programs.length;
        programs.push(bytes);
      }
      key = bytes ? `embedded:${at}` : `unreadable:${font}`;
    }
    fontKeys.set(font, key);
    return key;
  };
  const { lib } = pdfium;
  const rightEdges = new Map<number, number>();
  const glyphsByObject = new Map<number, NativeGlyph[]>();
  for (let at = 0; at < lib.FPDFText_CountChars(textPage); at += 1) {
    const unicode = lib.FPDFText_GetUnicode(textPage, at);
    if (
      !Number.isInteger(unicode) ||
      unicode < 32 ||
      unicode > 0x10ffff ||
      (unicode >= 0xd800 && unicode <= 0xdfff)
    )
      continue;
    const object = lib.FPDFText_GetTextObject(textPage, at);
    if (!object) continue;
    const text = String.fromCodePoint(unicode);
    const origin = pdfium.readNumbers(2, "double", ([x, y]) =>
      lib.FPDFText_GetCharOrigin(textPage, at, x!, y!),
    );
    if (origin) {
      const point = userToPage(geometry, origin[0]!, origin[1]!);
      const glyphs = glyphsByObject.get(object) ?? [];
      glyphs.push({ text, ...point });
      glyphsByObject.set(object, glyphs);
    }
    if (/\s/u.test(text)) continue;
    const box = pdfium.readNumbers(4, "float", ([out]) =>
      lib.FPDFText_GetLooseCharBox(textPage, at, out!),
    );
    if (box)
      rightEdges.set(
        object,
        Math.max(
          rightEdges.get(object) ?? -Infinity,
          userToPage(geometry, box[2]!, box[1]!).x,
        ),
      );
  }
  records.forEach((record, index) => {
    if (
      record.mark ||
      record.staleMarkId ||
      record.type !== OBJECT_TEXT ||
      record.id.includes("/")
    )
      return;
    const element = elements.find((item) => item.id === record.id);
    if (
      !element?.textStyle ||
      !element.text?.trim() ||
      element.text !== element.text.trimStart() ||
      element.rotation ||
      element.text.includes("\u0002") ||
      /[\r\n\t]/u.test(element.text)
    )
      return;
    const object = lib.FPDFPage_GetObject(page, index);
    const matrix = pdfium.readNumbers(6, "float", ([out]) =>
      lib.FPDFPageObj_GetMatrix(object, out!),
    );
    if (
      !matrix ||
      matrix[0]! <= 0 ||
      Math.abs(matrix[1]!) > 0.001 ||
      Math.abs(matrix[2]!) > 0.001 ||
      Math.abs(matrix[0]! - matrix[3]!) > 0.001
    )
      return;
    const rgba = pdfium.readNumbers(4, "i32", ([r, g, b, a]) =>
      lib.FPDFPageObj_GetFillColor(object, r!, g!, b!, a!),
    );
    if (rgba?.[3] !== 255 || lib.FPDFTextObj_GetTextRenderMode(object) !== 0)
      return;
    const font = lib.FPDFTextObj_GetFont(object);
    const origin = userToPage(geometry, matrix[4]!, matrix[5]!);
    const rawSize = pdfium.readNumbers(1, "float", ([out]) =>
      lib.FPDFTextObj_GetFontSize(object, out!),
    )?.[0];
    if (!rawSize) return;
    const effective = {
      ...element,
      textStyle: {
        ...element.textStyle,
        fontSize: round(rawSize * matrix[0]!),
      },
    };
    if (effective.textStyle.fontSize < 1 || effective.textStyle.fontSize > 500)
      return;
    // Reused fonts must draw the same advances as our native reflow. CFF and
    // other non-writable fonts already require an explicit font-substitution
    // warning on replacement; their metrics are not claimed preserved here.
    if (
      fontCanDraw(pdfium, font, element.text) &&
      !hasDefaultAdvances(
        glyphsByObject.get(object) ?? [],
        element.text,
        font,
        rawSize * matrix[0]!,
        measurer,
      )
    )
      return;
    const width =
      (rightEdges.get(object) ?? element.bounds.x + element.bounds.width) -
      origin.x;
    if (!(width > 0) || element.bounds.width > width + 2) return;
    rows.push({
      element: effective,
      index,
      font: fontKey(font),
      x: origin.x,
      baseline: origin.y,
      width,
    });
  });
  rows.sort((a, b) => a.baseline - b.baseline || a.x - b.x);
  const used = new Set<string>();
  for (const first of rows) {
    if (used.has(first.element.id)) continue;
    const aligned = rows.filter(
      (row) =>
        !used.has(row.element.id) &&
        row.baseline >= first.baseline &&
        compatible(first, row),
    );
    const gaps = aligned
      .slice(1)
      .map((row, index) => row.baseline - aligned[index]!.baseline)
      .filter(
        (gap) =>
          gap >= first.element.textStyle!.fontSize * 0.8 &&
          gap <= first.element.textStyle!.fontSize * 2.1,
      );
    if (gaps.length === 0) continue;
    const leading = Math.min(...gaps);
    const group: Row[] = [];
    for (const row of aligned) {
      const previous = group.at(-1);
      if (
        previous &&
        (Math.abs(row.baseline - previous.baseline - leading) >
          Math.max(1, leading * 0.15) ||
          /[-\u00ad\u0002]$/u.test(previous.element.text!.trim()))
      )
        break;
      group.push(row);
    }
    if (group.length < 2) continue;
    const bounds = unionRects(group.map((row) => row.element.bounds));
    const rect = roundRect({
      x: first.x,
      y: bounds.y,
      width: Math.max(...group.map((row) => row.width)),
      height: bounds.height,
    });
    const ids = new Set(group.map((row) => row.element.id));
    const firstIndex = Math.min(...group.map((row) => row.index));
    // A short non-final line may end an independent paragraph. A same-baseline
    // neighbor may be another table cell or a separate list label. Neither has
    // enough evidence of ordinary paragraph wrapping to merge safely.
    if (
      group
        .slice(0, -1)
        .some((row) => row.width < rect.width * MIN_CONTINUATION_FILL) ||
      elements.some((element) => {
        if (ids.has(element.id) || !element.text?.trim()) return false;
        const b = element.bounds;
        const sharesRow = group.some((row) => {
          const a = row.element.bounds;
          return a.y < b.y + b.height && a.y + a.height > b.y;
        });
        if (sharesRow) return true;
        // Do not expose the matching tail/head of an unsupported connected
        // section (indentation, mixed style, excluded source characters).
        const sharesColumn =
          b.x < rect.x + rect.width && b.x + b.width > rect.x;
        const edgeGap = Math.max(
          b.y - (rect.y + rect.height),
          rect.y - (b.y + b.height),
        );
        return sharesColumn && edgeGap <= leading * MAX_CONNECTED_EDGE_GAP;
      })
    )
      continue;
    // Ambiguous interleaving/overlap may be another run, heading or figure.
    // A background completely behind the group is safe to keep untouched.
    if (
      elements.some(
        (element) =>
          !ids.has(element.id) &&
          intersects(rect, element.bounds) &&
          !backgroundBehind(element, rect, records, firstIndex),
      )
    )
      continue;
    const id = `${first.element.id}:paragraph`;
    if (elements.some((element) => element.id === id)) continue;
    let text = "";
    const members = group.map((row) => {
      if (text) text += " ";
      const start = text.length;
      text += row.element.text!.trim();
      return { elementId: row.element.id, start, end: text.length };
    });
    if (text.length > 20000) continue;
    const style = {
      ...first.element.textStyle!,
      lineHeight: leading / first.element.textStyle!.fontSize,
    };
    const paragraph: PdfTextParagraph = {
      id,
      pageIndex,
      text,
      bounds: rect,
      textStyle: style,
      memberIds: [...ids],
      members,
    };
    const target: ParagraphTarget = {
      paragraph,
      spec: {
        kind: "paragraph",
        id,
        text,
        lines: group.map((row) => row.element.text!.trim()),
        rect,
        style,
        baselineOffset: (first.baseline - rect.y) / style.fontSize,
      },
      indexes: group.map((row) => row.index).sort((a, b) => a - b),
    };
    result.set(id, target);
    for (const member of members) {
      result.set(member.elementId, target);
      used.add(member.elementId);
    }
  }
  return result;
}

function compatible(a: Row, b: Row): boolean {
  const x = a.element.textStyle!;
  const y = b.element.textStyle!;
  return (
    Math.abs(a.x - b.x) <= 1 &&
    a.font === b.font &&
    x.fontSize === y.fontSize &&
    x.color === y.color &&
    x.bold === y.bold &&
    x.italic === y.italic
  );
}

function hasDefaultAdvances(
  glyphs: readonly NativeGlyph[],
  text: string,
  font: number,
  fontSize: number,
  measurer: TextMeasurer,
): boolean {
  const characters = [...text.trimEnd()];
  const first = glyphs[0];
  if (!first || glyphs.length < characters.length) return false;
  let advance = 0;
  for (const [index, character] of characters.entries()) {
    const glyph = glyphs[index];
    if (
      !glyph ||
      glyph.text !== character ||
      Math.abs(glyph.x - first.x - advance) > GLYPH_POSITION_PRECISION ||
      Math.abs(glyph.y - first.y) > GLYPH_POSITION_PRECISION
    )
      return false;
    const width = measurer.advance(font, fontSize, character);
    if (!Number.isFinite(width) || width <= 0) return false;
    advance += width;
  }
  return true;
}

export function intersects(a: PageRect, b: PageRect): boolean {
  return (
    a.x < b.x + b.width &&
    a.x + a.width > b.x &&
    a.y < b.y + b.height &&
    a.y + a.height > b.y
  );
}

function backgroundBehind(
  element: PdfElement,
  rect: PageRect,
  records: readonly ObjectRecord[],
  first: number,
): boolean {
  const b = element.bounds;
  return (
    element.kind !== "text" &&
    element.text === undefined &&
    records.findIndex((record) => record.id === element.id) < first &&
    b.x <= rect.x &&
    b.y <= rect.y &&
    b.x + b.width >= rect.x + rect.width &&
    b.y + b.height >= rect.y + rect.height
  );
}
