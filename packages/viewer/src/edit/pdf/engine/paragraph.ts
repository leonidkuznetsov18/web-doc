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
  readonly fontHandle: number;
  readonly x: number;
  readonly baseline: number;
  readonly width: number;
  /** Observed first word plus its following space, when another word follows. */
  readonly firstWordAdvance: number | undefined;
}

interface UprightText {
  readonly element: PdfElement;
  readonly baseline: number;
}

const MIN_CONTINUATION_FILL = 0.9;
const MAX_CONNECTED_EDGE_GAP = 0.65;
const MAX_LEFT_ALIGNMENT_DELTA = 1;
const ADJACENT_BASELINE_EM = 0.25;
const ADJACENT_TEXT_GAP_EM = 3;
const LIST_ITEM_START = /^(?:[•◦▪‣·*–—-]|\(?\d{1,3}[.)]|\(?[a-z][.)])\s/iu;
const NUMERIC_VALUE = /^[\d\s.,:/%()+−–—-]+$/u;
// Public PDF geometry is canonicalized to thousandths of a point. This only
// absorbs native floating-point rounding, not visible tracking or kerning.
const GLYPH_POSITION_PRECISION = 0.001;
const MAX_PAIR_ADJUSTMENT_EM = 0.1;
const MATRIX_RELATIVE_PRECISION = 0.000001;

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
  const indexesById = new Map<string, number[]>();
  records.forEach((record, index) => {
    const indexes = indexesById.get(record.id) ?? [];
    indexes.push(index);
    indexesById.set(record.id, indexes);
  });
  const elementsById = new Map<string, PdfElement>();
  for (const element of elements)
    if (!elementsById.has(element.id)) elementsById.set(element.id, element);
  for (const element of elements) {
    const indexes = indexesById.get(element.id) ?? [];
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
  const uprightTexts = new Map<string, UprightText>();
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
    const element = elementsById.get(record.id);
    if (!element?.text?.trim()) return;
    const object = lib.FPDFPage_GetObject(page, index);
    const matrix = pdfium.readNumbers(6, "float", ([out]) =>
      lib.FPDFPageObj_GetMatrix(object, out!),
    );
    if (
      !matrix ||
      matrix[0]! <= 0 ||
      matrix[3]! <= 0 ||
      Math.abs(matrix[1]!) > 0.001 ||
      Math.abs(matrix[2]!) > 0.001
    )
      return;
    const origin = userToPage(geometry, matrix[4]!, matrix[5]!);
    uprightTexts.set(element.id, { element, baseline: origin.y });
    if (
      !element.textStyle ||
      element.text !== element.text.trimStart() ||
      element.rotation ||
      element.text.includes("\u0002") ||
      /[\r\n\t]/u.test(element.text) ||
      LIST_ITEM_START.test(element.text) ||
      NUMERIC_VALUE.test(element.text) ||
      Math.abs(matrix[0]! - matrix[3]!) >
        Math.max(matrix[0]!, matrix[3]!) * MATRIX_RELATIVE_PRECISION
    )
      return;
    const rgba = pdfium.readNumbers(4, "i32", ([r, g, b, a]) =>
      lib.FPDFPageObj_GetFillColor(object, r!, g!, b!, a!),
    );
    if (rgba?.[3] !== 255 || lib.FPDFTextObj_GetTextRenderMode(object) !== 0)
      return;
    const font = lib.FPDFTextObj_GetFont(object);
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
    // Text/size reflow recomputes default native advances; bounded source pair
    // adjustments are not retained. Color-only changes retain source positions.
    // Non-writable fonts require the separate explicit substitution warning.
    if (
      fontCanDraw(pdfium, font, element.text) &&
      !hasSupportedAdvances(
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
      fontHandle: font,
      x: origin.x,
      baseline: origin.y,
      width,
      firstWordAdvance: observedFirstWordAdvance(
        glyphsByObject.get(object) ?? [],
        element.text,
      ),
    });
  });
  rows.sort((a, b) => a.baseline - b.baseline || a.x - b.x);
  // Neighboring cells/list labels share a baseline and a nearby horizontal
  // gap. Distant columns and rotated sidebars do not invalidate body text.
  // Keep blocked rows in `aligned` so leading inference is unchanged.
  const blockedRows = rowsWithAdjacentText(rows, [...uprightTexts.values()]);
  const used = new Set<string>();
  for (const first of rows) {
    if (used.has(first.element.id) || blockedRows.has(first.element.id))
      continue;
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
    if (
      group.length < 2 ||
      group.some((row) => blockedRows.has(row.element.id))
    )
      continue;
    const bounds = unionRects(group.map((row) => row.element.bounds));
    const rect = roundRect({
      x: first.x,
      y: bounds.y,
      width: Math.max(...group.map((row) => row.width)),
      height: bounds.height,
    });
    const ids = new Set(group.map((row) => row.element.id));
    const firstIndex = Math.min(...group.map((row) => row.index));
    // A short non-final line needs native evidence that the next word could
    // not fit; otherwise it may end an independent paragraph.
    if (
      group
        .slice(0, -1)
        .some(
          (row, index) =>
            !continuesAtWrap(
              row,
              group[index + 1]!,
              rect.width,
              pdfium,
              measurer,
            ),
        ) ||
      elements.some((element) => {
        if (ids.has(element.id) || !element.text?.trim()) return false;
        const b = element.bounds;
        const upright = uprightTexts.get(element.id);
        const sharesRow =
          upright && group.some((row) => isAdjacentText(row, upright));
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
    if (elementsById.has(id)) continue;
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
  return (
    Math.abs(a.x - b.x) <= MAX_LEFT_ALIGNMENT_DELTA && sameTypography(a, b)
  );
}

function sameTypography(a: Row, b: Row): boolean {
  const x = a.element.textStyle!;
  const y = b.element.textStyle!;
  return (
    a.font === b.font &&
    x.fontSize === y.fontSize &&
    x.color === y.color &&
    x.bold === y.bold &&
    x.italic === y.italic
  );
}

function rowsWithAdjacentText(
  rows: readonly Row[],
  texts: UprightText[],
): ReadonlySet<string> {
  texts.sort((a, b) => a.baseline - b.baseline);
  const blocked = new Set<string>();
  for (const row of rows) {
    const tolerance = row.element.textStyle!.fontSize * ADJACENT_BASELINE_EM;
    let low = 0;
    let high = texts.length;
    while (low < high) {
      const mid = (low + high) >>> 1;
      if (texts[mid]!.baseline < row.baseline - tolerance) low = mid + 1;
      else high = mid;
    }
    for (let at = low; at < texts.length; at += 1) {
      const other = texts[at]!;
      if (other.baseline > row.baseline + tolerance) break;
      if (isAdjacentText(row, other)) {
        blocked.add(row.element.id);
        break;
      }
    }
  }
  return blocked;
}

function isAdjacentText(row: Row, other: UprightText): boolean {
  if (row.element.id === other.element.id) return false;
  const size = row.element.textStyle!.fontSize;
  if (Math.abs(row.baseline - other.baseline) > size * ADJACENT_BASELINE_EM)
    return false;
  const a = row.element.bounds;
  const b = other.element.bounds;
  const gap = Math.max(b.x - (a.x + a.width), a.x - (b.x + b.width));
  return gap <= size * ADJACENT_TEXT_GAP_EM;
}

function continuesAtWrap(
  row: Row,
  next: Row,
  width: number,
  pdfium: Pdfium,
  measurer: TextMeasurer,
): boolean {
  if (row.width >= width * MIN_CONTINUATION_FILL) return true;
  // Source glyph origins also work for fonts whose encoding cannot be reused
  // (for example a subset missing U+0020). Do not measure its notdef space.
  if (next.firstWordAdvance !== undefined)
    return row.width + next.firstWordAdvance > width + GLYPH_POSITION_PRECISION;
  const nextWord = next.element.text!.trim().split(/\s/u)[0];
  if (!nextWord || !fontCanDraw(pdfium, row.fontHandle, ` ${nextWord}`))
    return false;
  const needed = measurer.advance(
    row.fontHandle,
    row.element.textStyle!.fontSize,
    ` ${nextWord}`,
  );
  return (
    Number.isFinite(needed) &&
    row.width + needed > width + GLYPH_POSITION_PRECISION
  );
}

function observedFirstWordAdvance(
  glyphs: readonly NativeGlyph[],
  text: string,
): number | undefined {
  const prefix = /^\S+\s+/u.exec(text)?.[0];
  if (!prefix) return undefined;
  const count = [...prefix].length;
  const first = glyphs[0];
  const nextWord = glyphs[count];
  if (
    !first ||
    !nextWord ||
    glyphs
      .slice(0, count)
      .map((glyph) => glyph.text)
      .join("") !== prefix
  )
    return undefined;
  const advance = nextWord.x - first.x;
  return Number.isFinite(advance) && advance > 0 ? advance : undefined;
}

/** A bounded eligibility rule, not a claim that source TJ positioning survives reflow. */
function hasSupportedAdvances(
  glyphs: readonly NativeGlyph[],
  text: string,
  font: number,
  fontSize: number,
  measurer: TextMeasurer,
): boolean {
  const characters = [...text.trimEnd()];
  const first = glyphs[0];
  if (!first || glyphs.length < characters.length) return false;
  const pairLimit = fontSize * MAX_PAIR_ADJUSTMENT_EM;
  let previousAdvance = 0;
  let cumulativeAdjustment = 0;
  let adjustedPairs = 0;
  let defaultPairs = 0;
  for (const [index, character] of characters.entries()) {
    const glyph = glyphs[index];
    if (
      !glyph ||
      glyph.text !== character ||
      !Number.isFinite(glyph.x) ||
      !Number.isFinite(glyph.y) ||
      Math.abs(glyph.y - first.y) > GLYPH_POSITION_PRECISION
    )
      return false;
    const previous = glyphs[index - 1];
    if (previous) {
      const adjustment = glyph.x - previous.x - previousAdvance;
      const changed = Math.abs(adjustment) > GLYPH_POSITION_PRECISION;
      // Real and inferred word spaces cannot be treated as pair adjustments.
      if (/\s/u.test(previous.text + character)) {
        if (changed) return false;
      } else if (changed) {
        if (Math.abs(adjustment) > pairLimit) return false;
        adjustedPairs += 1;
      } else {
        defaultPairs += 1;
      }
      cumulativeAdjustment += adjustment;
      if (Math.abs(cumulativeAdjustment) > pairLimit) return false;
    }
    previousAdvance = measurer.advance(font, fontSize, character);
    if (!Number.isFinite(previousAdvance) || previousAdvance <= 0) return false;
  }
  // Uniform tracking cannot pass merely because its per-glyph change is tiny.
  return adjustedPairs === 0 || defaultPairs > adjustedPairs;
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
