import type {
  PagePoint,
  PageRect,
  TextPosition,
  TextRange,
} from "../../types.js";
import type { PdfElement, TextLayout, TextLayoutLine } from "../types.js";
import { textStyle } from "./elements.js";
import {
  pageToUser,
  rectContains,
  round,
  roundRect,
  unionRects,
  userRectToPage,
  userToPage,
  type PageGeometry,
} from "./geometry.js";
import type { Pdfium } from "./pdfium.js";

/*
 * The overlay primitives' view of a page: glyph geometry from PDFium's text
 * page, mapped to the elements that own the glyphs and to page space. A
 * layout line is one text object — a text box's lines and a table's cells
 * are separate objects already — so lines come out in drawing order, which
 * is the reading order web-doc writes them in.
 */

/** A loaded page with its text page and the character-to-element mapping. */
export interface TextPageScan {
  readonly page: number;
  readonly textPage: number;
  readonly geometry: PageGeometry;
  /** Object handle → element id, in drawing order. */
  readonly byObject: ReadonlyMap<number, string>;
  readonly elements: readonly PdfElement[];
  /** Per character of the text page: its element and offset, when an object draws it. */
  readonly offsets: readonly (TextPosition | undefined)[];
}

interface Glyph {
  readonly index: number;
  readonly object: number;
  readonly position: TextPosition;
  readonly char: string;
  readonly box: PageRect;
  readonly advance: number;
  readonly origin: PagePoint;
}

/** How far a point may be from a glyph for `FPDFText_GetCharIndexAtPos`. */
const HIT_TOLERANCE = 2;

const TEXT_KINDS = new Set(["text", "textBox", "table", "paragraph"]);

/** The layout of one element, or `undefined` when it draws no text. */
export function layoutOf(
  pdfium: Pdfium,
  scan: TextPageScan,
  element: PdfElement,
): TextLayout | undefined {
  if (!TEXT_KINDS.has(element.kind)) return undefined;
  const glyphs = glyphsOf(
    pdfium,
    scan,
    (position) => position.elementId === element.id,
  );
  return {
    elementId: element.id,
    pageIndex: element.pageIndex,
    lines: linesOf(pdfium, scan, glyphs),
  };
}

/** The layouts of every text element of the page, in reading order, from one pass. */
export function layoutsOf(pdfium: Pdfium, scan: TextPageScan): TextLayout[] {
  const glyphs = glyphsOf(pdfium, scan, () => true);
  const byElement = new Map<string, Glyph[]>();
  for (const glyph of glyphs) {
    const own = byElement.get(glyph.position.elementId);
    if (own) own.push(glyph);
    else byElement.set(glyph.position.elementId, [glyph]);
  }
  const layouts: TextLayout[] = [];
  for (const element of scan.elements) {
    const own = byElement.get(element.id);
    if (!own || !TEXT_KINDS.has(element.kind)) continue;
    layouts.push({
      elementId: element.id,
      pageIndex: element.pageIndex,
      lines: linesOf(pdfium, scan, own),
    });
  }
  return layouts;
}

/** The caret position nearest to a page-space point, or none without text. */
export function positionIn(
  pdfium: Pdfium,
  scan: TextPageScan,
  point: PagePoint,
): TextPosition | undefined {
  const glyphs = glyphsOf(pdfium, scan, () => true);
  if (glyphs.length === 0) return undefined;
  const user = pageToUser(scan.geometry, point.x, point.y);
  const hit = pdfium.lib.FPDFText_GetCharIndexAtPos(
    scan.textPage,
    user.x,
    user.y,
    HIT_TOLERANCE,
    HIT_TOLERANCE,
  );
  let glyph = glyphs.find((entry) => entry.index === hit);
  if (!glyph) {
    let best = Number.POSITIVE_INFINITY;
    for (const entry of glyphs) {
      const distance = distanceToRect(point, entry.box);
      if (distance < best) {
        best = distance;
        glyph = entry;
      }
    }
  }
  // Past the glyph's middle along the reading direction, the caret goes after it.
  const after = alongReading(point, glyph!.box, scan.geometry.rotation) > 0.5;
  return {
    elementId: glyph!.position.elementId,
    offset: glyph!.position.offset + (after ? 1 : 0),
  };
}

/** The rectangles a range covers: one per line fragment, in reading order. */
export function rectsOf(
  pdfium: Pdfium,
  scan: TextPageScan,
  range: TextRange,
): PageRect[] {
  const order = scan.elements.map((element) => element.id);
  const first = order.indexOf(range.start.elementId);
  const last = order.indexOf(range.end.elementId);
  if (first < 0 || last < 0 || first > last) return [];
  const glyphs = glyphsOf(pdfium, scan, (position) => {
    const at = order.indexOf(position.elementId);
    if (at < first || at > last) return false;
    const from = at === first ? range.start.offset : 0;
    const to = at === last ? range.end.offset : Number.POSITIVE_INFINITY;
    return position.offset >= from && position.offset < to;
  });
  return linesOf(pdfium, scan, glyphs).map((line) => line.bounds);
}

function glyphsOf(
  pdfium: Pdfium,
  scan: TextPageScan,
  keep: (position: TextPosition) => boolean,
): Glyph[] {
  const { lib } = pdfium;
  const { textPage, geometry, offsets } = scan;
  const glyphs: Glyph[] = [];
  for (let index = 0; index < offsets.length; index += 1) {
    const position = offsets[index];
    if (!position || !keep(position)) continue;
    const object = lib.FPDFText_GetTextObject(textPage, index);
    const tight = pdfium.readNumbers(4, "double", ([l, r, b, t]) =>
      lib.FPDFText_GetCharBox(textPage, index, l!, r!, b!, t!),
    );
    // FS_RECTF: left, top, right, bottom.
    const loose = pdfium.readNumbers(4, "float", ([pointer]) =>
      lib.FPDFText_GetLooseCharBox(textPage, index, pointer!),
    );
    const origin = pdfium.readNumbers(2, "double", ([x, y]) =>
      lib.FPDFText_GetCharOrigin(textPage, index, x!, y!),
    );
    if (!loose || !origin) continue;
    const [looseLeft, looseTop, looseRight, looseBottom] = loose as [
      number,
      number,
      number,
      number,
    ];
    const [left, right, bottom, top] =
      tight && tight[0] !== tight[1] && tight[2] !== tight[3]
        ? (tight as [number, number, number, number])
        : [looseLeft, looseRight, looseBottom, looseTop];
    glyphs.push({
      index,
      object,
      position,
      char: String.fromCodePoint(lib.FPDFText_GetUnicode(textPage, index)),
      box: roundRect(userRectToPage(geometry, left, bottom, right, top)),
      advance: round(Math.abs(looseRight - looseLeft)),
      origin: roundPoint(userToPage(geometry, origin[0]!, origin[1]!)),
    });
  }
  return glyphs;
}

/** Groups glyphs into lines, one per text object, in drawing order. */
function linesOf(
  pdfium: Pdfium,
  scan: TextPageScan,
  glyphs: readonly Glyph[],
): TextLayoutLine[] {
  const byObject = new Map<number, Glyph[]>();
  for (const glyph of glyphs) {
    const line = byObject.get(glyph.object);
    if (line) line.push(glyph);
    else byObject.set(glyph.object, [glyph]);
  }
  const lines: TextLayoutLine[] = [];
  for (const object of scan.byObject.keys()) {
    const members = byObject.get(object);
    if (!members) continue;
    members.sort((a, b) => a.position.offset - b.position.offset);
    const style = textStyle(pdfium, object);
    const { elementId } = members[0]!.position;
    lines.push({
      range: {
        start: { elementId, offset: members[0]!.position.offset },
        end: { elementId, offset: members.at(-1)!.position.offset + 1 },
      },
      text: members.map((glyph) => glyph.char).join(""),
      bounds: roundRect(unionRects(members.map((glyph) => glyph.box))),
      baseline: members[0]!.origin,
      glyphs: members.map(({ position, box, advance }) => ({
        offset: position.offset,
        box,
        advance,
      })),
      fontFamily: style.fontFamily,
      fontSize: style.fontSize,
      color: style.color,
    });
  }
  return lines;
}

/** Where a point falls along a glyph in reading direction, 0 before it to 1 after. */
function alongReading(
  point: PagePoint,
  box: PageRect,
  rotation: number,
): number {
  const dx = box.width > 0 ? (point.x - box.x) / box.width : 0.5;
  const dy = box.height > 0 ? (point.y - box.y) / box.height : 0.5;
  // Page rotation turns the reading direction: 0 → right, 1 → down, 2 → left, 3 → up.
  switch (rotation) {
    case 1:
      return dy;
    case 2:
      return 1 - dx;
    case 3:
      return 1 - dy;
    default:
      return dx;
  }
}

function distanceToRect(point: PagePoint, rect: PageRect): number {
  if (rectContains(rect, point)) return 0;
  const dx = Math.max(rect.x - point.x, 0, point.x - (rect.x + rect.width));
  const dy = Math.max(rect.y - point.y, 0, point.y - (rect.y + rect.height));
  return Math.hypot(dx, dy);
}

function roundPoint(point: PagePoint): PagePoint {
  return { x: round(point.x), y: round(point.y) };
}
