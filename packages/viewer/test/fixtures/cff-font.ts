/*
 * Minimal CFF font programs for tests, written byte by byte: named glyphs
 * with advance widths and, where asked, a box for an outline, the shape a
 * PDF's FontFile3 /Type1C subset has. Every number is written as a five-byte integer, so the top
 * DICT has the same size whatever offsets it holds.
 */

/** String ids of the standard glyph names these fixtures use. */
const STANDARD_SIDS: Readonly<Record<string, number>> = {
  space: 1,
  zero: 17,
  A: 34,
  B: 35,
  C: 36,
  a: 66,
  b: 67,
};

export interface CffGlyph {
  readonly name: string;
  /** Without one, the glyph takes the font's default width. */
  readonly width?: number;
  /** The width is pushed before a call to a global subroutine that ends the glyph. */
  readonly viaSubroutine?: boolean;
  /** The height of a box drawn on the baseline, inset 50 units from each side; none without one. */
  readonly ink?: number;
}

export function cffFont({
  name = "ABCDEF+TestFace",
  glyphs,
  defaultWidth = 500,
  nominalWidth = 400,
  cid = false,
  bbox = [0, -200, 1000, 800],
}: {
  name?: string;
  glyphs: readonly CffGlyph[];
  defaultWidth?: number;
  nominalWidth?: number;
  cid?: boolean;
  bbox?: readonly [number, number, number, number];
}): Uint8Array {
  const strings: string[] = [];
  const sid = (glyphName: string): number => {
    const standard = STANDARD_SIDS[glyphName];
    if (standard !== undefined) return standard;
    if (!strings.includes(glyphName)) strings.push(glyphName);
    return 391 + strings.indexOf(glyphName);
  };
  // A CID-keyed font names its registry and ordering as strings.
  const ros = cid ? [sid("Adobe"), sid("Identity")] : [];
  // A CID-keyed font's charset holds CIDs, here 1, 2, 3 and so on.
  const charset = [
    0,
    ...glyphs.flatMap((glyph, index) => u16(cid ? index + 1 : sid(glyph.name))),
  ];
  const charStrings = [
    [14], // .notdef: endchar
    ...glyphs.map((glyph) => [
      ...(glyph.width === undefined
        ? []
        : number(glyph.width - nominalWidth, "charstring")),
      ...(glyph.ink === undefined
        ? []
        : box(glyph.width ?? defaultWidth, glyph.ink)),
      // Subroutine 0 is numbered -107 with the bias of a small set: callgsubr.
      ...(glyph.viaSubroutine ? [32, 29] : [14]),
    ]),
  ];
  const privateDict = [
    ...number(defaultWidth),
    20,
    ...number(nominalWidth),
    21,
  ];
  // Every glyph in the one font DICT, format 0.
  const fdSelect = cid ? [0, ...charStrings.map(() => 0)] : [];
  const fontDict = (privateAt: number) => [
    ...number(privateDict.length),
    ...number(privateAt),
    18,
  ];
  const fdArraySize = cid ? index([fontDict(0)]).length : 0;
  const nameIndex = index([ascii(name)]);
  const stringIndex = index(strings.map(ascii));
  // One global subroutine, endchar, for glyphs that call it.
  const globalSubrs = index(
    glyphs.some((glyph) => glyph.viaSubroutine) ? [[14]] : [],
  );
  const charStringIndex = index(charStrings);
  const topDict = (layout: Layout) => [
    ...(cid
      ? [
          ...ros.flatMap((id) => number(id)),
          ...number(0),
          12,
          30,
          ...number(charStrings.length),
          12,
          34,
          ...number(layout.fdArrayAt),
          12,
          36,
          ...number(layout.fdSelectAt),
          12,
          37,
        ]
      : [...number(privateDict.length), ...number(layout.privateAt), 18]),
    ...bbox.flatMap((value) => number(value)),
    5,
    ...number(layout.charsetAt),
    15,
    ...number(layout.charStringsAt),
    17,
  ];
  // Sizes first, with placeholder offsets of the same width, then the real ones.
  const topSize = index([
    topDict({
      charsetAt: 0,
      fdSelectAt: 0,
      charStringsAt: 0,
      fdArrayAt: 0,
      privateAt: 0,
    }),
  ]).length;
  const charsetAt =
    4 + nameIndex.length + topSize + stringIndex.length + globalSubrs.length;
  const fdSelectAt = charsetAt + charset.length;
  const charStringsAt = fdSelectAt + fdSelect.length;
  const fdArrayAt = charStringsAt + charStringIndex.length;
  const privateAt = fdArrayAt + fdArraySize;
  const layout = { charsetAt, fdSelectAt, charStringsAt, fdArrayAt, privateAt };
  return Uint8Array.from([
    1,
    0,
    4,
    4,
    ...nameIndex,
    ...index([topDict(layout)]),
    ...stringIndex,
    ...globalSubrs,
    ...charset,
    ...fdSelect,
    ...charStringIndex,
    ...(cid ? index([fontDict(privateAt)]) : []),
    ...privateDict,
  ]);
}

interface Layout {
  readonly charsetAt: number;
  readonly fdSelectAt: number;
  readonly charStringsAt: number;
  readonly fdArrayAt: number;
  readonly privateAt: number;
}

/** A closed box from x 50 to `width - 50` and y 0 to `height`: rmoveto, then rlineto. */
function box(width: number, height: number): number[] {
  const side = width - 100;
  return [
    ...[50, 0].flatMap((value) => number(value, "charstring")),
    21,
    ...[side, 0, 0, height, -side, 0].flatMap((value) =>
      number(value, "charstring"),
    ),
    5,
  ];
}

/** A DICT operand as a five-byte integer, or a charstring operand as its own 255 form (16.16). */
function number(value: number, kind: "dict" | "charstring" = "dict"): number[] {
  return kind === "dict"
    ? [29, ...i32(value)]
    : [255, ...i32(Math.round(value * 65536))];
}

function index(items: readonly (readonly number[])[]): number[] {
  if (items.length === 0) return [0, 0];
  const offsets = [1];
  for (const item of items) offsets.push(offsets.at(-1)! + item.length);
  return [
    ...u16(items.length),
    4,
    ...offsets.flatMap((offset) => i32(offset)),
    ...items.flat(),
  ];
}

function ascii(text: string): number[] {
  return [...text].map((character) => character.charCodeAt(0));
}

function u16(value: number): number[] {
  return [(value >> 8) & 0xff, value & 0xff];
}

function i32(value: number): number[] {
  return [
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff,
  ];
}
