import { GLYPH_UNICODE } from "./glyph-names.js";

/*
 * Reading a bare CFF font program (a PDF's FontFile3 /Type1C), far enough to
 * give a browser an OpenType font of it: its glyph names, advance widths, box
 * and units per em. A browser loads no bare CFF, and PDF.js loads PDF fonts
 * with their glyphs moved to the Private Use Area, so neither gives a face
 * typed text can use. Only what a well-formed font needs is read; anything
 * unexpected yields `undefined`, never a guess.
 *
 * The glyph names of the ISOAdobe charset are the standard strings of the
 * CFF specification (Adobe TN 5176, appendix C); see glyph-names.ts for the
 * code points names stand for.
 */

/** What a browser font is built from. */
export interface CffFont {
  /** The font's PostScript name, a subset tag included. */
  readonly name: string;
  /** CID-keyed: glyphs are numbered, not named, and only the PDF's CMap says what they draw. */
  readonly cid: boolean;
  /** Glyph names by glyph id; `undefined` where the name is not one this reader knows. */
  readonly glyphNames: readonly (string | undefined)[];
  /** Advance widths by glyph id, in font units. */
  readonly widths: readonly number[];
  readonly bbox: readonly [number, number, number, number];
  readonly unitsPerEm: number;
}

const ISO_ADOBE_NAMES = `
.notdef space exclam quotedbl numbersign dollar percent ampersand
quoteright parenleft parenright asterisk plus comma hyphen period slash
zero one two three four five six seven eight nine colon semicolon less
equal greater question at A B C D E F G H I J K L M N O P Q R S T U V W X Y
Z bracketleft backslash bracketright asciicircum underscore quoteleft a b c
d e f g h i j k l m n o p q r s t u v w x y z braceleft bar braceright
asciitilde exclamdown cent sterling fraction yen florin section currency
quotesingle quotedblleft guillemotleft guilsinglleft guilsinglright fi fl
endash dagger daggerdbl periodcentered paragraph bullet quotesinglbase
quotedblbase quotedblright guillemotright ellipsis perthousand questiondown
grave acute circumflex tilde macron breve dotaccent dieresis ring cedilla
hungarumlaut ogonek caron emdash AE ordfeminine Lslash Oslash OE
ordmasculine ae dotlessi lslash oslash oe germandbls onesuperior logicalnot
mu trademark Eth onehalf plusminus Thorn onequarter divide brokenbar degree
thorn threequarters twosuperior registered minus eth multiply threesuperior
copyright Aacute Acircumflex Adieresis Agrave Aring Atilde Ccedilla Eacute
Ecircumflex Edieresis Egrave Iacute Icircumflex Idieresis Igrave Ntilde
Oacute Ocircumflex Odieresis Ograve Otilde Scaron Uacute Ucircumflex
Udieresis Ugrave Yacute Ydieresis Zcaron aacute acircumflex adieresis
agrave aring atilde ccedilla eacute ecircumflex edieresis egrave iacute
icircumflex idieresis igrave ntilde oacute ocircumflex odieresis ograve
otilde scaron uacute ucircumflex udieresis ugrave yacute ydieresis zcaron
`
  .trim()
  .split(/\s+/);

/**
 * The code point a glyph name stands for: an Adobe Glyph List name, or
 * `uniXXXX` or `uXXXX[XX]` for one code point. Ligatures of several code
 * points, suffixed variants such as `a.sc`, and noncharacters stand for none.
 */
export function glyphUnicode(name: string | undefined): number | undefined {
  if (!name || name === ".notdef") return undefined;
  const known = GLYPH_UNICODE.get(name);
  if (known !== undefined) return known;
  const match = /^uni([0-9A-F]{4})$|^u([0-9A-F]{4,6})$/.exec(name);
  const hex = match?.[1] ?? match?.[2];
  if (!hex) return undefined;
  const code = Number.parseInt(hex, 16);
  const surrogate = code >= 0xd800 && code <= 0xdfff;
  const nonCharacter = (code & 0xfffe) === 0xfffe;
  return code <= 0x10ffff && !surrogate && !nonCharacter ? code : undefined;
}

/** Top and Private DICT operators this reader uses; two-byte ones as 1200 + their second byte. */
const OP = {
  fontBBox: 5,
  charset: 15,
  charStrings: 17,
  private: 18,
  defaultWidthX: 20,
  nominalWidthX: 21,
  subrs: 19,
  charstringType: 1206,
  fontMatrix: 1207,
  ros: 1230,
} as const;

/** Standard strings come first; a string id from here on names an entry of the String INDEX. */
const STANDARD_STRINGS = 391;

/** Reads a CFF font program, or `undefined` when it is not one this reader can use. */
export function readCff(bytes: Uint8Array): CffFont | undefined {
  try {
    const at = (offset: number): number => {
      if (!Number.isInteger(offset) || offset < 0 || offset >= bytes.length)
        throw new RangeError("Outside the font");
      return bytes[offset]!;
    };
    const u16 = (offset: number): number => (at(offset) << 8) | at(offset + 1);
    if (at(0) !== 1) return undefined;
    const names = readIndex(bytes, at(2), at, u16);
    const tops = readIndex(bytes, names.end, at, u16);
    const strings = readIndex(bytes, tops.end, at, u16);
    const globalSubrs = readIndex(bytes, strings.end, at, u16);
    // A PDF embeds one font per program.
    if (names.items.length !== 1 || tops.items.length !== 1) return undefined;
    const top = readDict(bytes, ...tops.items[0]!, at);
    const charStringsAt = top.get(OP.charStrings)?.[0];
    if (charStringsAt === undefined) return undefined;
    if ((top.get(OP.charstringType)?.[0] ?? 2) !== 2) return undefined;
    const charStrings = readIndex(bytes, charStringsAt, at, u16);
    const glyphCount = charStrings.items.length;
    if (glyphCount === 0) return undefined;
    const [scale = 0.001, skewX = 0, skewY = 0, scaleY = 0.001] =
      top.get(OP.fontMatrix) ?? [];
    if (skewX !== 0 || skewY !== 0 || scale !== scaleY || scale <= 0)
      return undefined;
    const unitsPerEm = Math.round(1 / scale);
    if (unitsPerEm < 16 || unitsPerEm > 16384) return undefined;
    const [xMin = 0, yMin = 0, xMax = 0, yMax = 0] = top.get(OP.fontBBox) ?? [];
    const cid = top.has(OP.ros);
    const [nameStart, nameEnd] = names.items[0]!;
    // A font name is 127 characters at most.
    const name = latin1(
      bytes.subarray(nameStart, Math.min(nameEnd, nameStart + 127)),
    );
    if (cid)
      return {
        name,
        cid,
        glyphNames: [],
        widths: [],
        bbox: [xMin, yMin, xMax, yMax],
        unitsPerEm,
      };
    // Glyphs may share a string; each is read once.
    const custom = new Map<number, string | undefined>();
    const stringOf = (sid: number): string | undefined => {
      if (sid < ISO_ADOBE_NAMES.length) return ISO_ADOBE_NAMES[sid];
      const item = strings.items[sid - STANDARD_STRINGS];
      if (!item) return undefined;
      if (!custom.has(sid))
        custom.set(
          sid,
          // A glyph name is 63 characters at most.
          item[1] - item[0] <= 63 ? latin1(bytes.subarray(...item)) : undefined,
        );
      return custom.get(sid);
    };
    const glyphNames = charsetNames(
      top.get(OP.charset)?.[0] ?? 0,
      glyphCount,
      stringOf,
      at,
      u16,
    );
    const [privateSize, privateAt] = top.get(OP.private) ?? [0, 0];
    const privateDict = privateSize
      ? readDict(bytes, privateAt!, privateAt! + privateSize!, at)
      : new Map<number, number[]>();
    const localAt = privateDict.get(OP.subrs)?.[0];
    const metrics: WidthContext = {
      bytes,
      defaultWidth: privateDict.get(OP.defaultWidthX)?.[0] ?? 0,
      nominalWidth: privateDict.get(OP.nominalWidthX)?.[0] ?? 0,
      global: globalSubrs.items,
      local:
        localAt === undefined
          ? []
          : readIndex(bytes, privateAt! + localAt, at, u16).items,
    };
    const widths = charStrings.items.map(([start, end]) =>
      charstringWidth(metrics, start, end),
    );
    return {
      name,
      cid,
      glyphNames,
      widths,
      bbox: [xMin, yMin, xMax, yMax],
      unitsPerEm,
    };
  } catch (error) {
    if (error instanceof RangeError) return undefined;
    throw error;
  }
}

/**
 * A copy of the program whose font name holds only what a browser's font
 * sanitizer accepts in a CFF table: printable ASCII without
 * `[](){}<>/%` or a space, each other byte becoming `_`; `undefined` past
 * the 127 characters it accepts or for a program that is not one font.
 */
export function cffWithSafeName(bytes: Uint8Array): Uint8Array | undefined {
  try {
    const at = (offset: number): number => {
      if (offset < 0 || offset >= bytes.length)
        throw new RangeError("Outside the font");
      return bytes[offset]!;
    };
    const u16 = (offset: number): number => (at(offset) << 8) | at(offset + 1);
    const names = readIndex(bytes, at(2), at, u16);
    const [start, end] = names.items[0] ?? [0, 0];
    if (names.items.length !== 1 || end - start > 127) return undefined;
    const safe = bytes.slice();
    for (let offset = start; offset < end; offset += 1) {
      const byte = safe[offset]!;
      // A name that starts with NUL marks a deleted font, which is allowed.
      if (offset === start && byte === 0) continue;
      if (byte < 33 || byte > 126 || UNSAFE_NAME_BYTES.has(byte))
        safe[offset] = 0x5f;
    }
    return safe;
  } catch (error) {
    if (error instanceof RangeError) return undefined;
    throw error;
  }
}

const UNSAFE_NAME_BYTES = new Set(
  [..."[](){}<>/%"].map((character) => character.charCodeAt(0)),
);

/** An INDEX: its items' byte ranges and where it ends. */
function readIndex(
  bytes: Uint8Array,
  start: number,
  at: (offset: number) => number,
  u16: (offset: number) => number,
): { readonly items: [number, number][]; readonly end: number } {
  const count = u16(start);
  if (count === 0) return { items: [], end: start + 2 };
  const offSize = at(start + 2);
  if (offSize < 1 || offSize > 4) throw new RangeError("Bad offset size");
  const offset = (index: number): number => {
    let value = 0;
    for (let byte = 0; byte < offSize; byte += 1)
      value = value * 256 + at(start + 3 + index * offSize + byte);
    return value;
  };
  if (offset(0) !== 1) throw new RangeError("Bad INDEX");
  // Offsets count from the byte before the data.
  const base = start + 2 + (count + 1) * offSize;
  const items: [number, number][] = [];
  for (let index = 0; index < count; index += 1) {
    const from = base + offset(index);
    const to = base + offset(index + 1);
    if (to < from || to > bytes.length) throw new RangeError("Bad INDEX");
    items.push([from, to]);
  }
  return { items, end: base + offset(count) };
}

/** A DICT's operands by operator. */
function readDict(
  bytes: Uint8Array,
  start: number,
  end: number,
  at: (offset: number) => number,
): Map<number, number[]> {
  if (end > bytes.length) throw new RangeError("DICT past the font");
  const entries = new Map<number, number[]>();
  let operands: number[] = [];
  let offset = start;
  while (offset < end) {
    const byte = at(offset);
    if (byte <= 21) {
      const operator = byte === 12 ? 1200 + at(offset + 1) : byte;
      offset += byte === 12 ? 2 : 1;
      entries.set(operator, operands);
      operands = [];
    } else if (byte === 28) {
      operands.push(int16((at(offset + 1) << 8) | at(offset + 2)));
      offset += 3;
    } else if (byte === 29) {
      operands.push(
        (at(offset + 1) << 24) |
          (at(offset + 2) << 16) |
          (at(offset + 3) << 8) |
          at(offset + 4),
      );
      offset += 5;
    } else if (byte === 30) {
      let text = "";
      offset += 1;
      for (let done = false; !done; offset += 1) {
        const pair = at(offset);
        for (const nibble of [pair >> 4, pair & 15]) {
          if (nibble === 15) {
            done = true;
            break;
          }
          text += REAL_NIBBLES[nibble] ?? "";
        }
      }
      operands.push(Number.parseFloat(text));
    } else if (byte >= 32 && byte <= 246) {
      operands.push(byte - 139);
      offset += 1;
    } else if (byte >= 247 && byte <= 250) {
      operands.push((byte - 247) * 256 + at(offset + 1) + 108);
      offset += 2;
    } else if (byte >= 251 && byte <= 254) {
      operands.push(-(byte - 251) * 256 - at(offset + 1) - 108);
      offset += 2;
    } else throw new RangeError("Bad DICT byte");
  }
  return entries;
}

const REAL_NIBBLES = [
  "0",
  "1",
  "2",
  "3",
  "4",
  "5",
  "6",
  "7",
  "8",
  "9",
  ".",
  "E",
  "E-",
  "",
  "-",
];

/** Glyph names by glyph id from the charset at `offset`; 0 is the ISOAdobe set. */
function charsetNames(
  offset: number,
  glyphCount: number,
  stringOf: (sid: number) => string | undefined,
  at: (offset: number) => number,
  u16: (offset: number) => number,
): (string | undefined)[] {
  const names: (string | undefined)[] = [".notdef"];
  if (offset === 0) {
    for (let glyph = 1; glyph < glyphCount; glyph += 1)
      names.push(stringOf(glyph));
    return names;
  }
  // The predefined expert sets name small capitals and old-style figures.
  if (offset <= 2) return [...names, ...Array<undefined>(glyphCount - 1)];
  const format = at(offset);
  let cursor = offset + 1;
  if (format === 0) {
    for (let glyph = 1; glyph < glyphCount; glyph += 1, cursor += 2)
      names.push(stringOf(u16(cursor)));
  } else if (format === 1 || format === 2) {
    while (names.length < glyphCount) {
      const first = u16(cursor);
      const left = format === 1 ? at(cursor + 2) : u16(cursor + 2);
      cursor += format === 1 ? 3 : 4;
      for (let step = 0; step <= left && names.length < glyphCount; step += 1)
        names.push(stringOf(first + step));
    }
  } else throw new RangeError("Bad charset format");
  return names;
}

interface WidthContext {
  readonly bytes: Uint8Array;
  readonly defaultWidth: number;
  readonly nominalWidth: number;
  readonly global: readonly [number, number][];
  readonly local: readonly [number, number][];
}

/**
 * A Type 2 charstring's advance width (Adobe TN 5177): an extra first operand
 * before the first operator that clears the stack, plus the nominal width;
 * the default width without one. Subroutines called on the way are followed,
 * ten deep at most. A charstring that breaks a rule gets the default width.
 */
function charstringWidth(
  context: WidthContext,
  start: number,
  end: number,
): number {
  const { bytes, defaultWidth, nominalWidth } = context;
  const stack: number[] = [];
  // A width, or `undefined` to go on after a subroutine returns.
  const walk = (
    from: number,
    to: number,
    depth: number,
  ): number | undefined => {
    let offset = from;
    while (offset < to) {
      const byte = bytes[offset]!;
      const size =
        byte === 28 ? 3 : byte === 255 ? 5 : byte >= 247 && byte <= 254 ? 2 : 1;
      if (offset + size > to) return defaultWidth;
      if (byte === 28 || byte >= 32) {
        const next = (step: number) => bytes[offset + step]!;
        stack.push(
          byte === 28
            ? int16((next(1) << 8) | next(2))
            : byte <= 246
              ? byte - 139
              : byte <= 250
                ? (byte - 247) * 256 + next(1) + 108
                : byte <= 254
                  ? -(byte - 251) * 256 - next(1) - 108
                  : ((next(1) << 24) |
                      (next(2) << 16) |
                      (next(3) << 8) |
                      next(4)) /
                    65536,
        );
        // The Type 2 stack holds 48 operands.
        if (stack.length > 48) return defaultWidth;
        offset += size;
        continue;
      }
      offset += 1;
      if (byte === 10 || byte === 29) {
        const subrs = byte === 10 ? context.local : context.global;
        const index = stack.pop();
        const subr =
          index === undefined
            ? undefined
            : subrs[index + subrBias(subrs.length)];
        if (!subr || depth >= 10) return defaultWidth;
        const width = walk(subr[0], subr[1], depth + 1);
        if (width !== undefined) return width;
        continue;
      }
      if (byte === 11) return depth > 0 ? undefined : defaultWidth;
      const takes = STACK_CLEARING[byte];
      if (takes === undefined) return defaultWidth;
      const count = stack.length;
      const widthFirst =
        takes === "pairs"
          ? count % 2 === 1
          : takes === "endchar"
            ? count === 1 || count === 5
            : count > takes;
      return widthFirst ? stack[0]! + nominalWidth : defaultWidth;
    }
    return depth > 0 ? undefined : defaultWidth;
  };
  return walk(start, end, 0) ?? defaultWidth;
}

/** What a subroutine number is offset by, for a set of `count` subroutines. */
function subrBias(count: number): number {
  return count < 1240 ? 107 : count < 33900 ? 1131 : 32768;
}

/** Operators that clear the stack, and the operands they take without a width. */
const STACK_CLEARING: Readonly<Record<number, number | "pairs" | "endchar">> = {
  1: "pairs", // hstem
  3: "pairs", // vstem
  18: "pairs", // hstemhm
  23: "pairs", // vstemhm
  19: "pairs", // hintmask, after implied vstems
  20: "pairs", // cntrmask
  21: 2, // rmoveto
  22: 1, // hmoveto
  4: 1, // vmoveto
  14: "endchar", // endchar: none, or four for an accented character
};

function int16(value: number): number {
  return value >= 0x8000 ? value - 0x10000 : value;
}

function latin1(bytes: Uint8Array): string {
  return String.fromCharCode(...bytes);
}
