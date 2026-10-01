import type { EditWorkerFont } from "../../../worker-protocol.js";
import type { Pdfium } from "./pdfium.js";

/*
 * Fonts for text web-doc writes into a PDF: the standard 14 fonts for
 * WinAnsi text, host-registered TrueType fonts, and a bundled Noto Sans
 * Latin/Cyrillic face as the last resort. Coverage is read from each font's
 * cmap, because PDFium draws a notdef glyph for anything else.
 */

const STANDARD_FAMILIES: Readonly<
  Record<string, readonly [string, string, string, string]>
> = {
  helvetica: [
    "Helvetica",
    "Helvetica-Bold",
    "Helvetica-Oblique",
    "Helvetica-BoldOblique",
  ],
  times: ["Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"],
  courier: [
    "Courier",
    "Courier-Bold",
    "Courier-Oblique",
    "Courier-BoldOblique",
  ],
};

/** Code points WinAnsiEncoding places in 0x80–0x9F. */
const WIN_ANSI_EXTRAS = new Set(
  "€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ".split("").map((c) => c.codePointAt(0)!),
);

/** FPDFText_LoadFont font types. */
const FONT_TRUETYPE = 1;

export const FALLBACK_FAMILY = "Noto Sans";

export function isStandardFamily(family: string): boolean {
  return family.toLowerCase() in STANDARD_FAMILIES;
}

/** The standard-14 name for a family and style, e.g. Helvetica + bold. */
export function standardFontName(
  family: string,
  bold: boolean,
  italic: boolean,
): string | undefined {
  const variants = STANDARD_FAMILIES[family.toLowerCase()];
  if (!variants) return undefined;
  return variants[(bold ? 1 : 0) + (italic ? 2 : 0)];
}

/** The first character a standard font cannot encode, if any. */
export function firstNonWinAnsi(text: string): string | undefined {
  for (const character of text) {
    const code = character.codePointAt(0)!;
    if (code === 0x0a) continue;
    if (
      (code >= 0x20 && code <= 0x7e) ||
      (code >= 0xa0 && code <= 0xff) ||
      WIN_ANSI_EXTRAS.has(code)
    )
      continue;
    return character;
  }
  return undefined;
}

export interface FontRequest {
  readonly family: string;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly text: string;
}

export interface ResolvedFont {
  readonly handle: number;
  /** Family actually used. */
  readonly family: string;
  /** Set when the font is not the one asked for. */
  readonly substitution?: string;
}

/** Why a request cannot be satisfied. */
export interface FontProblem {
  readonly path: string;
  readonly code: "unknown-font" | "font-unavailable";
  readonly message: string;
}

/**
 * Host-registered fonts and the fallback face, fetched on demand and loaded
 * into whichever PDFium document needs them.
 */
export class FontLibrary {
  readonly #registered: EditWorkerFont[] = [];
  readonly #fetch: (url: string) => Promise<Uint8Array>;
  readonly #bytes = new Map<string, Uint8Array>();
  readonly #coverage = new Map<Uint8Array, CmapCoverage>();
  /** Loaded PDFium font handles, keyed by document and font bytes. */
  readonly #handles = new Map<number, Map<Uint8Array, number>>();
  #fallbackUrl: string | undefined;
  #fallback: Uint8Array | undefined;

  constructor(fetchBytes: (url: string) => Promise<Uint8Array>) {
    this.#fetch = fetchBytes;
  }

  setFallbackUrl(url: string | undefined): void {
    this.#fallbackUrl = url;
  }

  register(fonts: readonly EditWorkerFont[]): void {
    this.#registered.length = 0;
    this.#registered.push(...fonts);
  }

  /** Whether `family` names a registered font. */
  hasFamily(family: string): boolean {
    return this.#registered.some((font) => sameFamily(font.family, family));
  }

  /**
   * Fetches whatever the given texts and families may need, so that
   * validation and drawing can run synchronously afterwards.
   */
  async prepare(
    requests: readonly { family: string; text: string }[],
  ): Promise<void> {
    const families = new Set(
      requests.map((request) => request.family.toLowerCase()),
    );
    const fonts = this.#registered.filter((font) =>
      families.has(font.family.toLowerCase()),
    );
    for (const font of fonts) await this.#bytesOf(font);
    // The fallback is only worth fetching for text no registered font covers.
    const needsFallback = requests.some(
      (request) =>
        // Text the standard fonts cannot encode, or a family only the
        // fallback can stand in for (an embedded font that lost a glyph).
        (firstNonWinAnsi(request.text) !== undefined ||
          (!isStandardFamily(request.family) &&
            !this.hasFamily(request.family))) &&
        !this.#coveringRegistered({ ...request, bold: false, italic: false }),
    );
    if (needsFallback && !this.#fallback && this.#fallbackUrl)
      this.#fallback = await this.#fetch(this.#fallbackUrl);
  }

  /** Checks a request against what is loaded; `undefined` means it will resolve. */
  problem(request: FontRequest): FontProblem | undefined {
    const standard = isStandardFamily(request.family);
    if (!standard && !this.hasFamily(request.family))
      return {
        path: "/style/fontFamily",
        code: "unknown-font",
        message: `Unknown font family ${request.family}; Helvetica, Times, Courier and registered fonts are available`,
      };
    if (standard && firstNonWinAnsi(request.text) === undefined)
      return undefined;
    if (this.#coveringRegistered(request)) return undefined;
    if (this.#fallback && this.#covers(this.#fallback, request.text))
      return undefined;
    const bad =
      firstNonWinAnsi(request.text) ?? firstUncovered(request.text, this);
    return {
      path: "/text",
      code: "font-unavailable",
      message: bad
        ? `No available font can draw "${bad}" (U+${bad.codePointAt(0)!.toString(16).toUpperCase().padStart(4, "0")})`
        : "No available font can draw the text",
    };
  }

  /** The font to draw `request` with in `document`; `prepare` must have run. */
  resolve(
    pdfium: Pdfium,
    document: number,
    request: FontRequest,
  ): ResolvedFont {
    const { lib } = pdfium;
    if (
      isStandardFamily(request.family) &&
      firstNonWinAnsi(request.text) === undefined
    ) {
      const name = standardFontName(
        request.family,
        request.bold,
        request.italic,
      )!;
      return {
        handle: lib.FPDFText_LoadStandardFont(document, name),
        family: name,
      };
    }
    const registered = this.#coveringRegistered(request);
    if (registered) {
      const exact =
        isBold(registered.font) === request.bold &&
        isItalic(registered.font) === request.italic;
      return {
        handle: this.#load(pdfium, document, registered.bytes),
        family: registered.font.family,
        ...(exact
          ? {}
          : {
              substitution: `${registered.font.family} has no matching bold or italic face`,
            }),
      };
    }
    if (!this.#fallback || !this.#covers(this.#fallback, request.text))
      throw new Error(`No font covers "${request.text}"`);
    return {
      handle: this.#load(pdfium, document, this.#fallback),
      family: FALLBACK_FAMILY,
      substitution: `${request.family} cannot draw this text; ${FALLBACK_FAMILY} is used`,
    };
  }

  /** Closes the fonts loaded into a document that is going away. */
  release(pdfium: Pdfium, document: number): void {
    const handles = this.#handles.get(document);
    if (!handles) return;
    for (const handle of handles.values()) pdfium.lib.FPDFFont_Close(handle);
    this.#handles.delete(document);
  }

  covers(bytes: Uint8Array, text: string): boolean {
    return this.#covers(bytes, text);
  }

  #coveringRegistered(
    request: FontRequest,
  ): { readonly font: EditWorkerFont; readonly bytes: Uint8Array } | undefined {
    const candidates = this.#registered
      .filter((font) => sameFamily(font.family, request.family))
      .map((font) => ({ font, bytes: this.#loadedBytes(font) }))
      .filter(
        (entry): entry is { font: EditWorkerFont; bytes: Uint8Array } =>
          entry.bytes !== undefined && this.#covers(entry.bytes, request.text),
      );
    // The closest style wins: matching bold and italic, then matching one.
    const score = (font: EditWorkerFont): number =>
      (isBold(font) === request.bold ? 2 : 0) +
      (isItalic(font) === request.italic ? 1 : 0);
    return candidates.sort((a, b) => score(b.font) - score(a.font))[0];
  }

  #covers(bytes: Uint8Array, text: string): boolean {
    let coverage = this.#coverage.get(bytes);
    if (!coverage) {
      coverage = parseCmap(bytes);
      this.#coverage.set(bytes, coverage);
    }
    for (const character of text) {
      const code = character.codePointAt(0)!;
      if (code === 0x0a || code === 0x20) continue;
      if (!coverage.has(code)) return false;
    }
    return true;
  }

  #load(pdfium: Pdfium, document: number, bytes: Uint8Array): number {
    let handles = this.#handles.get(document);
    if (!handles) {
      handles = new Map();
      this.#handles.set(document, handles);
    }
    let handle = handles.get(bytes);
    if (handle === undefined) {
      const pointer = pdfium.writeBytes(bytes);
      try {
        handle = pdfium.lib.FPDFText_LoadFont(
          document,
          pointer,
          bytes.byteLength,
          FONT_TRUETYPE,
          true,
        );
      } finally {
        pdfium.free(pointer);
      }
      if (!handle) throw new Error("PDFium could not load the font");
      handles.set(bytes, handle);
    }
    return handle;
  }

  #loadedBytes(font: EditWorkerFont): Uint8Array | undefined {
    return typeof font.source === "string"
      ? this.#bytes.get(font.source)
      : (this.#bytes.get(sourceKey(font)) ?? new Uint8Array(font.source));
  }

  async #bytesOf(font: EditWorkerFont): Promise<Uint8Array> {
    const key = sourceKey(font);
    let bytes = this.#bytes.get(key);
    if (!bytes) {
      bytes =
        typeof font.source === "string"
          ? await this.#fetch(font.source)
          : new Uint8Array(font.source);
      this.#bytes.set(key, bytes);
    }
    return bytes;
  }
}

function sourceKey(font: EditWorkerFont): string {
  return typeof font.source === "string"
    ? font.source
    : `${font.family}:${font.weight}:${font.style}:${font.source.byteLength}`;
}

function sameFamily(a: string, b: string): boolean {
  return a.toLowerCase() === b.toLowerCase();
}

function isBold(font: EditWorkerFont): boolean {
  return font.weight >= 600;
}

function isItalic(font: EditWorkerFont): boolean {
  return font.style !== "normal";
}

function firstUncovered(
  text: string,
  library: FontLibrary,
): string | undefined {
  void library;
  for (const character of text)
    if (character !== "\n" && character !== " ") return character;
  return undefined;
}

/** The code points a TrueType font maps to glyphs, from its cmap table. */
/**
 * Whether glyphs of a TrueType font have outline data: a subset font can keep
 * a cmap entry for a glyph it emptied. Returns undefined without a glyf table.
 */
function glyphOutlines(
  view: DataView,
  tables: ReadonlyMap<string, { offset: number; length: number }>,
): ((glyph: number) => boolean) | undefined {
  const head = tables.get("head");
  const maxp = tables.get("maxp");
  const loca = tables.get("loca");
  if (!head || !maxp || !loca || !tables.has("glyf")) return undefined;
  const longOffsets = view.getInt16(head.offset + 50) === 1;
  const count = view.getUint16(maxp.offset + 4);
  return (glyph) => {
    if (glyph < 0 || glyph >= count) return false;
    const at = loca.offset + glyph * (longOffsets ? 4 : 2);
    const start = longOffsets ? view.getUint32(at) : view.getUint16(at) * 2;
    const end = longOffsets
      ? view.getUint32(at + 4)
      : view.getUint16(at + 2) * 2;
    return end > start;
  };
}

export interface CmapCoverage {
  has(codePoint: number): boolean;
  /** Mapped to a glyph with outline data; false for an emptied glyph of a subset font. */
  drawable(codePoint: number): boolean;
  /** The glyph id a code point maps to, when the font's cmap was parsed. */
  glyph(codePoint: number): number | undefined;
}

export function parseCmap(bytes: Uint8Array): CmapCoverage {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const ranges: [number, number][] = [];
  /** Glyph ids by code point, for the outline check; filled for format 4 and 12. */
  const glyphs = new Map<number, number>();
  let outlines: ((glyph: number) => boolean) | undefined;
  try {
    const tables = view.getUint16(4);
    let cmapOffset = -1;
    const offsets = new Map<string, { offset: number; length: number }>();
    for (let index = 0; index < tables; index += 1) {
      const record = 12 + index * 16;
      const tag = String.fromCharCode(
        bytes[record]!,
        bytes[record + 1]!,
        bytes[record + 2]!,
        bytes[record + 3]!,
      );
      offsets.set(tag, {
        offset: view.getUint32(record + 8),
        length: view.getUint32(record + 12),
      });
      if (tag === "cmap") cmapOffset = view.getUint32(record + 8);
    }
    outlines = glyphOutlines(view, offsets);
    if (cmapOffset < 0)
      return {
        has: () => false,
        drawable: () => false,
        glyph: () => undefined,
      };
    const subtables = view.getUint16(cmapOffset + 2);
    let best: { offset: number; format: number } | undefined;
    for (let index = 0; index < subtables; index += 1) {
      const record = cmapOffset + 4 + index * 8;
      const platform = view.getUint16(record);
      const encoding = view.getUint16(record + 2);
      const offset = cmapOffset + view.getUint32(record + 4);
      const format = view.getUint16(offset);
      const unicode =
        platform === 0 ||
        (platform === 3 && (encoding === 1 || encoding === 10));
      if (!unicode || (format !== 4 && format !== 12)) continue;
      if (!best || format > best.format) best = { offset, format };
    }
    if (!best)
      return {
        has: () => false,
        drawable: () => false,
        glyph: () => undefined,
      };
    if (best.format === 12) {
      const groups = view.getUint32(best.offset + 12);
      for (let index = 0; index < groups; index += 1) {
        const group = best.offset + 16 + index * 12;
        const start = view.getUint32(group);
        const end = view.getUint32(group + 4);
        const startGlyph = view.getUint32(group + 8);
        ranges.push([start, end]);
        for (let code = start; code <= end && code - start < 0x10000; code += 1)
          glyphs.set(code, startGlyph + (code - start));
      }
    } else {
      const segments = view.getUint16(best.offset + 6) / 2;
      const ends = best.offset + 14;
      const starts = ends + segments * 2 + 2;
      const deltas = starts + segments * 2;
      const rangeOffsets = deltas + segments * 2;
      for (let index = 0; index < segments; index += 1) {
        const end = view.getUint16(ends + index * 2);
        const start = view.getUint16(starts + index * 2);
        if (start === 0xffff) continue;
        ranges.push([start, end]);
        const delta = view.getInt16(deltas + index * 2);
        const rangeOffset = view.getUint16(rangeOffsets + index * 2);
        for (let code = start; code <= end; code += 1) {
          let glyph: number;
          if (rangeOffset === 0) glyph = (code + delta) & 0xffff;
          else {
            const at =
              rangeOffsets + index * 2 + rangeOffset + (code - start) * 2;
            const raw = view.getUint16(at);
            glyph = raw === 0 ? 0 : (raw + delta) & 0xffff;
          }
          glyphs.set(code, glyph);
        }
      }
    }
  } catch {
    return { has: () => false, drawable: () => false, glyph: () => undefined };
  }
  ranges.sort((a, b) => a[0] - b[0]);
  return {
    glyph(codePoint) {
      return glyphs.get(codePoint);
    },
    drawable(codePoint) {
      const glyph = glyphs.get(codePoint);
      if (glyph === undefined || glyph === 0) return false;
      // Without a glyf table (CFF outlines) a mapped glyph is taken as drawn.
      return outlines ? outlines(glyph) : true;
    },
    has(codePoint) {
      let low = 0;
      let high = ranges.length - 1;
      while (low <= high) {
        const middle = (low + high) >> 1;
        const [start, end] = ranges[middle]!;
        if (codePoint < start) high = middle - 1;
        else if (codePoint > end) low = middle + 1;
        else return true;
      }
      return false;
    },
  };
}

/**
 * Measures text with PDFium's own metrics by building throw-away text
 * objects, so layout agrees with what PDFium draws. Advances are cached per
 * character and font at 1000 pt and scale linearly.
 */
export class TextMeasurer {
  readonly #pdfium: Pdfium;
  readonly #document: number;
  readonly #advances = new Map<number, Map<string, number>>();
  readonly #metrics = new Map<number, { ascent: number; descent: number }>();

  constructor(pdfium: Pdfium, document: number) {
    this.#pdfium = pdfium;
    this.#document = document;
  }

  /** Width of `text` in points at `fontSize`, without kerning. */
  advance(font: number, fontSize: number, text: string): number {
    let total = 0;
    for (const character of text) total += this.#advanceOf(font, character);
    return (total * fontSize) / 1000;
  }

  /** Ascent above and descent below the baseline, in points, both positive. */
  metrics(font: number, fontSize: number): { ascent: number; descent: number } {
    let unit = this.#metrics.get(font);
    if (!unit) {
      const { lib } = this.#pdfium;
      const ascent =
        this.#pdfium.readNumbers(1, "float", ([pointer]) =>
          lib.FPDFFont_GetAscent(font, 1000, pointer!),
        )?.[0] ?? 750;
      const descent =
        this.#pdfium.readNumbers(1, "float", ([pointer]) =>
          lib.FPDFFont_GetDescent(font, 1000, pointer!),
        )?.[0] ?? -250;
      unit = { ascent, descent: Math.abs(descent) };
      this.#metrics.set(font, unit);
    }
    return {
      ascent: (unit.ascent * fontSize) / 1000,
      descent: (unit.descent * fontSize) / 1000,
    };
  }

  #advanceOf(font: number, character: string): number {
    let cache = this.#advances.get(font);
    if (!cache) {
      cache = new Map();
      this.#advances.set(font, cache);
    }
    let advance = cache.get(character);
    if (advance === undefined) {
      // Ink boxes do not give advances, but the box of "H?H" minus the box
      // of "HH" is exactly the advance of "?", for any "?" including space.
      advance =
        this.#inkWidth(font, `H${character}H`) - this.#inkWidth(font, "HH");
      cache.set(character, advance);
    }
    return advance;
  }

  #inkWidth(font: number, text: string): number {
    const { lib } = this.#pdfium;
    const object = lib.FPDFPageObj_CreateTextObj(this.#document, font, 1000);
    try {
      const wide = this.#pdfium.writeWideString(text);
      try {
        lib.FPDFText_SetText(object, wide);
      } finally {
        this.#pdfium.free(wide);
      }
      const bounds = this.#pdfium.readNumbers(4, "float", ([l, b, r, t]) =>
        lib.FPDFPageObj_GetBounds(object, l!, b!, r!, t!),
      );
      return bounds ? bounds[2]! - bounds[0]! : 0;
    } finally {
      lib.FPDFPageObj_Destroy(object);
    }
  }
}
