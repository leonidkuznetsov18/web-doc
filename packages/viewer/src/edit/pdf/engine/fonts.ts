import type { Pdfium } from "./pdfium.js";

/*
 * Fonts for text web-doc writes into a PDF. The MVP uses the standard 14
 * fonts, which every reader has; text must then be WinAnsi-encodable.
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
