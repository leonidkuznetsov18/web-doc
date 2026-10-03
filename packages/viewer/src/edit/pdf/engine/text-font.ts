import type { TextFont } from "../types.js";
import { glyphUnicode, readCff } from "./cff.js";
import { textStyle } from "./elements.js";
import { cffOpenType, completeSfnt, hasUnicodeCmap } from "./opentype.js";
import type { Pdfium } from "./pdfium.js";

/*
 * The browser face of the font a text object is drawn in. Faces are cached by
 * font program, so every element in one font shares one face and one key.
 */

export type TextFace = Omit<TextFont, "elementId">;

export function textFaceOf(
  pdfium: Pdfium,
  object: number,
  cache: Map<string, TextFace>,
): TextFace {
  const { lib } = pdfium;
  const font = lib.FPDFTextObj_GetFont(object);
  const family = textStyle(pdfium, object).fontFamily;
  const baseName = pdfium.readUtf8String((buffer, bytes) =>
    lib.FPDFFont_GetBaseFontName(font, buffer, bytes),
  );
  if (!lib.FPDFFont_GetIsEmbedded(font))
    return { key: `name:${baseName}`, family, missing: "not-embedded" };
  // A Type 3 font is embedded as drawings, with no program to read.
  const program = fontProgram(pdfium, font);
  if (!program)
    return { key: `embedded:${baseName}`, family, missing: "unreadable" };
  const key = `${baseName}#${fnv1a(program).toString(16)}`;
  const known = cache.get(key);
  if (known) return { ...known, family };
  const face: TextFace = { key, family, ...faceOf(program, baseName) };
  cache.set(key, face);
  return face;
}

function faceOf(
  program: Uint8Array,
  baseName: string,
): Pick<TextFace, "face" | "missing"> {
  const tag = String.fromCharCode(...program.subarray(0, 4));
  const truetype =
    tag === "true" ||
    (program[0] === 0 &&
      program[1] === 1 &&
      program[2] === 0 &&
      program[3] === 0);
  if (truetype || tag === "OTTO") {
    if (!hasUnicodeCmap(program)) return { missing: "no-unicode" };
    const data = completeSfnt(program, baseName);
    return data
      ? { face: { data, format: truetype ? "truetype" : "opentype" } }
      : { missing: "unreadable" };
  }
  // Type 1 programs start with "%!" in clear text, or with a PFB segment.
  if (tag.startsWith("%!") || program[0] === 0x80) return { missing: "type1" };
  const cff = readCff(program);
  if (!cff) return { missing: "unreadable" };
  if (cff.cid) return { missing: "cid-keyed" };
  const unicode = new Map<number, number>();
  cff.glyphNames.forEach((name, glyph) => {
    const code = glyphUnicode(name);
    if (code !== undefined && !unicode.has(code)) unicode.set(code, glyph);
  });
  const data = cffOpenType(program, cff, unicode);
  return data
    ? { face: { data, format: "opentype" } }
    : { missing: "no-unicode" };
}

/** The font program a PDF font object embeds, or `undefined` for a font it does not embed. */
function fontProgram(pdfium: Pdfium, font: number): Uint8Array | undefined {
  const { lib } = pdfium;
  const size = pdfium.readNumbers(1, "i32", ([out]) =>
    lib.FPDFFont_GetFontData(font, 0, 0, out!),
  )?.[0];
  if (!size) return undefined;
  const buffer = pdfium.malloc(size);
  const out = pdfium.malloc(4);
  try {
    return lib.FPDFFont_GetFontData(font, buffer, size, out)
      ? pdfium.readBytes(buffer, size)
      : undefined;
  } finally {
    pdfium.free(out);
    pdfium.free(buffer);
  }
}

/** FNV-1a over the program: a cheap, stable key for one font. */
function fnv1a(bytes: Uint8Array): number {
  let hash = 0x811c9dc5;
  for (const byte of bytes) hash = Math.imul(hash ^ byte, 0x01000193);
  return hash >>> 0;
}
