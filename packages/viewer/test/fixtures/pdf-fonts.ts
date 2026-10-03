import { cffFont, type CffGlyph } from "./cff-font.js";
import { fixturePdfium } from "./pdf-builder.js";

/*
 * PDFs whose text is drawn in an embedded CFF font, written out by hand:
 * PDFium embeds TrueType fonts only, and PDF producers embed CFF subsets.
 */

const GLYPHS: readonly CffGlyph[] = [
  { name: "space", width: 250 },
  { name: "A", width: 600 },
  // The width before a subroutine call, as subroutinized subsets write it.
  { name: "B", width: 550, viaSubroutine: true },
  // A name from the font's own strings, for a character outside WinAnsi.
  { name: "uni0416", width: 700 },
  // No width: the font's default one.
  { name: "a" },
  // Beyond the BMP, so the cmap needs format 12.
  { name: "u1F600", width: 1000 },
];

/** The advance widths `cffTextPdf`'s font gives its glyphs, in font units. */
export const CFF_TEXT_WIDTHS: Readonly<Record<string, number>> = {
  " ": 250,
  A: 600,
  B: 550,
  Ж: 700,
  a: 500,
  "\u{1F600}": 1000,
};

/**
 * Two lines in an embedded Type1C subset, "AB" and "BA", and one in
 * Helvetica, which is not embedded.
 */
export function cffTextPdf(): Uint8Array {
  // A space in the CFF name, which browsers refuse there; the PDF's BaseFont has none.
  const program = cffFont({ name: "ABCDEF+Test Face", glyphs: GLYPHS });
  return pdf([
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R /F2 7 0 R >> >> /Contents 8 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /ABCDEF+TestFace /FirstChar 65 /LastChar 66 /Widths [600 550] /Encoding /WinAnsiEncoding /FontDescriptor 5 0 R >>",
    descriptor("ABCDEF+TestFace"),
    stream(program, "/Subtype /Type1C"),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    stream(
      latin1(
        "BT /F1 24 Tf 72 700 Td (AB) Tj ET\nBT /F1 24 Tf 72 650 Td (BA) Tj ET\nBT /F2 12 Tf 72 600 Td (Helvetica) Tj ET",
      ),
    ),
  ]);
}

/**
 * Three lines in an inked CFF subset: "AB", "BA" and "A B". The PDF gives
 * widths to the space, "0", "A" and "B"; the program has glyphs for the
 * space, "A", "B" and "a". So "0" has a width and no glyph, and "a" a glyph
 * and no width.
 */
export function cffInkTextPdf(): Uint8Array {
  const program = cffFont({
    name: "ABCDEF+InkFace",
    glyphs: [
      { name: "space", width: 250 },
      { name: "A", width: 600, ink: 700 },
      { name: "B", width: 550, ink: 400, viaSubroutine: true },
      { name: "a", width: 500, ink: 300 },
    ],
  });
  const widths = Array.from({ length: 35 }, (_, offset) =>
    offset === 0
      ? 250
      : offset === 16
        ? 500
        : offset === 33
          ? 600
          : offset === 34
            ? 550
            : 0,
  );
  return pdf([
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 7 0 R >>",
    `<< /Type /Font /Subtype /Type1 /BaseFont /ABCDEF+InkFace /FirstChar 32 /LastChar 66 /Widths [${widths.join(" ")}] /Encoding /WinAnsiEncoding /FontDescriptor 5 0 R >>`,
    descriptor("ABCDEF+InkFace"),
    stream(program, "/Subtype /Type1C"),
    stream(
      latin1(
        "BT /F1 24 Tf 72 700 Td (AB) Tj ET\nBT /F1 24 Tf 72 650 Td (BA) Tj ET\nBT /F1 24 Tf 72 600 Td (A B) Tj ET",
      ),
    ),
  ]);
}

/** One line in a CID-keyed CFF font through Identity-H. */
export function cidTextPdf(): Uint8Array {
  const program = cffFont({
    name: "ABCDEF+TestCid",
    glyphs: GLYPHS,
    cid: true,
  });
  return pdf([
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 8 0 R >>",
    "<< /Type /Font /Subtype /Type0 /BaseFont /ABCDEF+TestCid /Encoding /Identity-H /DescendantFonts [7 0 R] >>",
    descriptor("ABCDEF+TestCid"),
    stream(program, "/Subtype /CIDFontType0C"),
    "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /ABCDEF+TestCid /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 5 0 R /DW 500 >>",
    stream(latin1("BT /F1 24 Tf 72 700 Td <00020003> Tj ET")),
  ]);
}

/** One line in a Type 3 font, whose glyphs are drawings with no font program. */
export function type3TextPdf(): Uint8Array {
  return pdf([
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 6 0 R >>",
    "<< /Type /Font /Subtype /Type3 /FontBBox [0 0 1000 1000] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /square 5 0 R >> /Encoding << /Type /Encoding /Differences [65 /square] >> /FirstChar 65 /LastChar 65 /Widths [1000] /Resources << >> >>",
    stream(latin1("1000 0 0 0 1000 1000 d1 0 0 1000 1000 re f")),
    stream(latin1("BT /F1 24 Tf 72 700 Td (AAA) Tj ET")),
  ]);
}

function descriptor(name: string): string {
  return `<< /Type /FontDescriptor /FontName /${name} /Flags 32 /FontBBox [0 -200 1000 800] /ItalicAngle 0 /Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile3 6 0 R >>`;
}

type PdfObject = string | Uint8Array;

function stream(data: Uint8Array, dictionary = ""): Uint8Array {
  return concat([
    latin1(`<< ${dictionary} /Length ${data.length} >>\nstream\n`),
    data,
    latin1("\nendstream"),
  ]);
}

function pdf(objects: readonly PdfObject[]): Uint8Array {
  const parts: Uint8Array[] = [latin1("%PDF-1.7\n")];
  let length = parts[0]!.length;
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(length);
    const object = concat([
      latin1(`${index + 1} 0 obj\n`),
      typeof body === "string" ? latin1(body) : body,
      latin1("\nendobj\n"),
    ]);
    parts.push(object);
    length += object.length;
  });
  const xref = [
    `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`,
    ...offsets.map(
      (offset) => `${String(offset).padStart(10, "0")} 00000 n \n`,
    ),
    `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${length}\n%%EOF\n`,
  ].join("");
  parts.push(latin1(xref));
  return concat(parts);
}

function latin1(text: string): Uint8Array {
  return Uint8Array.from(text, (character) => character.charCodeAt(0));
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

/**
 * A TrueType font the way PDF producers subset it: without the OS/2, name
 * and post tables, and with the cmap off a four-byte boundary. FreeType and
 * PDFium draw it; a browser refuses it as it is.
 */
export function strippedTrueType(font: Uint8Array): Uint8Array {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const tables: { tag: string; data: Uint8Array }[] = [];
  for (let index = 0; index < view.getUint16(4); index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...font.subarray(record, record + 4));
    const offset = view.getUint32(record + 8);
    if (!["OS/2", "name", "post"].includes(tag))
      tables.push({
        tag,
        data: font.subarray(offset, offset + view.getUint32(record + 12)),
      });
  }
  const out: number[] = [
    ...u32(0x00010000),
    ...u16(tables.length),
    ...u16(0),
    ...u16(0),
    ...u16(0),
  ];
  let offset = 12 + tables.length * 16;
  const placed = tables.map(({ tag, data }) => {
    // Two bytes of padding put the cmap at an offset of 2 mod 4.
    if (tag === "cmap") offset += 2;
    const at = offset;
    offset += data.length;
    return { tag, data, at };
  });
  for (const { tag, data, at } of placed)
    out.push(...latin1(tag), ...u32(0), ...u32(at), ...u32(data.length));
  const file = new Uint8Array(offset);
  file.set(out);
  for (const { data, at } of placed) file.set(data, at);
  return file;
}

/** One line in a TrueType font embedded by PDFium from `font` as given. */
export async function trueTypeTextPdf(
  font: Uint8Array,
  text: string,
): Promise<Uint8Array> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.createDocument();
  const data = pdfium.writeBytes(font);
  try {
    const page = lib.FPDFPage_New(document.handle, 0, 612, 792);
    const handle = lib.FPDFText_LoadFont(
      document.handle,
      data,
      font.length,
      1,
      true,
    );
    if (!handle) throw new Error("PDFium did not load the font");
    const object = lib.FPDFPageObj_CreateTextObj(document.handle, handle, 14);
    const wide = pdfium.writeWideString(text);
    lib.FPDFText_SetText(object, wide);
    pdfium.free(wide);
    lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 72, 700);
    lib.FPDFPage_InsertObject(page, object);
    lib.FPDFPage_GenerateContent(page);
    lib.FPDF_ClosePage(page);
    return document.save("full");
  } finally {
    pdfium.free(data);
    document.close();
  }
}

function u16(value: number): number[] {
  return [(value >> 8) & 0xff, value & 0xff];
}

function u32(value: number): number[] {
  return [
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff,
  ];
}
