import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";

import { fontBytes, parseCmap } from "../src/edit/pdf/engine/fonts.js";
import { extractPageText, fixturePdfium } from "./fixtures/pdf-builder.js";

import { pdfSession } from "./fixtures/pdf-session.js";

const lines = [
  "Alpha beta gamma delta",
  "Alpha beta gamma delta",
  "Iota kappa lambda mu",
] as const;

type Spacing =
  | "default"
  | "character"
  | "tiny character"
  | "word"
  | "tiny word"
  | "positioned words"
  | "horizontal scale"
  | "tiny horizontal scale"
  | "pair kerning"
  | "compensated pair kerning"
  | "accumulating pair adjustments";

/** Raw PDF text operators exercise spacing unavailable in our PDFium builder. */
function paragraphPdf(
  spacing: Spacing,
  textLines: readonly string[] = lines,
): Uint8Array {
  const state = {
    default: "",
    character: "0.7 Tc",
    "tiny character": "0.005 Tc",
    word: "4 Tw",
    "tiny word": "0.005 Tw",
    "positioned words": "",
    "horizontal scale": "80 Tz",
    "tiny horizontal scale": "99.99 Tz",
    "pair kerning": "",
    "compensated pair kerning": "",
    "accumulating pair adjustments": "",
  }[spacing];
  const content = textLines
    .map((text, index) => {
      const show =
        spacing === "positioned words"
          ? `[(${text.split(" ").join(") -600 (")})] TJ`
          : spacing === "pair kerning"
            ? `[(${text.slice(0, 1)}) 60 (${text.slice(1)})] TJ`
            : spacing === "compensated pair kerning"
              ? `[(${text.slice(0, 1)}) 60 (${text.slice(1, 2)}) -60 (${text.slice(2)})] TJ`
              : spacing === "accumulating pair adjustments"
                ? `[(${text.slice(0, 1)}) 60 (${text.slice(1, 2)}) 60 (${text.slice(2, 3)}) 60 (${text.slice(3)})] TJ`
                : `(${text}) Tj`;
      return `BT /F1 12 Tf ${state} 1 0 0 1 72 ${720 - index * 20} Tm ${show} ET`;
    })
    .join("\n");
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`,
  ];
  let pdf = "%PDF-1.7\n";
  const offsets: number[] = [];
  for (const [index, object] of objects.entries()) {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  }
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets)
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new Uint8Array(Buffer.from(pdf));
}

describe("native PDF paragraph spacing eligibility", () => {
  for (const spacing of [
    "character",
    "tiny character",
    "word",
    "tiny word",
    "positioned words",
    "horizontal scale",
    "tiny horizontal scale",
    "accumulating pair adjustments",
  ] as const) {
    it(`keeps ${spacing} spacing independently editable instead of silently normalizing it`, async () => {
      const original = paragraphPdf(spacing);
      const { session, end } = await pdfSession(original);
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        assert.equal(elements.length, lines.length);
        for (const element of elements) {
          assert.equal(
            element.textEditingTarget,
            undefined,
            element.text ?? element.id,
          );
          assert.equal(
            (await session.getTextParagraph(element.id)).item,
            undefined,
          );
          assert.ok(element.operations.includes("replaceText"));
        }
        assert.equal(session.state.dirty, false);
        assert.deepEqual((await session.save()).bytes, original);
      } finally {
        await end();
      }
    });
  }

  it("refuses tiny systematic tracking even when there are no word spaces", async () => {
    const textLines = lines.map((line) => line.replaceAll(" ", ""));
    const { session, end } = await pdfSession(
      paragraphPdf("tiny character", textLines),
    );
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      assert.equal(elements.length, textLines.length);
      assert.ok(elements.every((element) => !element.textEditingTarget));
    } finally {
      await end();
    }
  });

  for (const spacing of ["pair kerning", "compensated pair kerning"] as const) {
    it(`keeps bounded ${spacing} editable, preserves its color-only layout, and saves reflowed text`, async () => {
      const textLines = [
        "AVATAR beta gamma delta",
        "AVATAR beta gamma delta",
        "AVATAR final words",
      ];
      const { session, end } = await pdfSession(
        paragraphPdf(spacing, textLines),
      );
      try {
        const elements = (await session.getElements({ pageIndex: 0 })).items;
        const first = elements[0];
        assert.ok(first);
        const paragraph = (await session.getTextParagraph(first.id)).item;
        assert.ok(
          paragraph,
          "ordinary pair kerning must not disable paragraph editing",
        );
        assert.equal(paragraph.text, textLines.join(" "));
        const layout = (await session.getTextLayout(paragraph.id)).item;
        await session.setTextStyle({
          target: paragraph.id,
          style: { color: "#c02040" },
        });
        assert.ok(layout);
        assert.deepEqual((await session.getTextLayout(paragraph.id)).item, {
          ...layout,
          lines: layout.lines.map((line) => ({ ...line, color: "#c02040" })),
        });
        const replacement = "AVATAR retains the whole edited paragraph.";
        const receipt = await session.replaceParagraphText({
          target: paragraph.id,
          text: replacement,
        });
        assert.deepEqual(receipt.warnings, []);
        const saved = await session.save();
        assert.equal(
          (await extractPageText(saved.bytes, 0)).replace(/\s+/gu, " ").trim(),
          replacement,
        );
        const reopened = await pdfSession(saved.bytes);
        try {
          assert.equal(
            (await reopened.session.getTextParagraph(paragraph.id)).item?.text,
            replacement,
          );
        } finally {
          await reopened.end();
        }
      } finally {
        await end();
      }
    });
  }

  it("exposes otherwise identical normally spaced wrapped lines as one paragraph", async () => {
    const original = paragraphPdf("default");
    const { session, end } = await pdfSession(original);
    try {
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      const first = elements[0];
      assert.ok(first);
      const paragraph = (await session.getTextParagraph(first.id)).item;
      assert.ok(paragraph);
      assert.equal(paragraph.text, lines.join(" "));
      assert.deepEqual(
        paragraph.memberIds,
        elements.map((element) => element.id),
      );
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });
});

/** Remove only the space mapping from the fixture font, as in imported subsets. */
function withoutSpaceMapping(font: Uint8Array): Uint8Array {
  const bytes = font.slice();
  const view = new DataView(bytes.buffer);
  let cmap = -1;
  for (let index = 0; index < view.getUint16(4); index += 1) {
    const record = 12 + index * 16;
    if (String.fromCharCode(...bytes.subarray(record, record + 4)) === "cmap")
      cmap = view.getUint32(record + 8);
  }
  assert.ok(cmap > 0);
  const visited = new Set<number>();
  for (let index = 0; index < view.getUint16(cmap + 2); index += 1) {
    const subtable = cmap + view.getUint32(cmap + 4 + index * 8 + 4);
    if (visited.has(subtable)) continue;
    visited.add(subtable);
    const format = view.getUint16(subtable);
    assert.equal(format, 4, "the packaged fixture uses a format 4 cmap");
    const count = view.getUint16(subtable + 6) / 2;
    const ends = subtable + 14;
    const starts = ends + count * 2 + 2;
    const ranges = starts + count * 4;
    for (let segment = 0; segment < count; segment += 1) {
      if (view.getUint16(starts + segment * 2) !== 32) continue;
      assert.ok(view.getUint16(ends + segment * 2) > 32);
      view.setUint16(starts + segment * 2, 33);
      const offset = view.getUint16(ranges + segment * 2);
      if (offset !== 0) view.setUint16(ranges + segment * 2, offset + 2);
    }
  }
  assert.equal(parseCmap(bytes).glyph(32), undefined);
  assert.equal(parseCmap(bytes).glyph(65), parseCmap(font).glyph(65));
  return bytes;
}

async function embeddedParagraph(font: Uint8Array): Promise<Uint8Array> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.createDocument();
  const data = pdfium.writeBytes(font);
  try {
    const page = lib.FPDFPage_New(document.handle, 0, 612, 792);
    try {
      const handle = lib.FPDFText_LoadFont(
        document.handle,
        data,
        font.length,
        1,
        true,
      );
      assert.ok(handle);
      for (const [index, text] of [
        "AlphaBetaGammaDelta",
        "AlphaBetaGammaDelta",
        "IotaKappaLambda",
      ].entries()) {
        const object = lib.FPDFPageObj_CreateTextObj(
          document.handle,
          handle,
          12,
        );
        const wide = pdfium.writeWideString(text);
        try {
          assert.ok(lib.FPDFText_SetText(object, wide));
        } finally {
          pdfium.free(wide);
        }
        lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 72, 720 - index * 20);
        lib.FPDFPage_InsertObject(page, object);
      }
      assert.ok(lib.FPDFPage_GenerateContent(page));
    } finally {
      lib.FPDF_ClosePage(page);
    }
    return document.save("full");
  } finally {
    pdfium.free(data);
    document.close();
  }
}

async function nativeFonts(bytes: Uint8Array) {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.openDocument(bytes);
  try {
    const page = lib.FPDF_LoadPage(document.handle, 0);
    try {
      return Array.from(
        { length: lib.FPDFPage_CountObjects(page) },
        (_, index) => {
          const object = lib.FPDFPage_GetObject(page, index);
          const font = lib.FPDFTextObj_GetFont(object);
          const bytes = fontBytes(pdfium, font);
          assert.ok(bytes);
          return {
            bytes,
            embedded: Boolean(lib.FPDFFont_GetIsEmbedded(font)),
            family: pdfium.readUtf8String((out, length) =>
              lib.FPDFFont_GetBaseFontName(font, out, length),
            ),
          };
        },
      );
    } finally {
      lib.FPDF_ClosePage(page);
    }
  } finally {
    document.close();
  }
}

describe("paragraph replacement with embedded space glyphs", () => {
  for (const coverage of ["mapped empty space", "missing space"] as const) {
    it(`preserves native color resources and ${coverage === "missing space" ? "reports substitution for" : "reuses"} a font with ${coverage}`, async () => {
      const complete = new Uint8Array(
        readFileSync(
          new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
        ),
      );
      assert.ok(parseCmap(complete).glyph(32));
      assert.equal(
        parseCmap(complete).drawable(32),
        false,
        "a legitimate space has no outline",
      );
      const original = await embeddedParagraph(
        coverage === "missing space" ? withoutSpaceMapping(complete) : complete,
      );
      const fonts = await nativeFonts(original);
      const { session, end } = await pdfSession(original, {
        fallbackFont: true,
      });
      try {
        const paragraph = (await session.getTextParagraph("p0:o0")).item;
        assert.ok(paragraph);
        await session.setTextStyle({
          target: paragraph.id,
          style: { color: "#c02040" },
        });
        assert.deepEqual(
          await nativeFonts((await session.save()).bytes),
          fonts,
        );
        const replacement = "Alpha beta gamma delta.";
        const receipt = await session.replaceParagraphText({
          target: paragraph.id,
          text: replacement,
        });
        assert.deepEqual(
          receipt.warnings.map((warning) => warning.code),
          coverage === "missing space" ? ["font-substitution"] : [],
        );
        const saved = await session.save();
        assert.equal(
          (await extractPageText(saved.bytes, 0)).replace(/\s+/gu, " ").trim(),
          replacement,
        );
        for (const savedFont of await nativeFonts(saved.bytes)) {
          if (savedFont.embedded) {
            assert.ok(
              parseCmap(savedFont.bytes).glyph(32),
              "embedded spaces must not use .notdef",
            );
          } else {
            assert.equal(
              savedFont.family,
              "Helvetica",
              "the standard fallback encodes WinAnsi spaces",
            );
          }
          if (coverage === "missing space")
            assert.notDeepEqual(savedFont.bytes, fonts[0]?.bytes);
        }
        const reopened = await pdfSession(saved.bytes);
        try {
          assert.equal(
            (await reopened.session.getTextParagraph(paragraph.id)).item?.text,
            replacement,
          );
        } finally {
          await reopened.end();
        }
      } finally {
        await end();
      }
    });
  }
});
