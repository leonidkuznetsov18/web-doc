import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { describe, it } from "node:test";

import {
  cffWithSafeName,
  glyphUnicode,
  readCff,
} from "../src/edit/pdf/engine/cff.js";
import { parseCmap } from "../src/edit/pdf/engine/fonts.js";
import {
  cffOpenType,
  completeSfnt,
  hasUnicodeCmap,
} from "../src/edit/pdf/engine/opentype.js";
import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import {
  loadPdfEditEngine,
  type PdfEditEngineClient,
} from "../src/edit/pdf/provider.js";
import type { PdfElement, TextFont } from "../src/index.js";
import { defaultResourceLimits } from "../src/index.js";
import { cffFont } from "./fixtures/cff-font.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";
import {
  CFF_TEXT_WIDTHS,
  cffTextPdf,
  cidTextPdf,
  strippedTrueType,
  trueTypeTextPdf,
  type3TextPdf,
} from "./fixtures/pdf-fonts.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * ACTION-922: the face a host types an element's text in. A CFF subset, the
 * font most PDF producers embed, becomes an OpenType font a browser loads; a
 * TrueType program is handed over as embedded; anything else is reported as
 * missing, never substituted.
 */

const ttf = new Uint8Array(
  readFileSync(
    new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
  ),
);
const signal = new AbortController().signal;

async function engineFor(original: Uint8Array): Promise<PdfEditEngineClient> {
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      decodeImage: async () => {
        throw new Error("no images");
      },
      fetchBytes: async (url) => {
        throw new Error(`No font at ${url}`);
      },
    }),
  );
  return loadPdfEditEngine(
    original,
    { format: "pdf", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker },
  );
}

/** Where each table of an sfnt file starts. */
function offsetsOf(font: Uint8Array): number[] {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  return Array.from({ length: view.getUint16(4) }, (_, index) =>
    view.getUint32(12 + index * 16 + 8),
  );
}

/** The tables of an sfnt file by tag. */
function tablesOf(font: Uint8Array): Map<string, Uint8Array> {
  const view = new DataView(font.buffer, font.byteOffset, font.byteLength);
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < view.getUint16(4); index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...font.subarray(record, record + 4));
    const offset = view.getUint32(record + 8);
    tables.set(
      tag,
      font.subarray(offset, offset + view.getUint32(record + 12)),
    );
  }
  return tables;
}

/** Advance widths from hmtx, by glyph id. */
function advances(font: Uint8Array): number[] {
  const tables = tablesOf(font);
  const hhea = tables.get("hhea")!;
  const hmtx = tables.get("hmtx")!;
  const count = (hhea[34]! << 8) | hhea[35]!;
  return Array.from(
    { length: count },
    (_, glyph) => (hmtx[glyph * 4]! << 8) | hmtx[glyph * 4 + 1]!,
  );
}

function checksum(data: Uint8Array): number {
  let sum = 0;
  for (let offset = 0; offset < data.length; offset += 4)
    sum =
      (sum +
        (((data[offset] ?? 0) << 24) |
          ((data[offset + 1] ?? 0) << 16) |
          ((data[offset + 2] ?? 0) << 8) |
          (data[offset + 3] ?? 0))) >>>
      0;
  return sum;
}

describe("reading CFF font programs", () => {
  it("reads glyph names, advance widths, the box and the em", () => {
    const font = readCff(
      cffFont({
        glyphs: [
          { name: "space", width: 250 },
          { name: "A", width: 600 },
          { name: "uni0416", width: 700 },
          { name: "b" },
        ],
        defaultWidth: 480,
        nominalWidth: 300,
      }),
    );
    assert.ok(font);
    assert.equal(font.name, "ABCDEF+TestFace");
    assert.equal(font.cid, false);
    assert.deepEqual(font.glyphNames, [
      ".notdef",
      "space",
      "A",
      "uni0416",
      "b",
    ]);
    // .notdef and "b" have no width of their own: the default one.
    assert.deepEqual(font.widths, [480, 250, 600, 700, 480]);
    assert.deepEqual(font.bbox, [0, -200, 1000, 800]);
    assert.equal(font.unitsPerEm, 1000);
  });

  it("reads a width pushed before a subroutine call", () => {
    const font = readCff(
      cffFont({
        glyphs: [
          { name: "A", width: 640, viaSubroutine: true },
          { name: "B", width: 610 },
        ],
      }),
    );
    assert.deepEqual(font?.widths, [500, 640, 610]);
  });

  it("reads a string many glyphs share once, and no overlong name", () => {
    // 8,000 glyphs named by one 60 KB string: a small file, which must not
    // become 8,000 copies of the string.
    const long = "x".repeat(60_000);
    const shared = Array.from({ length: 8_000 }, () => ({ name: long }));
    const started = performance.now();
    const font = readCff(cffFont({ glyphs: [...shared, { name: "A" }] }))!;
    assert.ok(performance.now() - started < 1_000, "fast");
    assert.equal(font.glyphNames[1], undefined);
    assert.equal(font.glyphNames[8_000], undefined);
    assert.equal(font.glyphNames[8_001], "A");
  });

  it("tells a CID-keyed font apart", () => {
    const font = readCff(cffFont({ glyphs: [{ name: "A" }], cid: true }));
    assert.equal(font?.cid, true);
  });

  it("reads no garbage, and no truncated or other font", () => {
    const program = cffFont({ glyphs: [{ name: "A", width: 600 }] });
    assert.equal(readCff(new Uint8Array([1, 0, 4])), undefined);
    assert.equal(readCff(program.subarray(0, program.length - 10)), undefined);
    assert.equal(readCff(ttf), undefined);
    assert.equal(readCff(new Uint8Array(0)), undefined);
  });

  it("maps glyph names to code points", () => {
    assert.equal(glyphUnicode("A"), 0x41);
    assert.equal(glyphUnicode("space"), 0x20);
    assert.equal(glyphUnicode("quoteright"), 0x2019);
    assert.equal(glyphUnicode("Zcaron"), 0x17d);
    assert.equal(glyphUnicode("uni0416"), 0x416);
    // Adobe Glyph List names beyond the ISOAdobe set.
    assert.equal(glyphUnicode("Euro"), 0x20ac);
    assert.equal(glyphUnicode("cacute"), 0x107);
    assert.equal(glyphUnicode("zdotaccent"), 0x17c);
    assert.equal(glyphUnicode("Scedilla"), 0x15e);
    assert.equal(glyphUnicode("ff"), 0xfb00);
    assert.equal(glyphUnicode("nbspace"), 0xa0);
    assert.equal(glyphUnicode("u1F600"), 0x1f600);
    // Surrogates, unknown names and .notdef stand for nothing.
    assert.equal(glyphUnicode("uniD800"), undefined);
    assert.equal(glyphUnicode("uniFFFF"), undefined);
    assert.equal(glyphUnicode("u1FFFE"), undefined);
    assert.equal(glyphUnicode("uni0416.alt"), undefined);
    assert.equal(glyphUnicode("Acute.sc"), undefined);
    assert.equal(glyphUnicode(".notdef"), undefined);
    assert.equal(glyphUnicode(undefined), undefined);
  });
});

describe("wrapping CFF as OpenType", () => {
  const program = cffFont({
    name: "QWERTY+Brand Face",
    glyphs: [
      { name: "A", width: 600 },
      { name: "B", width: 550 },
      { name: "C", width: 500 },
      { name: "u1F600", width: 1000 },
    ],
  });
  const font = readCff(program)!;
  const unicode = new Map([
    [0x41, 1],
    [0x42, 2],
    [0x43, 3],
    [0x1f600, 4],
  ]);
  const otf = cffOpenType(program, font, unicode)!;

  it("holds the program, its name made safe, in a well-formed OTTO file", () => {
    assert.equal(String.fromCharCode(...otf.subarray(0, 4)), "OTTO");
    const tables = tablesOf(otf);
    assert.deepEqual(
      [...tables.keys()],
      ["CFF ", "OS/2", "cmap", "head", "hhea", "hmtx", "maxp", "name", "post"],
    );
    // Only the space of "Brand Face" changes, which browsers refuse in a CFF name.
    const cff = tables.get("CFF ")!;
    assert.equal(cff.length, program.length);
    const changed = [...cff.keys()].filter((at) => cff[at] !== program[at]);
    assert.deepEqual(
      changed.map((at) => [program[at], cff[at]]),
      [[0x20, 0x5f]],
    );
    assert.equal(readCff(cff)!.name, "QWERTY+Brand_Face");
    // The head adjustment makes the whole file sum to the magic number.
    assert.equal(checksum(otf), 0xb1b0afba);
    assert.equal(otf.length % 4, 0);
  });

  it("lists cmap subtables in platform and encoding order", () => {
    const table = tablesOf(otf).get("cmap")!;
    const view = new DataView(table.buffer, table.byteOffset);
    const records = Array.from(
      { length: view.getUint16(2) },
      (_, index) =>
        `${view.getUint16(4 + index * 8)}/${view.getUint16(6 + index * 8)}`,
    );
    assert.deepEqual(records, ["0/3", "0/4", "3/1", "3/10"]);
  });

  it("gives a font without a box one of the usual proportions", () => {
    const boxless = cffFont({
      glyphs: [{ name: "A", width: 600 }],
      bbox: [0, 0, 0, 0],
    });
    const wrapped = cffOpenType(
      boxless,
      readCff(boxless)!,
      new Map([[0x41, 1]]),
    )!;
    const head = new DataView(
      tablesOf(wrapped).get("head")!.buffer,
      tablesOf(wrapped).get("head")!.byteOffset,
    );
    assert.deepEqual(
      [
        head.getInt16(36),
        head.getInt16(38),
        head.getInt16(40),
        head.getInt16(42),
      ],
      [0, -200, 1000, 800],
    );
    const huge = cffFont({
      glyphs: [{ name: "A", width: 600 }],
      bbox: [-10, -300, 40000, 900],
    });
    const clamped = cffOpenType(huge, readCff(huge)!, new Map([[0x41, 1]]))!;
    const hugeHead = new DataView(
      tablesOf(clamped).get("head")!.buffer,
      tablesOf(clamped).get("head")!.byteOffset,
    );
    assert.equal(hugeHead.getInt16(40), 0x7fff);
  });

  it("maps code points to their glyphs, beyond the BMP as well", () => {
    const cmap = parseCmap(otf);
    assert.equal(hasUnicodeCmap(otf), true);
    assert.equal(cmap.glyph(0x41), 1);
    assert.equal(cmap.glyph(0x43), 3);
    assert.equal(cmap.glyph(0x1f600), 4);
    assert.equal(cmap.glyph(0x44), undefined);
  });

  it("gives every glyph its advance width", () => {
    assert.deepEqual(advances(otf), [500, 600, 550, 500, 1000]);
  });

  it("names the font without its subset tag", () => {
    const name = tablesOf(otf).get("name")!;
    const text = String.fromCharCode(
      ...Array.from(name.subarray(6 + 6 * 12), (byte) => byte).filter(
        (byte) => byte !== 0,
      ),
    );
    assert.match(text, /^BrandFace/);
    assert.doesNotMatch(text, /QWERTY/);
  });

  it("makes no font of glyphs without code points", () => {
    assert.equal(cffOpenType(program, font, new Map()), undefined);
  });
});

describe("making a CFF name safe", () => {
  it("replaces what a browser refuses and keeps the length", () => {
    const program = cffFont({
      name: "AB+Bad (Name)/%",
      glyphs: [{ name: "A" }],
    });
    const safe = cffWithSafeName(program)!;
    assert.equal(safe.length, program.length);
    assert.equal(readCff(safe)!.name, "AB+Bad__Name___");
    assert.equal(readCff(program)!.name, "AB+Bad (Name)/%");
  });

  it("refuses a name longer than a sanitizer takes", () => {
    const program = cffFont({ name: "N".repeat(128), glyphs: [{ name: "A" }] });
    assert.equal(cffWithSafeName(program), undefined);
    assert.equal(cffWithSafeName(new Uint8Array([1, 0, 4])), undefined);
  });
});

describe("completing TrueType programs", () => {
  const stripped = strippedTrueType(ttf);

  it("adds the OS/2, name and post tables and keeps every other one", () => {
    const complete = completeSfnt(stripped, "ABCDEF+Noto-Subset")!;
    const own = tablesOf(stripped);
    const tables = tablesOf(complete);
    assert.deepEqual(
      [...tables.keys()].sort(),
      [...own.keys(), "OS/2", "name", "post"].sort(),
    );
    for (const [tag, table] of own)
      if (tag !== "head") assert.deepEqual(tables.get(tag), table, tag);
    assert.equal(complete.subarray(0, 4).join(), "0,1,0,0");
    const name = tables.get("name")!;
    assert.match(
      String.fromCharCode(...name.filter((byte) => byte !== 0)),
      /Noto-Subset/,
    );
    const os2 = new DataView(
      tables.get("OS/2")!.buffer,
      tables.get("OS/2")!.byteOffset,
    );
    assert.equal(os2.getUint16(0), 4);
    assert.ok(os2.getInt16(2) > 0, "an average width");
  });

  it("lays every table out on a four-byte boundary, with checksums", () => {
    assert.ok(offsetsOf(stripped).some((offset) => offset % 4 !== 0));
    const complete = completeSfnt(stripped, "Noto")!;
    assert.ok(offsetsOf(complete).every((offset) => offset % 4 === 0));
    assert.equal(checksum(complete), 0xb1b0afba);
  });

  it("keeps the tables a font already has", () => {
    const complete = completeSfnt(ttf, "Other")!;
    for (const [tag, table] of tablesOf(ttf))
      if (tag !== "head") assert.deepEqual(tablesOf(complete).get(tag), table);
  });

  it("tells a Unicode cmap from a symbol one", () => {
    assert.equal(hasUnicodeCmap(ttf), true);
    const symbol = ttf.slice();
    const view = new DataView(symbol.buffer);
    const tables = tablesOf(ttf);
    const cmapAt = tables.get("cmap")!.byteOffset;
    for (let index = 0; index < view.getUint16(cmapAt + 2); index += 1) {
      view.setUint16(cmapAt + 4 + index * 8, 3);
      view.setUint16(cmapAt + 6 + index * 8, 0);
    }
    assert.equal(hasUnicodeCmap(symbol), false);
    assert.equal(hasUnicodeCmap(cffFont({ glyphs: [{ name: "A" }] })), false);
  });

  it("refuses a broken directory or a font without its core tables", () => {
    assert.equal(completeSfnt(ttf.subarray(0, 40), "Noto"), undefined);
    const cut = ttf.slice();
    // A table that runs past the end of the file.
    new DataView(cut.buffer).setUint32(12 + 12, 0x7fffffff);
    assert.equal(completeSfnt(cut, "Noto"), undefined);
    assert.equal(
      completeSfnt(cffFont({ glyphs: [{ name: "A" }] }), "CFF"),
      undefined,
    );
  });
});

describe("the face of a PDF text element", () => {
  it("wraps an embedded CFF subset, one face for the font", async () => {
    const engine = await engineFor(cffTextPdf());
    try {
      const first = (await engine.textFont("p0:o0", signal)) as TextFont;
      const second = (await engine.textFont("p0:o1", signal)) as TextFont;
      const element = (await engine.getElement!("p0:o0", signal)) as PdfElement;
      assert.equal(first.elementId, "p0:o0");
      assert.equal(first.missing, undefined);
      assert.equal(first.face?.format, "opentype");
      assert.equal(first.family, element.textStyle?.fontFamily);
      assert.equal(second.key, first.key);
      const data = first.face!.data;
      const cmap = parseCmap(data);
      const widths = advances(data);
      const font = readCff(tablesOf(data).get("CFF ")!)!;
      for (const [character, width] of Object.entries(CFF_TEXT_WIDTHS)) {
        const glyph = cmap.glyph(character.codePointAt(0)!);
        assert.ok(glyph !== undefined, `glyph for ${character}`);
        assert.equal(widths[glyph], width, `width of ${character}`);
        assert.equal(font.widths[glyph], width);
      }
    } finally {
      await engine.dispose();
    }
  });

  it("reports a font the file does not embed by name", async () => {
    const engine = await engineFor(cffTextPdf());
    try {
      const font = (await engine.textFont("p0:o2", signal)) as TextFont;
      assert.equal(font.missing, "not-embedded");
      assert.equal(font.face, undefined);
      assert.equal(font.key, "name:Helvetica");
      assert.equal(font.family, "Helvetica");
    } finally {
      await engine.dispose();
    }
  });

  it("reports a CID-keyed CFF font as missing", async () => {
    const engine = await engineFor(cidTextPdf());
    try {
      const font = (await engine.textFont("p0:o0", signal)) as TextFont;
      assert.equal(font.missing, "cid-keyed");
      assert.equal(font.face, undefined);
    } finally {
      await engine.dispose();
    }
  });

  it("reports a Type 3 font, embedded as drawings, as unreadable", async () => {
    const engine = await engineFor(type3TextPdf());
    try {
      const font = (await engine.textFont("p0:o0", signal)) as TextFont;
      assert.equal(font.missing, "unreadable");
      assert.equal(font.face, undefined);
    } finally {
      await engine.dispose();
    }
  });

  it("hands over an embedded TrueType program with its own tables", async () => {
    const engine = await engineFor(await trueTypeTextPdf(ttf, "Привіт"));
    try {
      const font = (await engine.textFont("p0:o0", signal)) as TextFont;
      assert.equal(font.face?.format, "truetype");
      const cmap = parseCmap(font.face!.data);
      for (const character of "Привіт")
        assert.ok(cmap.drawable(character.codePointAt(0)!), character);
      const own = tablesOf(ttf);
      for (const [tag, table] of tablesOf(font.face!.data))
        if (tag !== "head") assert.deepEqual(table, own.get(tag), tag);
    } finally {
      await engine.dispose();
    }
  });

  it("completes a TrueType subset a browser would refuse", async () => {
    const engine = await engineFor(
      await trueTypeTextPdf(strippedTrueType(ttf), "Привіт"),
    );
    try {
      const font = (await engine.textFont("p0:o0", signal)) as TextFont;
      const tables = tablesOf(font.face!.data);
      for (const tag of ["OS/2", "name", "post"])
        assert.ok(tables.has(tag), tag);
      assert.ok(offsetsOf(font.face!.data).every((offset) => offset % 4 === 0));
    } finally {
      await engine.dispose();
    }
  });

  it("answers nothing for an element that is not text", async () => {
    const engine = await engineFor(cffTextPdf());
    try {
      assert.equal(await engine.textFont("p0:o9", signal), undefined);
      assert.equal(await engine.textFont("missing", signal), undefined);
    } finally {
      await engine.dispose();
    }
  });

  it("answers for a paragraph by its first row, before and after it is edited", async () => {
    const lines = [
      "Since 2013 our independent testing has tracked",
      "software quality across teams and organisations",
      "and shared the results with practitioners.",
    ];
    const { session, end } = await pdfSession(
      await buildPdf([
        {
          texts: lines.map((text, index) => ({
            text,
            x: 72,
            y: 700 - index * 20,
            fontSize: 11,
          })),
        },
      ]),
      { fallbackFont: true },
    );
    try {
      const first = (await session.getElements({ pageIndex: 0 })).items[0]!;
      const paragraph = (await session.getTextParagraph(first.id)).item!;
      assert.match(paragraph.id, /:paragraph$/);
      const row = (await session.getTextFont(first.id)).item!;
      const virtual = (await session.getTextFont(paragraph.id)).item!;
      assert.equal(virtual.elementId, paragraph.id);
      assert.equal(virtual.key, row.key);
      assert.equal(virtual.missing, "not-embedded");
      assert.equal(virtual.family, "Helvetica");

      await session.replaceParagraphText({
        target: paragraph.id,
        text: "Since 2013 our testing has tracked software quality.",
      });
      const edited = (await session.getTextFont(paragraph.id)).item;
      assert.equal(edited?.elementId, paragraph.id);
      assert.equal(edited?.family, "Helvetica");
    } finally {
      await end();
    }
  });

  it("reads through the session, a new text box included", async () => {
    const { session, end } = await pdfSession(cffTextPdf(), {
      fallbackFont: true,
    });
    try {
      const read = await session.getTextFont("p0:o0");
      assert.equal(read.sessionId, session.sessionId);
      assert.equal(read.revision, 0);
      assert.equal(read.item?.face?.format, "opentype");
      const receipt = await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 400, width: 200, height: 40 },
        text: "Новий",
      });
      const box = await session.getTextFont(receipt.createdIds[0]!);
      assert.equal(box.revision, 1);
      assert.equal(box.item?.face?.format, "truetype");
      assert.notEqual(box.item?.key, read.item?.key);
    } finally {
      await end();
    }
  });
});
