import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import { FontLibrary } from "../src/edit/pdf/engine/fonts.js";
import type { Pdfium } from "../src/edit/pdf/engine/pdfium.js";
import { setText } from "../src/edit/pdf/engine/text-box.js";
import type { PdfElement, PdfOperation } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";

/*
 * A page is written back to its content stream once per batch: PDFium
 * regenerates the whole stream, which takes seconds on pages with thousands
 * of objects, however small the change.
 */

const ttf = new Uint8Array(
  readFileSync(
    new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
  ),
);
const op = <T extends PdfOperation>(operation: T): T => operation;

/** Counts `FPDFPage_GenerateContent` calls while `run` runs. */
function generations(pdfium: Pdfium, run: () => void): number {
  const { lib } = pdfium;
  const generate = lib.FPDFPage_GenerateContent;
  let calls = 0;
  lib.FPDFPage_GenerateContent = (page) => {
    calls += 1;
    return generate.call(lib, page);
  };
  try {
    run();
  } finally {
    lib.FPDFPage_GenerateContent = generate;
  }
  return calls;
}

/** What a reader sees of an element, without the id a batch numbering gives it. */
const shown = (elements: readonly PdfElement[]) =>
  elements.map(({ kind, text, bounds, textStyle, shapeStyle }) => ({
    kind,
    text,
    bounds,
    textStyle,
    shapeStyle,
  }));

describe("PDF page writes", () => {
  let pdfium: Pdfium;
  let original: Uint8Array;

  /** A model whose fallback font is fetched for `operations`. */
  async function model(
    operations: readonly PdfOperation[],
    bytes = original,
  ): Promise<PdfEditDocument> {
    const fonts = new FontLibrary(async () => ttf);
    fonts.setFallbackUrl("https://fonts.test/noto.ttf");
    const document = new PdfEditDocument(pdfium, bytes, fonts);
    await fonts.prepare(document.fontRequests(operations));
    assert.deepEqual(document.validate(operations), []);
    return document;
  }

  before(async () => {
    pdfium = await fixturePdfium();
    original = await buildPdf([
      {
        texts: [
          { text: "Hello world", x: 72, y: 700 },
          { text: "Second line", x: 72, y: 650 },
          { text: "Third line", x: 72, y: 600 },
        ],
        rect: { x: 100, y: 100, width: 50, height: 30, fill: [0, 128, 255] },
      },
      "Another page",
    ]);
  });

  it("reads an object's new text and bounds from the loaded page as after regenerating it", () => {
    const { lib } = pdfium;
    const read = (page: number) => {
      const textPage = lib.FPDFText_LoadPage(page);
      try {
        const object = lib.FPDFPage_GetObject(page, 0);
        let authored = "";
        for (let at = 0; at < lib.FPDFText_CountChars(textPage); at += 1)
          if (
            lib.FPDFText_GetTextObject(textPage, at) === object &&
            lib.FPDFText_IsGenerated(textPage, at) === 0
          )
            authored += String.fromCodePoint(
              lib.FPDFText_GetUnicode(textPage, at),
            );
        return {
          text: pdfium.readWideString((buffer, bytes) =>
            lib.FPDFTextObj_GetText(object, textPage, buffer, bytes),
          ),
          authored,
          bounds: pdfium.readNumbers(4, "float", ([l, b, r, t]) =>
            lib.FPDFPageObj_GetBounds(object, l!, b!, r!, t!),
          ),
        };
      } finally {
        lib.FPDFText_ClosePage(textPage);
      }
    };
    // The second text reads back with one space: PDFium's text page folds
    // the pair whether or not the content was regenerated.
    for (const text of ["Hello there", "Hello  world", "Hello world "]) {
      const document = pdfium.openDocument(original);
      try {
        const page = lib.FPDF_LoadPage(document.handle, 0);
        setText(pdfium, lib.FPDFPage_GetObject(page, 0), text);
        const unwritten = read(page);
        lib.FPDFPage_GenerateContent(page);
        lib.FPDF_ClosePage(page);
        const reloaded = lib.FPDF_LoadPage(document.handle, 0);
        try {
          assert.deepEqual(unwritten, read(reloaded), JSON.stringify(text));
        } finally {
          lib.FPDF_ClosePage(reloaded);
        }
      } finally {
        document.close();
      }
    }
  });

  it("regenerates the page content once for an in-place text edit", async () => {
    const operations = [
      op({
        op: "replaceText",
        target: "p0:o0",
        text: "Big ",
        range: {
          start: { elementId: "p0:o0", offset: 0 },
          end: { elementId: "p0:o0", offset: 0 },
        },
      }),
    ];
    const document = await model(operations);
    try {
      let warnings: readonly unknown[] = [];
      assert.equal(
        generations(pdfium, () => {
          warnings = document.apply(operations).warnings;
        }),
        1,
      );
      assert.deepEqual(warnings, []);
      assert.equal(document.getElement("p0:o0")?.text, "Big Hello world");
      assert.equal(
        await extractPageText(document.materialize("save"), 0),
        "Big Hello world\r\nSecond line\r\nThird line",
      );
    } finally {
      document.dispose();
    }
  });

  it("regenerates the page once when an in-place write reads back wrong and the text is split", async () => {
    // A doubled space reads back as one, so the in-place write is undone and
    // the object split around the new span.
    const operations = [
      op({
        op: "replaceText",
        target: "p0:o0",
        text: "  ",
        range: {
          start: { elementId: "p0:o0", offset: 5 },
          end: { elementId: "p0:o0", offset: 6 },
        },
      }),
    ];
    const document = await model(operations);
    try {
      let codes: string[] = [];
      assert.equal(
        generations(pdfium, () => {
          codes = document.apply(operations).warnings.map((w) => w.code);
        }),
        1,
      );
      assert.deepEqual(codes, ["font-substitution"]);
      assert.deepEqual(
        document
          .getElements({ pageIndex: 0, kinds: ["text"] })
          .map((element) => element.text),
        ["Hello ", " ", "world", "Second line", "Third line"],
      );
    } finally {
      document.dispose();
    }
  });

  it("regenerates the page once for an edit in a substitute font", async () => {
    const operations = [
      op({
        op: "replaceText",
        target: "p0:o1",
        text: "Друга",
        range: {
          start: { elementId: "p0:o1", offset: 0 },
          end: { elementId: "p0:o1", offset: 6 },
        },
      }),
    ];
    const document = await model(operations);
    try {
      let codes: string[] = [];
      assert.equal(
        generations(pdfium, () => {
          codes = document.apply(operations).warnings.map((w) => w.code);
        }),
        1,
      );
      assert.deepEqual(codes, ["font-substitution"]);
      assert.equal(
        await extractPageText(document.materialize("save"), 0),
        "Hello world\r\nДруга line\r\nThird line",
      );
    } finally {
      document.dispose();
    }
  });

  it("regenerates each page a batch changes once", async () => {
    const operations = [
      op({ op: "replaceText", target: "p0:o0", text: "Hello there" }),
      op({
        op: "setTextStyle",
        target: "p0:o1",
        style: { color: "#cc0000", fontSize: 16 },
      }),
      op({ op: "replaceText", target: "p0:o2", text: "Третій" }),
      op({ op: "moveElement", target: "p0:o3", by: { dx: 10, dy: -5 } }),
      op({
        op: "insertShape",
        pageIndex: 0,
        shape: "rectangle",
        rect: { x: 300, y: 300, width: 40, height: 20 },
        fill: { color: "#00aa00" },
      }),
      op({ op: "replaceText", target: "p1:o0", text: "Last page" }),
    ];
    const document = await model(operations);
    try {
      assert.equal(
        generations(pdfium, () => document.apply(operations)),
        2,
      );
      const saved = document.materialize("save");
      assert.equal(
        await extractPageText(saved, 0),
        "Hello there\r\nSecond line\r\nТретій",
      );
      assert.equal(await extractPageText(saved, 1), "Last page");
    } finally {
      document.dispose();
    }
  });

  it("shows and saves a batch as its operations applied one by one", async () => {
    // Each operation reads what the ones before it wrote: the in-place
    // edit's new bounds anchor the resize, the split parts the move.
    const operations = [
      op({ op: "replaceText", target: "p0:o0", text: "Hello wider world" }),
      op({ op: "setTextStyle", target: "p0:o0", style: { fontSize: 18 } }),
      op({
        op: "replaceText",
        target: "p0:o1",
        text: "  ",
        range: {
          start: { elementId: "p0:o1", offset: 6 },
          end: { elementId: "p0:o1", offset: 7 },
        },
      }),
      op({ op: "moveElement", target: "p0:o1", by: { dx: 20, dy: 10 } }),
      op({ op: "moveElement", target: "p0:o3", by: { dx: 10, dy: -5 } }),
    ];
    const batched = await model(operations);
    const stepped = await model(operations);
    try {
      batched.apply(operations);
      for (const operation of operations) stepped.apply([operation]);
      const elements = batched.getElements({ pageIndex: 0 });
      assert.deepEqual(
        shown(elements),
        shown(stepped.getElements({ pageIndex: 0 })),
      );
      const saved = batched.materialize("save");
      assert.equal(
        await extractPageText(saved, 0),
        await extractPageText(stepped.materialize("save"), 0),
      );
      const reopened = new PdfEditDocument(pdfium, saved);
      try {
        assert.deepEqual(
          shown(reopened.getElements({ pageIndex: 0 })),
          shown(elements),
        );
      } finally {
        reopened.dispose();
      }
    } finally {
      batched.dispose();
      stepped.dispose();
    }
  });

  it("writes nothing back for a batch that fails, and restores cleanly after it", async () => {
    const operations = [
      op({ op: "replaceText", target: "p0:o0", text: "Hello there" }),
      // Operation 0 creates nothing, so the reference fails while applying.
      op({ op: "moveElement", target: "$0", by: { dx: 1, dy: 1 } }),
    ];
    const document = await model([operations[0]!]);
    try {
      assert.equal(
        generations(pdfium, () =>
          assert.throws(() => document.apply(operations)),
        ),
        0,
      );
      document.restore([]);
      document.apply([operations[0]!]);
      assert.equal(
        await extractPageText(document.materialize("save"), 0),
        "Hello there\r\nSecond line\r\nThird line",
      );
    } finally {
      document.dispose();
    }
  });
});
