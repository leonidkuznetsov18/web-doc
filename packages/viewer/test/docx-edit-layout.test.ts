import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { TextRun } from "../src/contracts.js";
import { prepareDocxForDisplay } from "../src/adapters/docx-prepass.js";
import { layParagraph, type TextMeasurer } from "../src/edit/docx/layout.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { defaultResourceLimits, type PageRect } from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";
import { docxSession, run, type FakePages } from "./fixtures/docx-session.js";

/*
 * The DOCX overlay reads: `getTextLayout` gives a paragraph's lines on its
 * first page from the renderer's runs (offsets aligned with the paragraph's
 * text, baselines, advance boxes, colour), `positionAt` the caret nearest
 * to a page point. The runs here come from a fake renderer, so glyph
 * advances are the runs' even shares (Node has no canvas to measure with).
 */

const limits = defaultResourceLimits;

async function idsOf(bytes: Uint8Array): Promise<string[]> {
  const display = await prepareDocxForDisplay(bytes, limits);
  const pkg = await OoxmlPackage.open(display.bytes, { limits });
  const xml = new TextDecoder().decode(await pkg.part("/word/document.xml"));
  return [...xml.matchAll(/w:name="_wd([0-9A-F]{8})"/g)].map((m) => m[1]!);
}

/** A run drawn at 16 CSS px in Arial on an 18 px line. */
function line(
  paragraphId: string,
  text: string,
  x: number,
  y: number,
): TextRun {
  return {
    ...run(paragraphId, text, x, y, text.length * 6, 18),
    font: "16px Arial",
    fontFamily: "Arial",
    fontSize: 16,
  };
}

function union(rects: readonly PageRect[]): PageRect {
  const x = Math.min(...rects.map((rect) => rect.x));
  const y = Math.min(...rects.map((rect) => rect.y));
  return {
    x,
    y,
    width: Math.max(...rects.map((rect) => rect.x + rect.width)) - x,
    height: Math.max(...rects.map((rect) => rect.y + rect.height)) - y,
  };
}

function contains(outer: PageRect, inner: PageRect): boolean {
  return (
    inner.x >= outer.x - 1e-9 &&
    inner.y >= outer.y - 1e-9 &&
    inner.x + inner.width <= outer.x + outer.width + 1e-9 &&
    inner.y + inner.height <= outer.y + outer.height + 1e-9
  );
}

const TEXT = "Hello brave new world";
const DOCUMENT = buildDocx({
  theme: { major: "Calibri Light", minor: "Calibri" },
  body:
    paragraph(TEXT) +
    `<w:p><w:r><w:rPr><w:color w:val="4472C4" w:themeColor="accent1"/></w:rPr><w:t>Heading</w:t></w:r></w:p>` +
    paragraph("go go go go go go") +
    sectPr(),
});

describe("DOCX edit session: text layout and caret positions", () => {
  it("lays a two-line paragraph out as two lines a line pitch apart", async () => {
    const [body, heading, long] = await idsOf(DOCUMENT);
    const pages: FakePages = [
      [
        line(body!, "Hello ", 72, 100),
        line(body!, "brave ", 108, 100),
        line(body!, "new world", 72, 118),
        line(heading!, "Heading", 72, 150),
      ],
    ];
    const { session, end } = await docxSession(DOCUMENT, pages, {
      cached: [0],
    });
    try {
      const id = `p:${body}`;
      const layout = (await session.getTextLayout(id)).item;
      assert.ok(layout, "a placed paragraph has a layout");
      assert.equal(layout.elementId, id);
      assert.equal(layout.pageIndex, 0);
      assert.equal(layout.lines.length, 2);
      const [first, second] = layout.lines;
      assert.equal(first!.text, "Hello brave ");
      assert.equal(second!.text, "new world");
      assert.deepEqual(first!.range, {
        start: { elementId: id, offset: 0 },
        end: { elementId: id, offset: 12 },
      });
      assert.deepEqual(second!.range, {
        start: { elementId: id, offset: 12 },
        end: { elementId: id, offset: TEXT.length },
      });
      // Both lines share the line pitch the renderer laid them out with.
      assert.ok(Math.abs(second!.baseline.y - first!.baseline.y - 18) < 1e-9);
      assert.equal(first!.baseline.x, 72);
      // Glyph offsets cover the paragraph text once, in order.
      const offsets = layout.lines.flatMap((entry) =>
        entry.glyphs.map((glyph) => glyph.offset),
      );
      assert.deepEqual(
        offsets,
        [...TEXT].map((_, index) => index),
      );
      // Pen origins step by the advances and start at the run's x.
      assert.deepEqual(
        first!.glyphs.slice(0, 3).map((glyph) => glyph.origin?.x),
        [72, 78, 84],
      );
      assert.equal(first!.glyphs[0]!.advance, 6);
      assert.equal(first!.fontFamily, "Arial");
      assert.equal(first!.fontSize, 16);
      assert.equal(first!.color, "#000000");
      // The advance box runs from the pen origin to the end of the last advance.
      assert.equal(first!.advanceBounds!.x, 72);
      assert.equal(first!.advanceBounds!.width, 72);
      assert.equal(second!.advanceBounds!.width, 54);
      // The baseline lies inside the advance box, at the ascent.
      for (const entry of layout.lines) {
        const box = entry.advanceBounds!;
        assert.ok(entry.baseline.y > box.y);
        assert.ok(entry.baseline.y < box.y + box.height);
      }
      // The frame is the union of the line boxes: two pitches high, as
      // wide as the widest line, holding every line's boxes.
      assert.deepEqual(layout.frame, { x: 72, y: 100, width: 72, height: 36 });
      assert.deepEqual(
        union([layout.frame!, ...layout.lines.map((entry) => entry.bounds)]),
        layout.frame,
      );
      for (const entry of layout.lines)
        assert.ok(contains(layout.frame!, entry.advanceBounds!));

      // A theme colour is reported as the colour it resolves to.
      const headingLayout = (await session.getTextLayout(`p:${heading}`)).item;
      assert.equal(headingLayout?.lines[0]?.color, "#4472C4");

      // Not a paragraph on any page, or not an element at all.
      assert.equal((await session.getTextLayout("p:00000000")).item, undefined);
      assert.equal((await session.getTextLayout(`p:${long}`)).item, undefined);
    } finally {
      await end();
    }
  });

  it("puts the caret at the glyph edge nearest to a point", async () => {
    const [body, heading] = await idsOf(DOCUMENT);
    const pages: FakePages = [
      [
        line(body!, "Hello ", 72, 100),
        line(body!, "brave ", 108, 100),
        line(body!, "new world", 72, 118),
        line(heading!, "Heading", 72, 150),
      ],
      [],
    ];
    const { session, end } = await docxSession(DOCUMENT, pages);
    try {
      const id = `p:${body}`;
      // Inside the second line, left of the middle of "w" (84 to 90).
      const before = (await session.positionAt(0, { x: 86, y: 125 })).item;
      assert.deepEqual(before, { elementId: id, offset: 14 });
      // Past the middle of "w", the caret goes after it.
      const after = (await session.positionAt(0, { x: 88, y: 125 })).item;
      assert.deepEqual(after, { elementId: id, offset: 15 });
      // Right of a line's end, the caret goes to its end.
      const end1 = (await session.positionAt(0, { x: 400, y: 104 })).item;
      assert.deepEqual(end1, { elementId: id, offset: 12 });
      // Between paragraphs, the nearest line wins.
      const near = (await session.positionAt(0, { x: 60, y: 147 })).item;
      assert.deepEqual(near, { elementId: `p:${heading}`, offset: 0 });
      // A page without text has no caret.
      assert.equal(
        (await session.positionAt(1, { x: 80, y: 80 })).item,
        undefined,
      );
    } finally {
      await end();
    }
  });

  it("reports the first page's lines and continues offsets on the next page", async () => {
    const [, , long] = await idsOf(DOCUMENT);
    const pages: FakePages = [
      [line(long!, "go go go ", 72, 700)],
      [line(long!, "go go go", 72, 72)],
    ];
    const { session, end } = await docxSession(DOCUMENT, pages, {
      cached: [1],
    });
    try {
      const id = `p:${long}`;
      const layout = (await session.getTextLayout(id)).item;
      assert.equal(layout?.pageIndex, 0);
      assert.deepEqual(
        layout?.lines.map((entry) => entry.text),
        ["go go go "],
      );
      // The repeated words on page 1 are the paragraph's later ones.
      const caret = (await session.positionAt(1, { x: 73, y: 80 })).item;
      assert.deepEqual(caret, { elementId: id, offset: 9 });
    } finally {
      await end();
    }
  });

  it("measures advances and the baseline from the run's font and ascent box", () => {
    // A canvas stand-in: "W" is four times as wide as "i"; ascent 12, descent 4.
    const measurer: TextMeasurer = {
      advances: (_font, text) => [...text].map((c) => (c === "W" ? 20 : 5)),
      metrics: () => ({ ascent: 12, descent: 4 }),
    };
    const base = { paragraphId: "00000001", textLayer: "docx" as const };
    const lines = layParagraph(
      "p:00000001",
      "Wi hello",
      [
        // List numbering the paragraph's text does not hold.
        {
          pageIndex: 0,
          run: {
            ...base,
            text: "1.",
            x: 50,
            y: 100,
            width: 10,
            height: 24,
            font: "16px Arial",
            fontSize: 16,
          },
        },
        {
          pageIndex: 0,
          run: {
            ...base,
            text: "Wi ",
            x: 72,
            y: 100,
            width: 36,
            height: 24,
            font: "16px Arial",
            fontSize: 16,
            advanceBounds: { x: 72, y: 103, width: 36, height: 16 },
          },
        },
        // Drawn in capitals (w:caps).
        {
          pageIndex: 0,
          run: {
            ...base,
            text: "HELLO",
            x: 108,
            y: 100,
            width: 50,
            height: 24,
            font: "16px Arial",
            fontSize: 16,
          },
        },
      ],
      { fontFamily: "Arial", fontSize: 12 },
      measurer,
    );
    assert.equal(lines.length, 1);
    const [only] = lines;
    // The numbering is left out: the line starts at the text.
    assert.equal(only!.glyphs[0]!.offset, 0);
    assert.equal(only!.baseline.x, 72);
    assert.deepEqual(
      only!.glyphs.map((glyph) => glyph.offset),
      [0, 1, 2, 3, 4, 5, 6, 7],
    );
    // Measured widths scaled to the run's 36 px: 20 : 5 : 5.
    assert.deepEqual(
      only!.glyphs.slice(0, 3).map((glyph) => glyph.advance),
      [24, 6, 6],
    );
    assert.deepEqual(
      only!.glyphs.slice(0, 3).map((glyph) => glyph.origin?.x),
      [72, 96, 102],
    );
    // The baseline divides the renderer's ascent box at the ascent share.
    assert.equal(only!.baseline.y, 103 + 16 * 0.75);
    assert.equal(only!.advanceBounds!.width, 36 + 50);
  });
});
