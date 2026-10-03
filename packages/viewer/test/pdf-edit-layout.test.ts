import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PageRect, PdfOperation, TextLayout } from "../src/index.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";
import { astralTextPdf } from "./fixtures/pdf-fonts.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * The overlay primitives of ACTION-825, task 32: text layout, the caret
 * position nearest to a point, and the rectangles of a range. The engine is
 * driven directly on fixtures PDFium builds; one test goes through the
 * session and the loopback worker to prove the plumbing and the envelope.
 */

const op = <T extends PdfOperation>(operation: T): T => operation;
const ROTATIONS = [0, 1, 2, 3] as const;

async function openModel(
  bytes: Uint8Array,
): Promise<{ model: PdfEditDocument; close(): void }> {
  const pdfium = await fixturePdfium();
  const model = new PdfEditDocument(pdfium, bytes);
  return { model, close: () => model.dispose() };
}

function inside(inner: PageRect, outer: PageRect, slack = 0.01): boolean {
  return (
    inner.x >= outer.x - slack &&
    inner.y >= outer.y - slack &&
    inner.x + inner.width <= outer.x + outer.width + slack &&
    inner.y + inner.height <= outer.y + outer.height + slack
  );
}

function union(rects: readonly PageRect[]): PageRect {
  const x = Math.min(...rects.map((rect) => rect.x));
  const y = Math.min(...rects.map((rect) => rect.y));
  const right = Math.max(...rects.map((rect) => rect.x + rect.width));
  const bottom = Math.max(...rects.map((rect) => rect.y + rect.height));
  return { x, y, width: right - x, height: bottom - y };
}

function nearlyEqual(a: PageRect, b: PageRect, slack = 0.01): boolean {
  return (
    Math.abs(a.x - b.x) <= slack &&
    Math.abs(a.y - b.y) <= slack &&
    Math.abs(a.width - b.width) <= slack &&
    Math.abs(a.height - b.height) <= slack
  );
}

/** Helvetica advance widths (Adobe AFM, 1000 units per em) of the test strings. */
const HELVETICA: Readonly<Record<string, number>> = {
  " ": 278,
  A: 667,
  H: 722,
  R: 722,
  T: 611,
  V: 667,
  W: 944,
  a: 556,
  d: 556,
  e: 556,
  l: 222,
  o: 556,
  r: 333,
  v: 500,
  w: 722,
};

/** The pen distance of `text` set in Helvetica at `fontSize`, without spacing. */
function helveticaAdvance(text: string, fontSize: number): number {
  let units = 0;
  for (const character of text) units += HELVETICA[character]!;
  return (units * fontSize) / 1000;
}

/** Ascent and descent (both positive) PDFium reports for an object's font. */
async function fontMetrics(
  bytes: Uint8Array,
  objectIndex: number,
  fontSize: number,
): Promise<{ ascent: number; descent: number }> {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.openDocument(bytes);
  try {
    const page = lib.FPDF_LoadPage(document.handle, 0);
    try {
      const font = lib.FPDFTextObj_GetFont(
        lib.FPDFPage_GetObject(page, objectIndex),
      );
      const ascent = pdfium.readNumbers(1, "float", ([pointer]) =>
        lib.FPDFFont_GetAscent(font, fontSize, pointer!),
      )![0]!;
      const descent = pdfium.readNumbers(1, "float", ([pointer]) =>
        lib.FPDFFont_GetDescent(font, fontSize, pointer!),
      )![0]!;
      return { ascent, descent: Math.abs(descent) };
    } finally {
      lib.FPDF_ClosePage(page);
    }
  } finally {
    document.close();
  }
}

/** A one-page PDF whose content stream is `content`, with Helvetica as /F1. */
function rawPdf(content: string, toUnicode?: string): Uint8Array {
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
    `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica ${toUnicode ? "/ToUnicode 6 0 R" : ""} >>`,
    `<< /Length ${Buffer.byteLength(content)} >>\nstream\n${content}\nendstream`,
    ...(toUnicode
      ? [
          `<< /Length ${Buffer.byteLength(toUnicode)} >>\nstream\n${toUnicode}\nendstream`,
        ]
      : []),
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

/** A point `fraction` of the way through a glyph box in reading direction. */
function along(
  box: PageRect,
  rotation: number,
  fraction: number,
): { x: number; y: number } {
  const centre = { x: box.x + box.width / 2, y: box.y + box.height / 2 };
  switch (rotation) {
    case 1:
      return { x: centre.x, y: box.y + box.height * fraction };
    case 2:
      return { x: box.x + box.width * (1 - fraction), y: centre.y };
    case 3:
      return { x: centre.x, y: box.y + box.height * (1 - fraction) };
    default:
      return { x: box.x + box.width * fraction, y: centre.y };
  }
}

describe("text layout (overlay primitives)", () => {
  for (const rotation of ROTATIONS)
    it(`lays out a text object on a page rotated ${rotation * 90}°`, async () => {
      const { model, close } = await openModel(
        await buildPdf([
          {
            texts: [{ text: "Hello world", x: 72, y: 700, fontSize: 24 }],
            rotation,
          },
        ]),
      );
      try {
        const element = model.getElement("p0:o0")!;
        const layout = model.textLayout("p0:o0")!;
        assert.equal(layout.elementId, "p0:o0");
        assert.equal(layout.pageIndex, 0);
        assert.equal(layout.lines.length, 1);
        const [line] = layout.lines;
        assert.equal(line!.text, "Hello world");
        assert.deepEqual(line!.range, {
          start: { elementId: "p0:o0", offset: 0 },
          end: { elementId: "p0:o0", offset: 11 },
        });
        assert.equal(line!.glyphs.length, 11);
        assert.equal(line!.fontFamily, "Helvetica");
        assert.equal(line!.fontSize, 24);
        assert.equal(line!.color, "#000000");
        for (const glyph of line!.glyphs) {
          assert.ok(
            inside(glyph.box, element.bounds),
            `glyph ${glyph.offset} inside the element: ${JSON.stringify(glyph.box)} in ${JSON.stringify(element.bounds)}`,
          );
          assert.ok(glyph.advance > 0, "advance");
        }
        assert.ok(
          nearlyEqual(line!.bounds, union(line!.glyphs.map((g) => g.box))),
          "line bounds are the union of the glyph boxes",
        );
        // The baseline starts at the first glyph's pen position: a side
        // bearing left of the ink, and within the ink's vertical extent.
        assert.ok(
          inside(
            { ...line!.baseline, width: 0, height: 0 },
            element.bounds,
            line!.fontSize / 4,
          ),
          `baseline ${JSON.stringify(line!.baseline)} near ${JSON.stringify(element.bounds)}`,
        );
        // The advance box turns with the page like the ink: it holds the
        // ink and runs the pen distance along the reading direction.
        const advance = line!.advanceBounds!;
        assert.ok(
          inside(line!.bounds, advance),
          `ink ${JSON.stringify(line!.bounds)} inside ${JSON.stringify(advance)}`,
        );
        const length = rotation % 2 === 0 ? advance.width : advance.height;
        assert.ok(
          Math.abs(length - helveticaAdvance("Hello world", 24)) < 0.01,
          `advance length ${length}`,
        );
        assert.deepEqual(line!.glyphs[0]!.origin, line!.baseline);
        assert.deepEqual(layout.frame, advance);
      } finally {
        close();
      }
    });

  it("honours a crop box", async () => {
    const { model, close } = await openModel(
      await buildPdf([
        {
          texts: [{ text: "Cropped", x: 100, y: 120, fontSize: 18 }],
          cropBox: [50, 50, 400, 500],
        },
      ]),
    );
    try {
      const element = model.getElement("p0:o0")!;
      const [line] = model.textLayout("p0:o0")!.lines;
      // Page space starts at the crop box: the glyphs sit 50 pt in from its left edge.
      assert.ok(
        line!.bounds.x >= 49 && line!.bounds.x <= 52,
        `x ${line!.bounds.x}`,
      );
      for (const glyph of line!.glyphs)
        assert.ok(inside(glyph.box, element.bounds), "glyph inside bounds");
    } finally {
      close();
    }
  });

  it("reports a line per drawn line of a text box and per cell of a table", async () => {
    const { model, close } = await openModel(await buildPdf([{}]));
    try {
      const change = model.apply([
        op({
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 50, y: 50, width: 120, height: 200 },
          text: "one two three four five six seven eight nine ten",
          style: { fontSize: 14 },
        }),
        op({
          op: "insertTable",
          pageIndex: 0,
          at: { x: 50, y: 300 },
          width: 300,
          rows: [
            ["a1", "b1"],
            ["a2", "b2"],
          ],
        }),
      ]);
      const [boxId, tableId] = change.createdIds;
      const box = model.textLayout(boxId!)!;
      assert.ok(
        box.lines.length >= 3,
        `wrapped into ${box.lines.length} lines`,
      );
      const boxText = model.getElement(boxId!)!.text!;
      let previous = -1;
      for (const line of box.lines) {
        assert.ok(line.range.start.offset > previous, "lines in reading order");
        assert.equal(
          boxText.slice(line.range.start.offset, line.range.end.offset),
          line.text,
          "a line's range covers its text",
        );
        previous = line.range.start.offset;
      }
      const table = model.textLayout(tableId!)!;
      assert.deepEqual(
        table.lines.map((line) => line.text),
        ["a1", "b1", "a2", "b2"],
      );
    } finally {
      close();
    }
  });

  for (const rotation of ROTATIONS)
    it(`resolves positions and range rectangles on a page rotated ${rotation * 90}°`, async () => {
      const { model, close } = await openModel(
        await buildPdf([
          {
            texts: [{ text: "Hello world", x: 72, y: 700, fontSize: 24 }],
            rotation,
          },
        ]),
      );
      try {
        const [line] = model.textLayout("p0:o0")!.lines;
        const glyph = line!.glyphs[3]!;
        assert.deepEqual(
          model.positionAt(0, along(glyph.box, rotation, 0.25)),
          {
            elementId: "p0:o0",
            offset: 3,
          },
        );
        assert.deepEqual(
          model.positionAt(0, along(glyph.box, rotation, 0.75)),
          {
            elementId: "p0:o0",
            offset: 4,
          },
        );
        // Far from any glyph the nearest one still answers.
        const far = model.positionAt(0, { x: 5, y: 5 });
        assert.equal(far?.elementId, "p0:o0");
        const range = {
          start: { elementId: "p0:o0", offset: 2 },
          end: { elementId: "p0:o0", offset: 6 },
        };
        const rects = model.rangeRects(range);
        assert.equal(rects.length, 1);
        for (const offset of [2, 3, 4, 5]) {
          const point = along(line!.glyphs[offset]!.box, rotation, 0.25);
          assert.ok(
            inside({ ...point, width: 0, height: 0 }, rects[0]!),
            `the rect contains glyph ${offset}`,
          );
          const resolved = model.positionAt(0, point)!;
          assert.ok(
            resolved.offset >= 2 && resolved.offset < 6,
            `glyph ${offset} resolves into the range (${resolved.offset})`,
          );
        }
      } finally {
        close();
      }
    });

  it("answers nothing for unknown, non-text and empty targets", async () => {
    const { model, close } = await openModel(
      await buildPdf([
        { rect: { x: 10, y: 10, width: 50, height: 50, fill: [255, 0, 0] } },
        { text: "Second" },
      ]),
    );
    try {
      assert.equal(model.textLayout("nope"), undefined);
      assert.equal(
        model.textLayout("p0:o0"),
        undefined,
        "a shape has no layout",
      );
      assert.equal(
        model.positionAt(0, { x: 20, y: 20 }),
        undefined,
        "no text on the page",
      );
      assert.equal(model.positionAt(9, { x: 20, y: 20 }), undefined);
      assert.deepEqual(
        model.rangeRects({
          start: { elementId: "p0:o0", offset: 0 },
          end: { elementId: "p1:o0", offset: 2 },
        }),
        [],
        "a range across pages",
      );
      assert.deepEqual(
        model.rangeRects({
          start: { elementId: "p1:o0", offset: 4 },
          end: { elementId: "p1:o0", offset: 4 },
        }),
        [],
        "an empty range",
      );
    } finally {
      close();
    }
  });

  it("serves the primitives through the session with a stamped envelope", async () => {
    const original = await buildPdf(["Session text"]);
    const { session, end } = await pdfSession(original);
    try {
      const layout = await session.getTextLayout("p0:o0");
      assert.equal(layout.sessionId, session.sessionId);
      assert.equal(layout.revision, 0);
      const lines = (layout.item as TextLayout).lines;
      assert.equal(lines[0]!.text, "Session text");
      const position = await session.positionAt(0, lines[0]!.glyphs[0]!.box);
      assert.deepEqual(position.item, { elementId: "p0:o0", offset: 0 });
      const rects = await session.rangeRects(lines[0]!.range);
      assert.equal(rects.items.length, 1);
      await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 36, y: 36, width: 200, height: 40 },
        text: "later",
      });
      assert.equal((await session.getTextLayout("p0:o0")).revision, 1);
      assert.equal((await session.getTextLayout("missing")).item, undefined);
    } finally {
      await end();
    }
  });
});

describe("text field frame (ACTION-922)", () => {
  it("gives a line the advance box a text field needs, wider than its ink by the side bearings", async () => {
    const bytes = await buildPdf([
      { texts: [{ text: "Wave AVATAR", x: 72, y: 700, fontSize: 24 }] },
    ]);
    const { ascent, descent } = await fontMetrics(bytes, 0, 24);
    const { model, close } = await openModel(bytes);
    try {
      const element = model.getElement("p0:o0")!;
      const [line] = model.textLayout("p0:o0")!.lines;
      const advance = line!.advanceBounds!;
      assert.ok(advance, "the line has an advance box");
      // From the first pen position to the last glyph's advance end: the
      // sum of the font's advances, which a browser lays the text out by.
      assert.ok(
        Math.abs(advance.x - line!.baseline.x) < 0.01,
        `starts at the pen origin: ${advance.x} vs ${line!.baseline.x}`,
      );
      assert.ok(
        Math.abs(advance.width - helveticaAdvance("Wave AVATAR", 24)) < 0.01,
        `advance width ${advance.width}, AFM ${helveticaAdvance("Wave AVATAR", 24)}`,
      );
      assert.ok(
        advance.width > line!.bounds.width &&
          advance.width > element.bounds.width,
        `advance ${advance.width} wider than ink ${line!.bounds.width}`,
      );
      assert.ok(
        element.bounds.x > advance.x,
        "the ink starts right of the pen",
      );
      // Vertically the font's ascent above the baseline to its descent below.
      assert.ok(
        Math.abs(advance.y - (line!.baseline.y - ascent)) < 0.01,
        `top ${advance.y}, baseline ${line!.baseline.y}, ascent ${ascent}`,
      );
      assert.ok(
        Math.abs(advance.height - (ascent + descent)) < 0.01,
        `height ${advance.height}, ascent + descent ${ascent + descent}`,
      );
      assert.ok(inside(line!.bounds, advance), "the advance box holds the ink");
      // `bounds` stays the ink: hit tests and paragraph checks compare it.
      assert.ok(
        nearlyEqual(element.bounds, line!.bounds),
        "element bounds are still the ink",
      );
    } finally {
      close();
    }
  });

  it("places each glyph at its drawn pen position, character spacing and kerning included", async () => {
    const bytes = rawPdf(
      "BT /F1 20 Tf 2 Tc 1 0 0 1 72 700 Tm [(Wa) 100 (ve)] TJ ET",
    );
    const { model, close } = await openModel(bytes);
    try {
      const [line] = model.textLayout("p0:o0")!.lines;
      assert.equal(line!.text, "Wave");
      const origins = line!.glyphs.map((glyph) => glyph.origin!);
      assert.equal(origins.length, 4);
      assert.deepEqual(origins[0], line!.baseline);
      // Each step is the glyph's width, plus 2 pt of Tc, less the 100/1000 em
      // the TJ array moves "v" back.
      const steps = [
        helveticaAdvance("W", 20) + 2,
        helveticaAdvance("a", 20) + 2 - 2,
        helveticaAdvance("v", 20) + 2,
      ];
      for (const [index, step] of steps.entries()) {
        const drawn = origins[index + 1]!.x - origins[index]!.x;
        assert.ok(
          Math.abs(drawn - step) < 0.01,
          `step ${index}: ${drawn} vs ${step}`,
        );
        assert.equal(origins[index + 1]!.y, origins[index]!.y);
      }
      const advance = line!.advanceBounds!;
      const last = origins[3]!.x + helveticaAdvance("e", 20) - origins[0]!.x;
      assert.ok(
        Math.abs(advance.width - last) < 0.01,
        `advance width ${advance.width} vs ${last}`,
      );
    } finally {
      close();
    }
  });

  it("frames plain text by its advance box and a table by its cells' advance boxes", async () => {
    const { model, close } = await openModel(
      await buildPdf([
        { texts: [{ text: "Hello world", x: 72, y: 700, fontSize: 24 }] },
      ]),
    );
    try {
      const text = model.textLayout("p0:o0")!;
      assert.deepEqual(text.frame, text.lines[0]!.advanceBounds);
      const [tableId] = model.apply([
        op({
          op: "insertTable",
          pageIndex: 0,
          at: { x: 50, y: 300 },
          width: 300,
          rows: [
            ["a1", "b1"],
            ["a2", "b2"],
          ],
        }),
      ]).createdIds;
      const table = model.textLayout(tableId!)!;
      assert.equal(table.lines.length, 4);
      assert.ok(
        nearlyEqual(
          table.frame!,
          union(table.lines.map((line) => line.advanceBounds!)),
        ),
        `frame ${JSON.stringify(table.frame)}`,
      );
      // A page layout reports the same geometry as the element's own read.
      const page = model.pageLayout(0)!;
      for (const layout of [text, table])
        assert.deepEqual(
          page.layouts.find((entry) => entry.elementId === layout.elementId),
          layout,
        );
    } finally {
      close();
    }
  });

  it("frames a text box by its own rect", async () => {
    const { model, close } = await openModel(await buildPdf([{}]));
    try {
      const rect = { x: 50, y: 60, width: 200, height: 120 };
      const [boxId] = model.apply([
        op({
          op: "insertTextBox",
          pageIndex: 0,
          rect,
          text: "one two three four five six seven eight nine ten",
          style: { fontSize: 14 },
        }),
      ]).createdIds;
      const layout = model.textLayout(boxId!)!;
      assert.ok(layout.lines.length >= 2, `${layout.lines.length} lines`);
      assert.deepEqual(layout.frame, rect);
      for (const line of layout.lines) {
        // Left-aligned lines start at the box's edge and fit its width.
        const advance = line.advanceBounds!;
        assert.ok(Math.abs(advance.x - rect.x) < 0.01, `x ${advance.x}`);
        assert.ok(advance.x + advance.width <= rect.x + rect.width + 0.01);
      }
      assert.deepEqual(
        model.pageLayout(0)!.layouts.find((entry) => entry.elementId === boxId)
          ?.frame,
        rect,
      );
      model.apply([
        op({ op: "moveElement", target: boxId!, by: { dx: 10, dy: 20 } }),
      ]);
      assert.deepEqual(model.textLayout(boxId!)!.frame, {
        ...rect,
        x: 60,
        y: 80,
      });
    } finally {
      close();
    }
  });

  it("frames a paragraph by its own rect, before and after it is rewritten", async () => {
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
            x: 90,
            y: 700 - index * 20,
            fontSize: 11,
          })),
        },
      ]),
    );
    try {
      const row = (await session.getElements({ pageIndex: 0 })).items[0]!;
      const paragraph = (await session.getTextParagraph(row.id)).item!;
      assert.ok(paragraph, "the rows form a paragraph");
      const layout = (await session.getTextLayout(paragraph.id)).item!;
      assert.equal(layout.lines.length, 3);
      assert.deepEqual(layout.frame, paragraph.bounds);
      // The paragraph's rect is advance-based already: it starts at the pen
      // and is as wide as its widest line's advance box.
      const advances = layout.lines.map((line) => line.advanceBounds!);
      assert.ok(Math.abs(layout.frame!.x - advances[0]!.x) < 0.01);
      assert.ok(
        Math.abs(
          layout.frame!.width -
            Math.max(...advances.map((advance) => advance.width)),
        ) < 0.01,
        `frame ${layout.frame!.width}`,
      );
      // A row read on its own is plain text, framed by its own advance box.
      const own = (await session.getTextLayout(row.id)).item!;
      assert.deepEqual(own.frame, own.lines[0]!.advanceBounds);
      await session.replaceParagraphText({
        target: paragraph.id,
        text: "A shorter paragraph that still wraps across two of its lines.",
      });
      const element = (await session.getElement(paragraph.id)).item!;
      const rewritten = (await session.getTextLayout(paragraph.id)).item!;
      assert.deepEqual(rewritten.frame, element.bounds);
      // The rewrite wraps inside the paragraph's own width and left edge.
      assert.equal(rewritten.frame!.x, paragraph.bounds.x);
      assert.equal(rewritten.frame!.width, paragraph.bounds.width);
    } finally {
      await end();
    }
  });

  it("frames a line turned 90° on the page as an upright box turned with it", async () => {
    const bytes = await buildPdf([
      {
        texts: [
          { text: "Wave AVATAR", x: 72, y: 700, fontSize: 24 },
          // Reads down the page: the text turned 90° clockwise.
          {
            text: "Wave AVATAR",
            x: 300,
            y: 400,
            fontSize: 1,
            matrix: [0, -24, 24, 0],
          },
        ],
      },
    ]);
    const { ascent, descent } = await fontMetrics(bytes, 1, 24);
    const { model, close } = await openModel(bytes);
    try {
      const element = model.getElement("p0:o1")!;
      assert.equal(element.rotation, 90);
      const upright = model.textLayout("p0:o0")!.lines[0]!.advanceBounds!;
      const layout = model.textLayout("p0:o1")!;
      const [line] = layout.lines;
      const advance = line!.advanceBounds!;
      // Axis-aligned in page space like `bounds`: the advance runs down the
      // page from the pen origin, ascent to the right, descent to the left.
      assert.ok(
        Math.abs(advance.y - line!.baseline.y) < 0.01,
        `y ${advance.y}`,
      );
      assert.ok(
        Math.abs(advance.height - helveticaAdvance("Wave AVATAR", 24)) < 0.01,
        `height ${advance.height}`,
      );
      assert.ok(
        Math.abs(advance.x - (line!.baseline.x - descent)) < 0.01,
        `x ${advance.x}`,
      );
      assert.ok(
        Math.abs(advance.width - (ascent + descent)) < 0.01,
        `width ${advance.width}`,
      );
      assert.ok(inside(line!.bounds, advance), "the advance box holds the ink");
      assert.ok(inside(element.bounds, advance), "and the element's ink");
      // Turned back by the element's rotation it is the upright line's box.
      assert.ok(Math.abs(advance.width - upright.height) < 0.01);
      assert.ok(Math.abs(advance.height - upright.width) < 0.01);
      assert.deepEqual(layout.frame, advance);
      const origins = line!.glyphs.map((glyph) => glyph.origin!);
      for (const [index, origin] of origins.slice(1).entries()) {
        assert.equal(origin.x, origins[index]!.x);
        assert.ok(origin.y > origins[index]!.y, "the pen moves down the page");
      }
    } finally {
      close();
    }
  });
});

describe("PDF UTF-16 positions", () => {
  it("keeps source spans when case folding expands a BMP character", async () => {
    const cmap = `/CIDInit /ProcSet findresource begin
12 dict begin begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /CaseUnicode def /CMapType 2 def
1 begincodespacerange <00> <FF> endcodespacerange
1 beginbfchar <49> <0130> endbfchar
endcmap CMapName currentdict /CMap defineresource pop end end`;
    const { session, end } = await pdfSession(
      rawPdf("BT /F1 24 Tf 72 700 Td (IB) Tj ET", cmap),
    );
    try {
      assert.equal((await session.getElement("p0:o0")).item?.text, "İB");
      for (const [query, text, from, to] of [
        ["i", "İ", 0, 1],
        ["B", "B", 1, 2],
        ["İB", "İB", 0, 2],
      ] as const) {
        const [match] = (await session.findText(query)).items;
        assert.equal(match?.text, text, query);
        assert.deepEqual(match?.ranges, [
          {
            start: { elementId: "p0:o0", offset: from },
            end: { elementId: "p0:o0", offset: to },
          },
        ]);
      }
      assert.deepEqual(
        (await session.findText("i", { caseSensitive: true })).items,
        [],
      );
    } finally {
      await end();
    }
  });
  for (const encoding of ["glyph-names", "to-unicode"] as const) {
    it(`keeps one astral glyph and complete caret boundaries with ${encoding}`, async () => {
      const { session, end } = await pdfSession(astralTextPdf(encoding));
      try {
        const element = (await session.getElement("p0:o0")).item;
        assert.equal(element?.text, "A😀B");
        const layout = (await session.getTextLayout("p0:o0")).item;
        assert.ok(layout);
        const [line] = layout.lines;
        assert.ok(line);
        assert.equal(line.text, "A😀B");
        assert.equal(
          line.range.end.offset,
          4,
          "line ranges count UTF-16 code units",
        );
        assert.deepEqual(
          line.glyphs.map((glyph) => glyph.offset),
          [0, 1, 3],
        );
        assert.ok(
          Math.abs(
            line.glyphs.reduce((sum, glyph) => sum + glyph.advance, 0) - 51.6,
          ) < 0.01,
          "one physical advance per glyph",
        );
        const emoji = line.glyphs[1];
        const following = line.glyphs[2];
        assert.ok(emoji);
        assert.ok(following);
        for (const [glyph, before, after] of [
          [emoji, 1, 3],
          [following, 3, 4],
        ] as const) {
          assert.equal(
            (await session.positionAt(0, along(glyph.box, 0, 0.25))).item
              ?.offset,
            before,
          );
          assert.equal(
            (await session.positionAt(0, along(glyph.box, 0, 0.75))).item
              ?.offset,
            after,
          );
        }
        const range = (start: number, end: number) => ({
          start: { elementId: "p0:o0", offset: start },
          end: { elementId: "p0:o0", offset: end },
        });
        assert.deepEqual((await session.rangeRects(range(1, 3))).items, [
          emoji.box,
        ]);
        assert.deepEqual((await session.rangeRects(range(3, 4))).items, [
          following.box,
        ]);
      } finally {
        await end();
      }
    });

    it(`returns the exact UTF-16 search target after an astral glyph with ${encoding}`, async () => {
      const { session, end } = await pdfSession(astralTextPdf(encoding), {
        fallbackFont: true,
      });
      try {
        for (const [query, start, finish] of [
          ["😀", 1, 3],
          ["B", 3, 4],
          ["😀B", 1, 4],
        ] as const) {
          const result = await session.findText(query);
          assert.equal(result.items.length, 1, `one match for ${query}`);
          const [match] = result.items;
          assert.ok(match);
          assert.equal(match.text, query);
          assert.deepEqual(match.ranges, [
            {
              start: { elementId: "p0:o0", offset: start },
              end: { elementId: "p0:o0", offset: finish },
            },
          ]);
        }
        assert.deepEqual(
          (await session.findText("\ud83d")).items,
          [],
          "no half-surrogate match",
        );
        const [match] = (await session.findText("B")).items;
        const range = match?.ranges[0];
        assert.ok(range);
        await session.replaceText({ target: "p0:o0", text: "A", range });
        assert.equal(
          (await session.getElements({ pageIndex: 0 })).items
            .map((element) => element.text ?? "")
            .join(""),
          "A😀A",
        );
        const reopened = await pdfSession((await session.save()).bytes);
        try {
          assert.equal(
            (await reopened.session.getElements({ pageIndex: 0 })).items
              .map((element) => element.text ?? "")
              .join(""),
            "A😀A",
          );
          assert.deepEqual((await reopened.session.findText("B")).items, []);
          const [saved] = (await reopened.session.findText("😀")).items;
          assert.equal(saved?.text, "😀");
          assert.deepEqual(saved?.ranges, [
            {
              start: { elementId: "p0:o0", offset: 1 },
              end: { elementId: "p0:o0", offset: 3 },
            },
          ]);
          const layout = (await reopened.session.getTextLayout("p0:o0")).item;
          assert.deepEqual(
            layout?.lines[0]?.glyphs.map((glyph) => glyph.offset),
            [0, 1],
          );
          assert.equal(layout?.lines[0]?.range.end.offset, 3);
        } finally {
          await reopened.end();
        }
        await session.undo();
        assert.equal((await session.getElement("p0:o0")).item?.text, "A😀B");
        assert.deepEqual((await session.findText("B")).items[0]?.ranges, [
          range,
        ]);
      } finally {
        await end();
      }
    });
  }
});
