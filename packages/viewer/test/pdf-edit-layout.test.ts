import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import { loadPdfEditEngine } from "../src/edit/pdf/provider.js";
import { PdfSession } from "../src/edit/pdf/session.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../src/edit/session.js";
import type {
  PageRect,
  PdfEditSession,
  PdfOperation,
  TextLayout,
} from "../src/index.js";
import { defaultResourceLimits } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";

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

async function pageCountOf(bytes: Uint8Array): Promise<number> {
  const pdfium = await fixturePdfium();
  const document = pdfium.openDocument(bytes);
  try {
    return pdfium.lib.FPDF_GetPageCount(document.handle);
  } finally {
    document.close();
  }
}

/** A PDF session over the loopback worker, with a host that only counts pages. */
async function pdfSession(
  original: Uint8Array,
): Promise<{ session: PdfEditSession; end(): Promise<void> }> {
  const signal = new AbortController().signal;
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      fetchBytes: async () => {
        throw new Error("no fonts");
      },
      decodeImage: async () => {
        throw new Error("no images");
      },
    }),
  );
  const engine = await loadPdfEditEngine(
    original,
    { format: "pdf", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "pdf",
    limits: defaultResourceLimits,
    prepareDocument: async (bytes) => ({ pageCount: await pageCountOf(bytes) }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
  };
  const core = new EditSessionController(
    engine,
    host,
    original,
    await pageCountOf(original),
  );
  return { session: new PdfSession(core), end: () => core.end() };
}
