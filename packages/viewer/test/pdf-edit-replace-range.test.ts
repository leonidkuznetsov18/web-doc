import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { validateSchema } from "../src/edit/schema.js";
import { pdfOperationSchemas } from "../src/edit/pdf/schemas.js";
import type { PdfElement, TextLayout, TextRange } from "../src/index.js";
import { ViewerError } from "../src/index.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * Range-scoped replaceText of ACTION-825, task 35: a text object changes in
 * place when its font covers the new text, is split around the range when
 * only a fallback font can draw the replacement, and a text box reflows.
 */

const range = (elementId: string, start: number, end: number): TextRange => ({
  start: { elementId, offset: start },
  end: { elementId, offset: end },
});

describe("range-scoped replaceText (overlay primitives)", () => {
  it("rewrites a text object in place and keeps the other glyphs put", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        { texts: [{ text: "Hello brave world", fontSize: 20 }] },
      ]),
    );
    try {
      const before = (await session.getTextLayout("p0:o0")).item as TextLayout;
      const receipt = await session.replaceText({
        target: "p0:o0",
        text: "bold",
        range: range("p0:o0", 6, 11),
      });
      assert.deepEqual(receipt.createdIds, []);
      assert.equal(receipt.warnings.length, 0);
      const elements = (await session.getElements({ pageIndex: 0 })).items;
      assert.equal(elements.length, 1, "still one object");
      assert.equal(elements[0]!.text, "Hello bold world");
      const after = (await session.getTextLayout("p0:o0")).item as TextLayout;
      for (let offset = 0; offset < 6; offset += 1)
        assert.deepEqual(
          after.lines[0]!.glyphs[offset]!.box,
          before.lines[0]!.glyphs[offset]!.box,
          `glyph ${offset} stays`,
        );
      // The range a host held moves with the text.
      assert.deepEqual(
        (await session.mapRange(range("p0:o0", 12, 17), 0)).item,
        range("p0:o0", 11, 16),
      );
    } finally {
      await end();
    }
  });

  it("splits the object around text only the fallback font can draw", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        { texts: [{ text: "Hello brave world", fontSize: 20 }] },
      ]),
      { fallbackFont: true },
    );
    try {
      const original = (await session.getTextLayout("p0:o0"))
        .item as TextLayout;
      const receipt = await session.replaceText({
        target: "p0:o0",
        text: "смелый",
        range: range("p0:o0", 6, 11),
      });
      assert.equal(receipt.createdIds.length, 2, "two new parts");
      assert.equal(receipt.warnings[0]?.code, "font-substitution");
      const ids = ["p0:o0", ...receipt.createdIds];
      const parts = await Promise.all(
        ids.map(
          async (id) => (await session.getElement(id)).item as PdfElement,
        ),
      );
      // PDFium's extraction appends a space to an object a gap follows.
      assert.deepEqual(
        parts.map((part) => part.text?.trimEnd()),
        ["Hello", "смелый", " world"],
      );
      const layouts = await Promise.all(
        ids.map(
          async (id) => (await session.getTextLayout(id)).item as TextLayout,
        ),
      );
      const baselines = layouts.map((layout) => layout.lines[0]!.baseline.y);
      for (const y of baselines)
        assert.ok(Math.abs(y - baselines[0]!) < 0.01, `baseline ${y}`);
      // The parts follow each other: each starts where the previous one ends.
      const firstBox = layouts[0]!.lines[0]!.glyphs[0]!.box;
      assert.ok(
        Math.abs(firstBox.x - original.lines[0]!.glyphs[0]!.box.x) < 0.01,
      );
      for (let at = 1; at < layouts.length; at += 1) {
        const previous = layouts[at - 1]!.lines[0]!;
        const lastGlyph = previous.glyphs.at(-1)!;
        const expected = lastGlyph.box.x + lastGlyph.advance;
        const actual = layouts[at]!.lines[0]!.glyphs[0]!.box.x;
        assert.ok(
          Math.abs(actual - expected) < 2,
          `part ${at} starts at ${actual}, after ${expected}`,
        );
      }
      // The tail keeps its font and colour; the middle is in the fallback.
      assert.equal(parts[0]!.textStyle?.fontFamily, "Helvetica");
      assert.equal(parts[2]!.textStyle?.fontFamily, "Helvetica");
      assert.notEqual(parts[1]!.textStyle?.fontFamily, "Helvetica");
      assert.equal(parts[2]!.textStyle?.fontSize, 20);
      // Undo brings the one object back.
      await session.undo();
      assert.equal(
        (await session.getElements({ pageIndex: 0 })).items.length,
        1,
      );
      assert.equal(
        (await session.getElement("p0:o0")).item?.text,
        "Hello brave world",
      );
    } finally {
      await end();
    }
  });

  it("reflows a text box around the replaced range", async () => {
    const { session, end } = await pdfSession(await buildPdf([{}]));
    try {
      const { createdIds } = await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 50, y: 50, width: 120, height: 200 },
        text: "one two three four five",
        style: { fontSize: 14 },
      });
      const [box] = createdIds;
      await session.replaceText({
        target: box!,
        text: "TWO",
        range: range(box!, 4, 7),
      });
      const element = (await session.getElement(box!)).item!;
      assert.equal(element.text, "one TWO three four five");
      const layout = (await session.getTextLayout(box!)).item as TextLayout;
      assert.ok(layout.lines.length >= 2, "still wrapped");
      assert.equal(
        layout.lines.map((line) => line.text).join(" "),
        element.text,
      );
    } finally {
      await end();
    }
  });

  it("refuses ranges outside the text, on other elements and on tables", async () => {
    const { session, end } = await pdfSession(await buildPdf(["Short"]));
    try {
      const { createdIds } = await session.insertTable({
        pageIndex: 0,
        at: { x: 50, y: 300 },
        width: 200,
        rows: [["a", "b"]],
      });
      const bad = async (
        target: string,
        text: string,
        at: TextRange,
      ): Promise<{ code: string; path: string } | undefined> => {
        try {
          await session.replaceText({ target, text, range: at });
          return undefined;
        } catch (error) {
          const issue = (error as ViewerError).details?.issues as
            { code: string; path: string }[] | undefined;
          return issue?.[0] && { code: issue[0].code, path: issue[0].path };
        }
      };
      assert.deepEqual(await bad("p0:o0", "x", range("p0:o0", 2, 9)), {
        code: "invalid-range",
        path: "/range",
      });
      assert.deepEqual(await bad("p0:o0", "x", range("p0:o0", 4, 2)), {
        code: "invalid-range",
        path: "/range",
      });
      assert.deepEqual(await bad("p0:o0", "x", range("p0:o1", 0, 1)), {
        code: "invalid-range",
        path: "/range",
      });
      assert.equal(
        (await bad(createdIds[0]!, "x", range(createdIds[0]!, 0, 1)))?.code,
        "unsupported-target",
      );
      // The schema knows the field and its shape.
      const schema = pdfOperationSchemas.operations.replaceText!;
      assert.deepEqual(
        validateSchema(
          {
            op: "replaceText",
            target: "p0:o0",
            text: "x",
            range: range("p0:o0", 0, 1),
          },
          schema,
          0,
        ),
        [],
      );
      assert.ok(
        validateSchema(
          {
            op: "replaceText",
            target: "p0:o0",
            text: "x",
            range: { start: { elementId: "p0:o0" } },
          },
          schema,
          0,
        ).length > 0,
      );
    } finally {
      await end();
    }
  });
});
