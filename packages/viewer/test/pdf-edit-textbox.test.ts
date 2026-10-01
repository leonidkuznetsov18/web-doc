import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import { layoutText } from "../src/edit/pdf/engine/text-layout.js";
import { pdfOperationSchemas } from "../src/edit/pdf/schemas.js";
import { checkOperations } from "../src/edit/operations.js";
import { assertSupportedSchema } from "../src/edit/schema.js";
import type { InsertTextBoxOperation, PageRect } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";

/** Six points per character at any size: widths are easy to predict. */
const monospace = (text: string) => text.length * 6;

function layout(
  text: string,
  width: number,
  height = 1000,
  align = "left" as const,
) {
  return layoutText({
    text,
    width,
    height,
    fontSize: 12,
    lineHeight: 1.2,
    align,
    ascent: 10,
    advance: monospace,
  });
}

function textBox(
  fields: Partial<InsertTextBoxOperation> & { readonly text: string },
): InsertTextBoxOperation {
  return {
    op: "insertTextBox",
    pageIndex: 0,
    rect: { x: 72, y: 72, width: 200, height: 100 },
    ...fields,
  };
}

async function objectCount(bytes: Uint8Array, pageIndex: number) {
  const pdfium = await fixturePdfium();
  const document = pdfium.openDocument(bytes);
  try {
    const page = pdfium.lib.FPDF_LoadPage(document.handle, pageIndex);
    try {
      return pdfium.lib.FPDFPage_CountObjects(page);
    } finally {
      pdfium.lib.FPDF_ClosePage(page);
    }
  } finally {
    document.close();
  }
}

function within(rect: PageRect, outer: PageRect, slack: number): boolean {
  return (
    rect.x >= outer.x - slack &&
    rect.y >= outer.y - slack &&
    rect.x + rect.width <= outer.x + outer.width + slack &&
    rect.y + rect.height <= outer.y + outer.height + slack
  );
}

describe("text layout", () => {
  it("wraps greedily on spaces and keeps paragraphs", () => {
    const result = layout("aaa bbb ccc\n\nddd", 60);
    assert.deepEqual(
      result.lines.map((line) => line.text),
      ["aaa bbb", "ccc", "", "ddd"],
    );
    assert.deepEqual(
      result.lines.map((line) => Math.round(line.baseline * 1000) / 1000),
      [10, 24.4, 38.8, 53.2],
    );
    assert.equal(Math.round(result.height * 1000) / 1000, 57.6);
    assert.equal(result.overflow, false);
  });

  it("breaks words wider than the box and aligns lines", () => {
    const broken = layout("abcdefghij xy", 30);
    assert.deepEqual(
      broken.lines.map((line) => line.text),
      ["abcde", "fghij", "xy"],
    );
    const right = layout("ab\nabcd", 60, 1000, "right" as never);
    assert.deepEqual(
      right.lines.map((line) => line.x),
      [48, 36],
    );
    const center = layoutText({
      text: "ab",
      width: 60,
      height: 10,
      fontSize: 12,
      lineHeight: 1.2,
      align: "center",
      ascent: 10,
      advance: monospace,
    });
    assert.equal(center.lines[0]?.x, 24);
    assert.equal(center.overflow, true);
    assert.equal(layout("\t x", 100).lines[0]?.text, "  x");
  });
});

describe("insertTextBox", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;
  let original: Uint8Array;

  before(async () => {
    pdfium = await fixturePdfium();
    original = await buildPdf([
      "Existing",
      { width: 300, height: 400, rotation: 1 },
    ]);
    for (const schema of Object.values(pdfOperationSchemas.operations))
      assertSupportedSchema(schema);
  });

  it("lays text out inside the box, saves it and lists it as one element", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const operation = textBox({
        text: "Hello wrapped text that is long enough to wrap onto several lines",
      });
      assert.deepEqual(checkOperations([operation], pdfOperationSchemas), []);
      assert.deepEqual(model.validate([operation]), []);
      const change = model.apply([operation]);
      assert.deepEqual(change, {
        createdIds: ["p0:n1.0.0"],
        changedPages: [0],
        pageCount: 2,
        warnings: [],
      });
      const elements = model.getElements({ pageIndex: 0 });
      assert.deepEqual(
        elements.map((element) => [element.kind, element.id]),
        [
          ["text", "p0:o0"],
          ["textBox", "p0:n1.0.0"],
        ],
      );
      const box = elements[1]!;
      assert.equal(box.text, operation.text);
      assert.equal(
        within(box.bounds, operation.rect, 1),
        true,
        JSON.stringify(box.bounds),
      );
      assert.ok(box.bounds.height > 20, "two lines of 12 pt text");
      assert.deepEqual(box.textStyle, {
        fontFamily: "Helvetica",
        fontSize: 12,
        bold: false,
        italic: false,
        color: "#000000",
      });
      assert.ok(box.operations.includes("replaceText"));

      const saved = model.materialize();
      assert.equal(saved.length > original.length, true);
      assert.equal(
        original.every((byte, index) => saved[index] === byte),
        true,
      );
      const text = await extractPageText(saved, 0);
      for (const word of ["Hello", "wrapped", "several", "lines"])
        assert.ok(text.includes(word), `${word} in ${text}`);
      assert.equal(await objectCount(saved, 0), 3, "one object per line");
    } finally {
      model.dispose();
    }
  });

  it("honours alignment, style and the page's rotation", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const rect = { x: 50, y: 40, width: 150, height: 60 };
      model.apply([
        textBox({
          rect,
          text: "Right",
          style: { align: "right", bold: true, color: "#ff0000", fontSize: 20 },
        }),
        textBox({
          pageIndex: 1,
          rect: { x: 20, y: 30, width: 120, height: 50 },
          text: "Turned",
          style: { fontFamily: "Times", italic: true },
        }),
      ]);
      const right = model.getElement("p0:n1.0.0")!;
      assert.ok(
        Math.abs(right.bounds.x + right.bounds.width - (rect.x + rect.width)) <
          2,
        JSON.stringify(right.bounds),
      );
      assert.ok(right.bounds.y >= rect.y - 0.5 && right.bounds.y < rect.y + 8);
      assert.deepEqual(right.textStyle, {
        fontFamily: "Helvetica",
        fontSize: 20,
        bold: true,
        italic: false,
        color: "#ff0000",
      });

      const turned = model.getElement("p1:n1.1.0")!;
      assert.equal(turned.pageIndex, 1);
      assert.ok(
        turned.bounds.x >= 19.5 && turned.bounds.x < 23,
        JSON.stringify(turned.bounds),
      );
      assert.ok(turned.bounds.y >= 29.5 && turned.bounds.y < 36);
      assert.ok(turned.bounds.width > 25 && turned.bounds.width < 60);
      assert.equal(turned.rotation, undefined, "upright on the displayed page");
      assert.equal(turned.textStyle?.italic, true);
      assert.equal(turned.textStyle?.fontFamily, "Times");
      assert.equal(await extractPageText(model.materialize(), 1), "Turned");
    } finally {
      model.dispose();
    }
  });

  it("rejects what it cannot draw and changes nothing", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const shape = checkOperations(
        [
          textBox({ rect: { x: 0, y: 0, width: 0, height: 10 }, text: "x" }),
          { ...textBox({ text: "x" }), style: { color: "red", fontSize: 900 } },
          { ...textBox({ text: "x" }), extra: 1 },
        ],
        pdfOperationSchemas,
      ).map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(shape, [
        "0/rect/width:minimum",
        "1/style/color:pattern",
        "1/style/fontSize:maximum",
        "2/extra:additional-property",
      ]);
      const engine = model
        .validate([
          textBox({ pageIndex: 5, text: "x" }),
          textBox({
            rect: { x: 500, y: 700, width: 200, height: 100 },
            text: "x",
          }),
          textBox({ text: "Привіт" }),
          textBox({ text: "ok", style: { fontFamily: "Comic Sans" } }),
          textBox({ text: "Ünïcode € fine" }),
        ])
        .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(engine, [
        "0/pageIndex:unknown-target",
        "1/rect:range",
        "2/text:font-unavailable",
        "3/style/fontFamily:unknown-font",
      ]);
      assert.deepEqual(model.materialize(), original);
    } finally {
      model.dispose();
    }
  });

  it("warns when the text overflows the box", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const change = model.apply([
        textBox({
          rect: { x: 72, y: 72, width: 100, height: 10 },
          text: "far too much text for ten points of height",
        }),
      ]);
      assert.equal(change.warnings[0]?.code, "fidelity-degraded");
      assert.equal(change.warnings[0]?.details?.elementId, "p0:n1.0.0");
    } finally {
      model.dispose();
    }
  });

  it("is deterministic and replays its history", async () => {
    const first = new PdfEditDocument(pdfium, original);
    const second = new PdfEditDocument(pdfium, original);
    try {
      const batch = [textBox({ text: "Same" })];
      first.apply(batch);
      second.apply(batch);
      const bytes = first.materialize();
      assert.deepEqual(second.materialize(), bytes);
      first.restore([]);
      assert.deepEqual(first.materialize(), original);
      assert.deepEqual(
        first.getElements({ pageIndex: 0 }).map((element) => element.id),
        ["p0:o0"],
      );
      first.restore([batch]);
      assert.deepEqual(first.materialize(), bytes);
      assert.equal(first.getElement("p0:n1.0.0")?.kind, "textBox");
    } finally {
      first.dispose();
      second.dispose();
    }
  });
});
