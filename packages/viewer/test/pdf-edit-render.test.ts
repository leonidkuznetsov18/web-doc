import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PdfOperation } from "../src/index.js";
import { ViewerError } from "../src/index.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";

/*
 * The suppressed render of ACTION-825, task 33: a page rendered by PDFium
 * with chosen elements inactive, as a stand-in behind a host's input
 * surface. The objects are active again afterwards and the saved bytes do
 * not change.
 */

const op = <T extends PdfOperation>(operation: T): T => operation;

/** The RGBA pixel at page-space point (x, y) of a bitmap rendered at `scale`. */
function pixelAt(
  bitmap: { width: number; height: number; data: ArrayBuffer },
  scale: number,
  x: number,
  y: number,
): [number, number, number, number] {
  const column = Math.min(bitmap.width - 1, Math.floor(x * scale));
  const row = Math.min(bitmap.height - 1, Math.floor(y * scale));
  const bytes = new Uint8Array(bitmap.data);
  const at = (row * bitmap.width + column) * 4;
  return [bytes[at]!, bytes[at + 1]!, bytes[at + 2]!, bytes[at + 3]!];
}

function sameBytes(a: ArrayBuffer, b: ArrayBuffer): boolean {
  if (a.byteLength !== b.byteLength) return false;
  const x = new Uint8Array(a);
  const y = new Uint8Array(b);
  for (let index = 0; index < x.length; index += 1)
    if (x[index] !== y[index]) return false;
  return true;
}

describe("suppressed render (overlay primitives)", () => {
  it("leaves the element out, keeps the rest and restores it for the next render", async () => {
    const pdfium = await fixturePdfium();
    // A red square at user-space (100, 600)–(200, 700): page-space y 92–192.
    const original = await buildPdf([
      {
        rect: { x: 100, y: 600, width: 100, height: 100, fill: [255, 0, 0] },
        texts: [{ text: "Kept", x: 300, y: 400, fontSize: 36 }],
      },
    ]);
    const model = new PdfEditDocument(pdfium, original);
    try {
      const before = model.materialize("save", "full");
      const normal = model.renderPageWithout(0, [], 1);
      assert.equal(normal.width, 612);
      assert.equal(normal.height, 792);
      assert.equal(normal.data.byteLength, 612 * 792 * 4);
      assert.deepEqual(pixelAt(normal, 1, 150, 142), [255, 0, 0, 255]);

      // The builder draws texts before shapes: the square is the second object.
      const square = model
        .getElements({ pageIndex: 0 })
        .find((element) => element.kind === "shape")!;
      assert.equal(square.id, "p0:o1");
      const without = model.renderPageWithout(0, [square.id], 1);
      assert.deepEqual(
        pixelAt(without, 1, 150, 142),
        [255, 255, 255, 255],
        "the square is gone",
      );
      // Outside the square the two renders agree, text included.
      let differentOutside = 0;
      const a = new Uint8Array(normal.data);
      const b = new Uint8Array(without.data);
      for (let row = 0; row < 792; row += 1)
        for (let column = 0; column < 612; column += 1) {
          if (row >= 90 && row <= 194 && column >= 98 && column <= 202)
            continue;
          const at = (row * 612 + column) * 4;
          if (
            a[at] !== b[at] ||
            a[at + 1] !== b[at + 1] ||
            a[at + 2] !== b[at + 2]
          )
            differentOutside += 1;
        }
      assert.equal(differentOutside, 0, "nothing else changed");

      const again = model.renderPageWithout(0, [], 1);
      assert.ok(sameBytes(again.data, normal.data), "the square is back");
      const after = model.materialize("save", "full");
      assert.deepEqual(after, before, "the bytes did not change");
    } finally {
      model.dispose();
    }
  });

  it("renders at a scale, ignores unknown ids and respects the raster limit", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf([
      {
        width: 200,
        height: 100,
        rect: { x: 50, y: 25, width: 100, height: 50, fill: [0, 0, 255] },
      },
    ]);
    const model = new PdfEditDocument(pdfium, original, undefined, {
      ...(await import("../src/index.js")).defaultResourceLimits,
      maxDecodedPixels: 200 * 100 * 4,
    });
    try {
      const scaled = model.renderPageWithout(0, ["nope"], 2);
      assert.equal(scaled.width, 400);
      assert.equal(scaled.height, 200);
      assert.deepEqual(pixelAt(scaled, 2, 100, 50), [0, 0, 255, 255]);
      assert.throws(
        () => model.renderPageWithout(0, [], 3),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "resource-limit",
      );
      assert.throws(
        () => model.renderPageWithout(1, [], 1),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
    } finally {
      model.dispose();
    }
  });

  it("hides every object of a text box and shows them again", async () => {
    const pdfium = await fixturePdfium();
    const model = new PdfEditDocument(pdfium, await buildPdf([{}]));
    try {
      const { createdIds } = model.apply([
        op({
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 50, y: 50, width: 300, height: 100 },
          text: "Box line one\nBox line two",
          style: { fontSize: 24 },
        }),
      ]);
      const normal = model.renderPageWithout(0, [], 1);
      const without = model.renderPageWithout(0, createdIds, 1);
      const inked = (bitmap: { width: number; data: ArrayBuffer }) => {
        const bytes = new Uint8Array(bitmap.data);
        let count = 0;
        for (let at = 0; at < bytes.length; at += 4)
          if (bytes[at]! < 128) count += 1;
        return count;
      };
      assert.ok(inked(normal) > 100, "the box draws ink");
      assert.equal(inked(without), 0, "the box is gone entirely");
      assert.ok(sameBytes(model.renderPageWithout(0, [], 1).data, normal.data));
    } finally {
      model.dispose();
    }
  });
});
