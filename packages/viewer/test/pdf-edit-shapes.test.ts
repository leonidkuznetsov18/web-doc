import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PageRect, PdfOperation } from "../src/index.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";

const op = <T extends PdfOperation>(operation: T): T => operation;

function near(actual: PageRect, expected: PageRect, slack: number): void {
  for (const key of ["x", "y", "width", "height"] as const)
    assert.ok(
      Math.abs(actual[key] - expected[key]) <= slack,
      `${key}: ${actual[key]} vs ${expected[key]} in ${JSON.stringify(actual)}`,
    );
}

describe("shapes", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;
  let original: Uint8Array;

  before(async () => {
    pdfium = await fixturePdfium();
    original = await buildPdf([
      "One",
      { width: 300, height: 400, rotation: 1 },
    ]);
  });

  it("draws rectangles, ellipses and lines where asked, on rotated pages too", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const change = model.apply([
        op({
          op: "insertShape",
          pageIndex: 0,
          shape: "rectangle",
          rect: { x: 100, y: 200, width: 150, height: 50 },
          fill: { color: "#ff8800" },
        }),
        op({
          op: "insertShape",
          pageIndex: 0,
          shape: "ellipse",
          rect: { x: 300, y: 300, width: 80, height: 40 },
          stroke: { color: "#0000ff", width: 2 },
          fill: { color: "#00ff00" },
        }),
        op({
          op: "insertShape",
          pageIndex: 0,
          shape: "line",
          from: { x: 50, y: 50 },
          to: { x: 250, y: 150 },
          stroke: { color: "#000000", width: 1 },
        }),
        op({
          op: "insertShape",
          pageIndex: 1,
          shape: "rectangle",
          rect: { x: 20, y: 30, width: 60, height: 100 },
          stroke: { color: "#ff0000", width: 0 },
        }),
      ]);
      assert.deepEqual(change.createdIds, [
        "p0:n1.0.0",
        "p0:n1.1.0",
        "p0:n1.2.0",
        "p1:n1.3.0",
      ]);
      const [rectangle, ellipse, line] = model
        .getElements({ pageIndex: 0, kinds: ["shape"] })
        .map((element) => element);
      near(rectangle!.bounds, { x: 100, y: 200, width: 150, height: 50 }, 0.5);
      assert.deepEqual(rectangle!.shapeStyle, { fill: { color: "#ff8800" } });
      // PDFium reports stroked bounds grown by the stroke width on each side.
      near(ellipse!.bounds, { x: 298, y: 298, width: 84, height: 44 }, 0.6);
      assert.deepEqual(ellipse!.shapeStyle, {
        stroke: { color: "#0000ff", width: 2 },
        fill: { color: "#00ff00" },
      });
      // A diagonal stroke's caps push the box out a little further.
      near(line!.bounds, { x: 49, y: 49, width: 202, height: 102 }, 1);
      assert.deepEqual(line!.shapeStyle, {
        stroke: { color: "#000000", width: 1 },
      });
      const turned = model.getElement("p1:n1.3.0")!;
      // A hairline stroke still widens the reported box by half a unit.
      near(turned.bounds, { x: 19.5, y: 29.5, width: 61, height: 101 }, 0.6);

      const reopened = new PdfEditDocument(pdfium, model.materialize());
      try {
        assert.equal(reopened.getElements({ kinds: ["shape"] }).length, 4);
      } finally {
        reopened.dispose();
      }
    } finally {
      model.dispose();
    }
  });

  it("changes and removes stroke and fill", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([
        op({
          op: "insertShape",
          pageIndex: 0,
          shape: "rectangle",
          rect: { x: 100, y: 200, width: 150, height: 50 },
          fill: { color: "#ff8800" },
        }),
      ]);
      model.apply([
        op({
          op: "setShapeStyle",
          target: "p0:n1.0.0",
          stroke: { color: "#123456", width: 3 },
        }),
      ]);
      assert.deepEqual(model.getElement("p0:n1.0.0")!.shapeStyle, {
        stroke: { color: "#123456", width: 3 },
        fill: { color: "#ff8800" },
      });
      model.apply([
        op({ op: "setShapeStyle", target: "p0:n1.0.0", fill: null }),
      ]);
      assert.deepEqual(model.getElement("p0:n1.0.0")!.shapeStyle, {
        stroke: { color: "#123456", width: 3 },
      });
      const issues = model
        .validate([
          op({ op: "setShapeStyle", target: "p0:n1.0.0", stroke: null }),
          op({
            op: "setShapeStyle",
            target: "p0:o0",
            fill: { color: "#000000" },
          }),
          op({
            op: "insertShape",
            pageIndex: 0,
            shape: "rectangle",
            rect: { x: 0, y: 0, width: 10, height: 10 },
          }),
          op({
            op: "insertShape",
            pageIndex: 0,
            shape: "line",
            from: { x: 0, y: 0 },
            to: { x: 700, y: 0 },
            stroke: { color: "#000000", width: 1 },
          }),
          op({
            op: "insertShape",
            pageIndex: 3,
            shape: "ellipse",
            rect: { x: 0, y: 0, width: 1, height: 1 },
          }),
        ])
        .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(issues, [
        "0:required",
        "1/target:unsupported-target",
        "2:required",
        "3/to:range",
        "4/pageIndex:unknown-target",
      ]);
    } finally {
      model.dispose();
    }
  });

  it("replays shapes deterministically", async () => {
    const batch = [
      op({
        op: "insertShape",
        pageIndex: 0,
        shape: "ellipse",
        rect: { x: 10, y: 10, width: 30, height: 20 },
        fill: { color: "#000000" },
      }),
      op({
        op: "setShapeStyle",
        target: "p0:n1.0.0",
        stroke: { color: "#ff0000", width: 1 },
      }),
    ];
    const first = new PdfEditDocument(pdfium, original);
    const second = new PdfEditDocument(pdfium, original);
    try {
      for (const operation of batch) first.apply([operation]);
      second.restore(batch.map((operation) => [operation]));
      assert.deepEqual(second.materialize(), first.materialize());
    } finally {
      first.dispose();
      second.dispose();
    }
  });
});
