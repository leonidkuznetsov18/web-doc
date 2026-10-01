import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import {
  displayedSize,
  userToPage,
  type PageGeometry,
} from "../src/edit/pdf/engine/geometry.js";
import type { PdfElement } from "../src/index.js";
import {
  buildPdf,
  fixturePdfium,
  type FixturePage,
} from "./fixtures/pdf-builder.js";

const rotations = [0, 1, 2, 3] as const;

/** Geometry as PDFium reports it for a loaded page. */
function geometryOf(
  pdfium: Awaited<ReturnType<typeof fixturePdfium>>,
  page: number,
): PageGeometry {
  const box = pdfium.readNumbers(4, "float", ([pointer]) =>
    pdfium.lib.FPDF_GetPageBoundingBox(page, pointer!),
  ) as [number, number, number, number];
  return {
    box: { left: box[0], top: box[1], right: box[2], bottom: box[3] },
    rotation: pdfium.lib.FPDFPage_GetRotation(page),
  };
}

function close(actual: number, expected: number, tolerance: number): void {
  assert.ok(
    Math.abs(actual - expected) <= tolerance,
    `expected ${actual} to be within ${tolerance} of ${expected}`,
  );
}

describe("PDF document model", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;

  before(async () => {
    pdfium = await fixturePdfium();
  });

  it("maps user space to page space like PDFium on rotated and cropped pages", async () => {
    const pages: FixturePage[] = rotations.flatMap((rotation) => [
      { rotation, width: 612, height: 792 },
      { rotation, width: 612, height: 792, cropBox: [50, 60, 350, 460] },
    ]);
    const bytes = await buildPdf(pages);
    const document = pdfium.openDocument(bytes);
    try {
      const SCALE = 1000;
      pages.forEach((_, pageIndex) => {
        const page = pdfium.lib.FPDF_LoadPage(document.handle, pageIndex);
        const geometry = geometryOf(pdfium, page);
        const size = displayedSize(geometry);
        assert.equal(size.width, pdfium.lib.FPDF_GetPageWidthF(page));
        assert.equal(size.height, pdfium.lib.FPDF_GetPageHeightF(page));
        for (const [x, y] of [
          [0, 0],
          [612, 792],
          [50, 60],
          [350, 460],
          [123.25, 456.75],
          [-10, 900],
        ]) {
          const device = pdfium.readNumbers(2, "i32", ([dx, dy]) =>
            pdfium.lib.FPDF_PageToDevice(
              page,
              0,
              0,
              Math.round(size.width * SCALE),
              Math.round(size.height * SCALE),
              0,
              x!,
              y!,
              dx!,
              dy!,
            ),
          ) as [number, number];
          const point = userToPage(geometry, x!, y!);
          close(point.x, device[0] / SCALE, 0.01);
          close(point.y, device[1] / SCALE, 0.01);
        }
        pdfium.lib.FPDF_ClosePage(page);
      });
    } finally {
      document.close();
    }
  });

  it("lists elements with kinds, text, styles and page-space bounds", async () => {
    const bytes = await buildPdf([
      {
        texts: [
          { text: "Hello", x: 72, y: 700 },
          {
            text: "Loud",
            x: 72,
            y: 600,
            font: "Helvetica-Bold",
            fontSize: 20,
            color: [255, 0, 0],
          },
        ],
        rect: { x: 100, y: 100, width: 50, height: 30, fill: [0, 128, 255] },
        image: { x: 300, y: 500, width: 160, height: 80 },
      },
      {
        width: 300,
        height: 400,
        rotation: 1,
        image: { x: 10, y: 20, width: 100, height: 50 },
        rect: {
          x: 150,
          y: 150,
          width: 40,
          height: 40,
          stroke: [0, 0, 0],
          strokeWidth: 2,
        },
      },
    ]);
    const model = new PdfEditDocument(pdfium, bytes);
    try {
      assert.equal(model.pageCount, 2);
      const first = model.getElements({ pageIndex: 0 });
      assert.deepEqual(
        first.map((element) => element.kind),
        ["text", "text", "shape", "image"],
      );
      const [hello, loud, rect, image] = first as [
        PdfElement,
        PdfElement,
        PdfElement,
        PdfElement,
      ];
      assert.equal(hello.id, "p0:o0");
      assert.equal(hello.text, "Hello");
      assert.deepEqual(hello.textStyle, {
        fontFamily: "Helvetica",
        fontSize: 12,
        bold: false,
        italic: false,
        color: "#000000",
      });
      close(hello.bounds.x, 72, 1);
      assert.ok(
        hello.bounds.y > 78 && hello.bounds.y < 92,
        String(hello.bounds.y),
      );
      assert.ok(hello.bounds.width > 20);
      assert.equal(loud.textStyle?.bold, true);
      assert.equal(loud.textStyle?.fontSize, 20);
      assert.equal(loud.textStyle?.color, "#ff0000");
      assert.deepEqual(rect.bounds, { x: 100, y: 662, width: 50, height: 30 });
      assert.deepEqual(rect.shapeStyle, { fill: { color: "#0080ff" } });
      assert.deepEqual(image.bounds, {
        x: 300,
        y: 212,
        width: 160,
        height: 80,
      });
      assert.deepEqual(image.operations, [
        "moveElement",
        "resizeElement",
        "deleteElement",
      ]);
      assert.equal(hello.rotation, undefined);

      const second = model.getElements({ pageIndex: 1, kinds: ["image"] });
      assert.equal(second.length, 1);
      // A 300×400 page turned a quarter clockwise shows 400×300; the image at
      // user (10, 20)–(110, 70) lands at the top-left, 50 wide and 100 tall.
      assert.deepEqual(second[0]!.bounds, {
        x: 20,
        y: 10,
        width: 50,
        height: 100,
      });
      assert.equal(second[0]!.rotation, 90);
      const stroked = model.getElements({ pageIndex: 1, kinds: ["shape"] })[0]!;
      assert.deepEqual(stroked.shapeStyle, {
        stroke: { color: "#000000", width: 2 },
      });
      assert.deepEqual(
        model
          .getElements({
            pageIndex: 0,
            intersects: { x: 0, y: 650, width: 200, height: 50 },
          })
          .map((element) => element.id),
        ["p0:o2"],
      );
      assert.deepEqual(model.getElements({ pageIndex: 7 }), []);
    } finally {
      model.dispose();
    }
  });

  it("answers point and text queries", async () => {
    const bytes = await buildPdf([
      {
        texts: [
          { text: "Hello world", x: 72, y: 700 },
          { text: "hello again", x: 72, y: 650 },
        ],
        image: { x: 300, y: 500, width: 160, height: 80 },
      },
      "Hello elsewhere",
    ]);
    const model = new PdfEditDocument(pdfium, bytes);
    try {
      assert.deepEqual(
        model.elementsAt(0, { x: 310, y: 220 }).map((element) => element.id),
        ["p0:o2"],
      );
      assert.deepEqual(model.elementsAt(0, { x: 5, y: 5 }), []);
      assert.deepEqual(model.elementsAt(9, { x: 5, y: 5 }), []);

      const hits = model.findText("hello", {});
      assert.deepEqual(
        hits.map((hit) => [hit.pageIndex, hit.text, hit.elementIds]),
        [
          [0, "Hello", ["p0:o0"]],
          [0, "hello", ["p0:o1"]],
          [1, "Hello", ["p1:o0"]],
        ],
      );
      const [first] = hits;
      assert.equal(first!.rects.length, 1);
      close(first!.rects[0]!.x, 72, 1);
      assert.ok(first!.rects[0]!.width > 20);
      assert.ok(first!.rects[0]!.y > 78 && first!.rects[0]!.y < 92);
      assert.equal(model.findText("hello", { caseSensitive: true }).length, 1);
      assert.equal(model.findText("hello", { maxResults: 2 }).length, 2);
      assert.deepEqual(
        model
          .findText("hello", { pageRange: [1, 1] })
          .map((hit) => hit.pageIndex),
        [1],
      );
      assert.deepEqual(model.findText("absent", {}), []);
      assert.deepEqual(model.findText("", {}), []);
    } finally {
      model.dispose();
    }
  });

  it("groups marked objects into composite elements and ignores foreign marks", async () => {
    const mark = { kind: "textBox", id: "p0:n1.0.0", text: "Box\nText" };
    const bytes = await buildPdf([
      {
        texts: [
          { text: "Box", x: 72, y: 700, mark },
          { text: "Text", x: 72, y: 686, mark },
          { text: "Plain", x: 72, y: 600, brokenMark: true },
          { text: "Foreign", x: 72, y: 500, mark: { kind: "table", id: "x1" } },
        ],
      },
    ]);
    const model = new PdfEditDocument(pdfium, bytes);
    try {
      const elements = model.getElements({ pageIndex: 0 });
      assert.deepEqual(
        elements.map((element) => [element.kind, element.id]),
        [
          ["textBox", "p0:n1.0.0"],
          ["text", "p0:o2"],
          ["table", "p0:x1"],
        ],
      );
      const box = elements[0]!;
      assert.equal(box.text, "Box\nText");
      assert.ok(box.operations.includes("replaceText"));
      close(box.bounds.x, 72, 1);
      assert.ok(box.bounds.height > 20, "union of both lines");
      assert.equal(model.getElement("p0:n1.0.0")?.kind, "textBox");
      assert.equal(model.getElement("p0:o2")?.text, "Plain");
      assert.equal(model.getElement("nope"), undefined);
      assert.equal(model.getElement("p0:zzz"), undefined);
    } finally {
      model.dispose();
    }
  });

  it("keeps ids stable across restore and hands the original back", async () => {
    const bytes = await buildPdf(["One", "Two"]);
    const model = new PdfEditDocument(pdfium, bytes);
    try {
      const before = model.getElements({}).map((element) => element.id);
      assert.deepEqual(before, ["p0:o0", "p1:o0"]);
      model.restore([]);
      assert.deepEqual(
        model.getElements({}).map((element) => element.id),
        before,
      );
      assert.deepEqual(model.materialize(), bytes);
      assert.deepEqual(
        model.validate([{ op: "x" }])[0]?.code,
        "unknown-operation",
      );
    } finally {
      model.dispose();
    }
  });
});
