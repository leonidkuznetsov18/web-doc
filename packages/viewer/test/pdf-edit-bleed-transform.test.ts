import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { ViewerError, type PageRect, type PdfOperation } from "../src/index.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

const PAGE_SIZE = 200;

function imagePdf(bounds: PageRect): Promise<Uint8Array> {
  return buildPdf([
    {
      width: PAGE_SIZE,
      height: PAGE_SIZE,
      image: { ...bounds, y: PAGE_SIZE - bounds.y - bounds.height },
    },
  ]);
}

function near(actual: PageRect, expected: PageRect): void {
  for (const key of ["x", "y", "width", "height"] as const)
    assert.ok(
      Math.abs(actual[key] - expected[key]) < 0.002,
      `${key}: ${actual[key]} vs ${expected[key]}`,
    );
}

function rangeError(
  operationIndex = 0,
  message?: RegExp,
): (error: unknown) => boolean {
  return (error) => {
    assert.ok(error instanceof ViewerError);
    assert.equal(error.code, "invalid-operation");
    const issues = error.details?.issues;
    assert.ok(Array.isArray(issues));
    assert.equal(issues.length, 1);
    const issue: unknown = issues[0];
    assert.ok(issue && typeof issue === "object");
    assert.ok("operationIndex" in issue && "code" in issue);
    assert.equal(issue.operationIndex, operationIndex);
    assert.equal(issue.code, "range");
    if (message) assert.match(error.message, message);
    return true;
  };
}

describe("transforms of imported PDF objects outside the crop box", () => {
  for (const { edge, bounds, by } of [
    {
      edge: "left",
      bounds: { x: -20, y: 60, width: 50, height: 40 },
      by: { dx: 1, dy: 0 },
    },
    {
      edge: "right",
      bounds: { x: 170, y: 60, width: 50, height: 40 },
      by: { dx: -1, dy: 0 },
    },
    {
      edge: "top",
      bounds: { x: 60, y: -20, width: 40, height: 50 },
      by: { dx: 0, dy: 1 },
    },
    {
      edge: "bottom",
      bounds: { x: 60, y: 170, width: 40, height: 50 },
      by: { dx: 0, dy: -1 },
    },
  ]) {
    it(`moves an image one point inward from the ${edge} without changing its other axis or size`, async () => {
      const { session, end } = await pdfSession(await imagePdf(bounds));
      try {
        await session.moveElement({ target: "p0:o0", by });
        const moved = (await session.getElement("p0:o0")).item;
        assert.ok(moved);
        assert.deepEqual(moved.bounds, {
          ...bounds,
          x: bounds.x + by.dx,
          y: bounds.y + by.dy,
        });
      } finally {
        await end();
      }
    });
  }

  it("accepts unchanged bounds and preserves two-edge bleed during an orthogonal move", async () => {
    const bounds = { x: 50, y: -7, width: 14, height: 214 };
    const { session, end } = await pdfSession(await imagePdf(bounds));
    try {
      await session.moveElement({ target: "p0:o0", by: { dx: 0, dy: 0 } });
      await session.resizeElement({ target: "p0:o0", rect: bounds });
      await session.moveElement({ target: "p0:o0", to: { x: 51, y: -7 } });
      const moved = (await session.getElement("p0:o0")).item;
      assert.ok(moved);
      assert.deepEqual(moved.bounds, { ...bounds, x: 51 });
    } finally {
      await end();
    }
  });

  it("keeps inward moves and resizing as one undoable edit through Save and reopen", async () => {
    const bounds = { x: -20, y: -30, width: 80, height: 100 };
    const original = await imagePdf(bounds);
    const { session, end } = await pdfSession(original);
    try {
      const receipt = await session.apply([
        { op: "moveElement", target: "p0:o0", by: { dx: 1, dy: 2 } },
        {
          op: "resizeElement",
          target: "p0:o0",
          rect: { x: -10, y: -20, width: 70, height: 90 },
        },
      ]);
      const expected = { x: -10, y: -20, width: 70, height: 90 };
      assert.equal(receipt.operationCount, 2);
      assert.deepEqual(receipt.changedPages, [0]);
      assert.deepEqual(receipt.createdIds, []);
      assert.deepEqual(receipt.removedIds, []);
      assert.deepEqual(
        (await session.getElement("p0:o0")).item?.bounds,
        expected,
      );
      await session.undo();
      assert.deepEqual(
        (await session.getElement("p0:o0")).item?.bounds,
        bounds,
      );
      assert.equal(session.state.canUndo, false);
      assert.deepEqual((await session.save()).bytes, original);
      await session.redo();
      assert.deepEqual(
        (await session.getElement("p0:o0")).item?.bounds,
        expected,
      );
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.deepEqual(
          (await reopened.session.getElement("p0:o0")).item?.bounds,
          expected,
        );
        await reopened.session.moveElement({
          target: "p0:o0",
          by: { dx: 1, dy: 1 },
        });
        assert.deepEqual(
          (await reopened.session.getElement("p0:o0")).item?.bounds,
          { ...expected, x: -9, y: -19 },
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });

  for (const { rotation, bounds, by } of [
    {
      rotation: 0,
      bounds: { x: -10, y: 50, width: 50, height: 40 },
      by: { dx: 1, dy: 0 },
    },
    {
      rotation: 1,
      bounds: { x: 50, y: -10, width: 40, height: 50 },
      by: { dx: 0, dy: 1 },
    },
    {
      rotation: 2,
      bounds: { x: 120, y: 50, width: 50, height: 40 },
      by: { dx: -1, dy: 0 },
    },
    {
      rotation: 3,
      bounds: { x: 50, y: 120, width: 40, height: 50 },
      by: { dx: 0, dy: -1 },
    },
  ] as const) {
    it(`uses displayed crop coordinates on a page rotated ${rotation * 90} degrees`, async () => {
      const original = await buildPdf([
        {
          width: 200,
          height: 200,
          cropBox: [20, 30, 180, 170],
          rotation,
          image: { x: 10, y: 80, width: 50, height: 40 },
        },
      ]);
      const { session, end } = await pdfSession(original);
      try {
        assert.deepEqual(
          (await session.getElement("p0:o0")).item?.bounds,
          bounds,
        );
        await session.moveElement({ target: "p0:o0", by });
        assert.deepEqual((await session.getElement("p0:o0")).item?.bounds, {
          ...bounds,
          x: bounds.x + by.dx,
          y: bounds.y + by.dy,
        });
      } finally {
        await end();
      }
    });
  }

  it("uses ordinary text and stroked path bounds without dropping their native styles", async () => {
    const original = await buildPdf([
      {
        width: 200,
        height: 200,
        texts: [{ text: "Hello", x: -5, y: 140 }],
        rect: {
          x: 50,
          y: -7,
          width: 14,
          height: 214,
          stroke: [255, 255, 255],
          strokeWidth: 2,
        },
      },
    ]);
    const { session, end } = await pdfSession(original);
    try {
      for (const target of ["p0:o0", "p0:o1"]) {
        const before = (await session.getElement(target)).item;
        assert.ok(before);
        assert.ok(before.operations.includes("moveElement"));
        assert.ok(before.bounds.x < 0 || before.bounds.y < 0);
        await session.moveElement({ target, by: { dx: 1, dy: 0 } });
        const after = (await session.getElement(target)).item;
        assert.ok(after);
        near(after.bounds, { ...before.bounds, x: before.bounds.x + 1 });
        assert.equal(after.text, before.text);
        assert.deepEqual(after.textStyle, before.textStyle);
        assert.deepEqual(after.shapeStyle, before.shapeStyle);
      }
    } finally {
      await end();
    }
  });

  for (const { name, bounds, operation } of [
    {
      name: "increased existing overflow",
      bounds: { x: -20, y: 60, width: 50, height: 40 },
      operation: { op: "moveElement", target: "p0:o0", by: { dx: -1, dy: 0 } },
    },
    {
      name: "a new opposite overflow edge",
      bounds: { x: -20, y: 60, width: 50, height: 40 },
      operation: { op: "moveElement", target: "p0:o0", to: { x: 180, y: 60 } },
    },
    {
      name: "a two-edge bleed moved toward either edge",
      bounds: { x: 50, y: -7, width: 14, height: 214 },
      operation: { op: "moveElement", target: "p0:o0", by: { dx: 0, dy: 1 } },
    },
    {
      name: "a resized object entirely outside the page",
      bounds: { x: -20, y: 60, width: 50, height: 40 },
      operation: {
        op: "resizeElement",
        target: "p0:o0",
        rect: { x: -20, y: 60, width: 19, height: 40 },
      },
    },
    {
      name: "an outside object only touching the page edge",
      bounds: { x: -20, y: 60, width: 50, height: 40 },
      operation: {
        op: "resizeElement",
        target: "p0:o0",
        rect: { x: -20, y: 60, width: 20, height: 40 },
      },
    },
  ] satisfies readonly {
    name: string;
    bounds: PageRect;
    operation: PdfOperation;
  }[]) {
    it(`rejects ${name} without changing history or saved bytes`, async () => {
      const original = await imagePdf(bounds);
      const { session, end } = await pdfSession(original);
      try {
        const before = session.state;
        await assert.rejects(
          session.apply([operation]),
          rangeError(0, /preserve or reduce.*partly visible/),
        );
        assert.deepEqual(session.state, before);
        assert.deepEqual(
          (await session.getElement("p0:o0")).item?.bounds,
          bounds,
        );
        assert.deepEqual((await session.save()).bytes, original);
      } finally {
        await end();
      }
    });
  }

  it("revalidates the reduced overflow before a later operation and rolls the batch back", async () => {
    const bounds = { x: -20, y: 60, width: 50, height: 40 };
    const original = await imagePdf(bounds);
    const { session, end } = await pdfSession(original);
    try {
      const before = session.state;
      await assert.rejects(
        session.apply([
          { op: "moveElement", target: "p0:o0", by: { dx: 1, dy: 0 } },
          { op: "resizeElement", target: "p0:o0", rect: bounds },
        ]),
        rangeError(1),
      );
      assert.deepEqual(session.state, before);
      assert.deepEqual(
        (await session.getElement("p0:o0")).item?.bounds,
        bounds,
      );
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("does not allow repeated outward steps smaller than the page tolerance", async () => {
    const bounds = { x: -20, y: 60, width: 50, height: 40 };
    const original = await imagePdf(bounds);
    const { session, end } = await pdfSession(original);
    try {
      for (let step = 0; step < 3; step += 1)
        await assert.rejects(
          session.moveElement({ target: "p0:o0", by: { dx: -0.005, dy: 0 } }),
          rangeError(),
        );
      assert.deepEqual(
        (await session.getElement("p0:o0")).item?.bounds,
        bounds,
      );
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("does not allow inserts to inherit another object's overflow", async () => {
    const original = await imagePdf({ x: -20, y: 60, width: 50, height: 40 });
    const { session, end } = await pdfSession(original);
    try {
      await assert.rejects(
        session.insertShape({
          pageIndex: 0,
          shape: "rectangle",
          rect: { x: -1, y: 60, width: 30, height: 40 },
          fill: { color: "#000000" },
        }),
        rangeError(0, /must lie within the 200×200 pt page/),
      );
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("keeps crop-filling shading transforms bounded by their reported clipping rectangle", async () => {
    const content = "/Sh1 sh";
    const objects = [
      "<< /Type /Catalog /Pages 2 0 R >>",
      "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /CropBox [20 20 180 180] /Resources << /Shading << /Sh1 4 0 R >> >> /Contents 5 0 R >>",
      "<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 100 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [0 1 1] /N 1 >> /Extend [true true] /BBox [0 40 100 120] >>",
      `<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
    ];
    let pdf = "%PDF-1.7\n";
    const offsets: number[] = [];
    for (const [index, object] of objects.entries()) {
      offsets.push(pdf.length);
      pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
    }
    const xref = pdf.length;
    pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
    for (const offset of offsets)
      pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
    pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
    const { session, end } = await pdfSession(new TextEncoder().encode(pdf));
    try {
      const before = (await session.getElement("p0:o0")).item;
      assert.ok(before);
      assert.equal(before.kind, "other");
      assert.ok(before.operations.includes("moveElement"));
      assert.deepEqual(before.bounds, { x: 0, y: 0, width: 160, height: 160 });
      await session.moveElement({ target: before.id, by: { dx: 0, dy: 0 } });
      const after = (await session.getElement(before.id)).item;
      assert.ok(after);
      assert.deepEqual(after.bounds, before.bounds);
      const saved = (await session.save()).bytes;
      await assert.rejects(
        session.moveElement({ target: before.id, by: { dx: 1, dy: 0 } }),
        rangeError(),
      );
      assert.deepEqual((await session.save()).bytes, saved);
      const reopened = await pdfSession((await session.save()).bytes);
      try {
        assert.deepEqual(
          (await reopened.session.getElement(before.id)).item?.bounds,
          after.bounds,
        );
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
});
