import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import { ViewerError, type PageRect, type PdfOperation } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

function near(actual: PageRect, expected: PageRect, slack = 0.5): void {
  for (const key of ["x", "y", "width", "height"] as const)
    assert.ok(
      Math.abs(actual[key] - expected[key]) <= slack,
      `${key}: ${actual[key]} vs ${expected[key]} in ${JSON.stringify(actual)}`,
    );
}

const op = <T extends PdfOperation>(operation: T): T => operation;

describe("moveElement, resizeElement and deleteElement", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;
  let original: Uint8Array;

  before(async () => {
    pdfium = await fixturePdfium();
    original = await buildPdf([
      {
        texts: [{ text: "Hello", x: 72, y: 700 }],
        rect: { x: 100, y: 100, width: 50, height: 30, fill: [0, 128, 255] },
        image: { x: 300, y: 500, width: 160, height: 80 },
      },
      {
        width: 300,
        height: 400,
        rotation: 1,
        image: { x: 10, y: 20, width: 100, height: 50 },
      },
    ]);
  });

  it("moves plain objects by an offset or to a point, on rotated pages too", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const change = model.apply([
        op({ op: "moveElement", target: "p0:o2", by: { dx: 10, dy: -5 } }),
        op({ op: "moveElement", target: "p0:o0", to: { x: 200, y: 50 } }),
        op({ op: "moveElement", target: "p1:o0", by: { dx: 5, dy: 7 } }),
      ]);
      assert.deepEqual(change.changedPages, [0, 1]);
      near(model.getElement("p0:o2")!.bounds, {
        x: 310,
        y: 207,
        width: 160,
        height: 80,
      });
      const text = model.getElement("p0:o0")!.bounds;
      assert.ok(Math.abs(text.x - 200) <= 0.6 && Math.abs(text.y - 50) <= 0.6);
      near(model.getElement("p1:o0")!.bounds, {
        x: 25,
        y: 17,
        width: 50,
        height: 100,
      });

      // The moved objects survive saving and reopening.
      const reopened = new PdfEditDocument(pdfium, model.materialize());
      try {
        near(reopened.getElement("p1:o0")!.bounds, {
          x: 25,
          y: 17,
          width: 50,
          height: 100,
        });
        assert.equal(await extractPageText(model.materialize(), 0), "Hello");
      } finally {
        reopened.dispose();
      }
    } finally {
      model.dispose();
    }
  });

  it("resizes images and shapes to exact bounds", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([
        op({
          op: "resizeElement",
          target: "p0:o2",
          rect: { x: 50, y: 60, width: 80, height: 40 },
        }),
        op({
          op: "resizeElement",
          target: "p0:o1",
          rect: { x: 400, y: 600, width: 25, height: 90 },
        }),
        op({
          op: "resizeElement",
          target: "p1:o0",
          rect: { x: 30, y: 40, width: 60, height: 120 },
        }),
      ]);
      near(model.getElement("p0:o2")!.bounds, {
        x: 50,
        y: 60,
        width: 80,
        height: 40,
      });
      near(model.getElement("p0:o1")!.bounds, {
        x: 400,
        y: 600,
        width: 25,
        height: 90,
      });
      near(model.getElement("p1:o0")!.bounds, {
        x: 30,
        y: 40,
        width: 60,
        height: 120,
      });
    } finally {
      model.dispose();
    }
  });

  it("deletes elements and keeps the other ids", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([op({ op: "deleteElement", target: "p0:o1" })]);
      assert.deepEqual(
        model.getElements({ pageIndex: 0 }).map((element) => element.id),
        ["p0:o0", "p0:o2"],
      );
      assert.equal(model.getElement("p0:o1"), undefined);
      near(model.getElement("p0:o2")!.bounds, {
        x: 300,
        y: 212,
        width: 160,
        height: 80,
      });
      const reopened = new PdfEditDocument(pdfium, model.materialize());
      try {
        assert.equal(reopened.getElements({ pageIndex: 0 }).length, 2);
      } finally {
        reopened.dispose();
      }
    } finally {
      model.dispose();
    }
  });

  it("moves and deletes text boxes through their stored inputs", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([
        op({
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 72, y: 300, width: 200, height: 40 },
          text: "Box",
        }),
      ]);
      const before = model.getElement("p0:n1.0.0")!.bounds;
      model.apply([
        op({ op: "moveElement", target: "p0:n1.0.0", by: { dx: 10, dy: 20 } }),
      ]);
      const moved = model.getElement("p0:n1.0.0")!.bounds;
      near(moved, { ...before, x: before.x + 10, y: before.y + 20 }, 0.6);
      // A later rebuild starts from the moved rectangle, not the original one.
      model.apply([
        op({ op: "replaceText", target: "p0:n1.0.0", text: "Box again" }),
      ]);
      const rebuilt = model.getElement("p0:n1.0.0")!.bounds;
      assert.ok(
        Math.abs(rebuilt.x - moved.x) < 0.6 &&
          Math.abs(rebuilt.y - moved.y) < 0.6,
      );

      model.apply([op({ op: "deleteElement", target: "p0:n1.0.0" })]);
      assert.equal(model.getElement("p0:n1.0.0"), undefined);
      assert.equal(await extractPageText(model.materialize(), 0), "Hello");
    } finally {
      model.dispose();
    }
  });

  it("validates targets, offsets and page bounds", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const issues = model
        .validate([
          op({ op: "moveElement", target: "p0:o2" }),
          op({
            op: "moveElement",
            target: "p0:o2",
            to: { x: 1, y: 1 },
            by: { dx: 1, dy: 1 },
          }),
          op({ op: "moveElement", target: "p0:o2", by: { dx: 500, dy: 0 } }),
          op({ op: "moveElement", target: "nope", by: { dx: 1, dy: 1 } }),
          op({ op: "deleteElement", target: "p9:o0" }),
          op({
            op: "resizeElement",
            target: "p0:o2",
            rect: { x: 600, y: 0, width: 50, height: 50 },
          }),
        ])
        .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(issues, [
        "0:one-of",
        "1:one-of",
        "2/by:range",
        "3/target:unknown-target",
        "4/target:unknown-target",
        "5/rect:range",
      ]);
      assert.deepEqual(model.materialize(), original);
    } finally {
      model.dispose();
    }
  });

  it("replays transforms deterministically", async () => {
    const batch = [
      op({ op: "moveElement", target: "p0:o2", by: { dx: 3, dy: 4 } }),
      op({
        op: "resizeElement",
        target: "p0:o1",
        rect: { x: 10, y: 10, width: 20, height: 20 },
      }),
      op({ op: "deleteElement", target: "p0:o0" }),
    ];
    const first = new PdfEditDocument(pdfium, original);
    const second = new PdfEditDocument(pdfium, original);
    try {
      first.apply(batch);
      second.restore([batch]);
      assert.deepEqual(second.materialize(), first.materialize());
      first.restore([]);
      assert.deepEqual(first.materialize(), original);
    } finally {
      first.dispose();
      second.dispose();
    }
  });
});

describe("sequential PDF transform batch validation", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf([
      {
        width: 200,
        height: 200,
        image: { x: 80, y: 80, width: 40, height: 40 },
      },
    ]);
  });

  for (const dryRun of [false, true]) {
    it(`rejects cumulative off-page moves atomically${dryRun ? " in a dry run" : ""}`, async () => {
      const { session, end } = await pdfSession(original);
      try {
        const state = session.state;
        const before = (await session.getElement("p0:o0")).item;
        assert.ok(before);
        await assert.rejects(
          session.apply(
            Array.from({ length: 3 }, () =>
              op({
                op: "moveElement",
                target: "p0:o0",
                by: { dx: 70, dy: 0 },
              }),
            ),
            { dryRun },
          ),
          (error: unknown) => {
            assert.ok(error instanceof ViewerError);
            assert.equal(error.code, "invalid-operation");
            assert.deepEqual(error.details?.issues, [
              {
                operationIndex: 1,
                path: "/by",
                code: "range",
                message: "The rectangle must lie within the 200×200 pt page",
              },
            ]);
            return true;
          },
        );
        assert.deepEqual(session.state, state);
        assert.deepEqual((await session.getElement("p0:o0")).item, before);
        assert.deepEqual((await session.save()).bytes, original);
      } finally {
        await end();
      }
    });
  }

  it("uses the size from an earlier resize when validating a later move", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const state = session.state;
      const before = (await session.getElement("p0:o0")).item;
      await assert.rejects(
        session.apply([
          {
            op: "resizeElement",
            target: "p0:o0",
            rect: { x: 80, y: 80, width: 100, height: 40 },
          },
          { op: "moveElement", target: "p0:o0", by: { dx: 30, dy: 0 } },
        ]),
        (error: unknown) => {
          assert.ok(error instanceof ViewerError);
          assert.equal(error.code, "invalid-operation");
          assert.deepEqual(error.details?.issues, [
            {
              operationIndex: 1,
              path: "/by",
              code: "range",
              message: "The rectangle must lie within the 200×200 pt page",
            },
          ]);
          return true;
        },
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual((await session.getElement("p0:o0")).item, before);
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("reports a target deleted earlier in the batch without committing its deletion", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const state = session.state;
      const before = (await session.getElement("p0:o0")).item;
      await assert.rejects(
        session.apply([
          { op: "deleteElement", target: "p0:o0" },
          { op: "moveElement", target: "p0:o0", by: { dx: 1, dy: 0 } },
        ]),
        (error: unknown) => {
          assert.ok(error instanceof ViewerError);
          assert.equal(error.code, "invalid-operation");
          assert.deepEqual(error.details?.issues, [
            {
              operationIndex: 1,
              path: "/target",
              code: "unknown-target",
              message: "No element p0:o0",
            },
          ]);
          return true;
        },
      );
      assert.deepEqual(session.state, state);
      assert.deepEqual((await session.getElement("p0:o0")).item, before);
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });

  it("keeps a valid sequential batch as one undoable, saveable edit", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const receipt = await session.apply([
        { op: "moveElement", target: "p0:o0", by: { dx: 10, dy: 0 } },
        { op: "moveElement", target: "p0:o0", by: { dx: 10, dy: 0 } },
        {
          op: "resizeElement",
          target: "p0:o0",
          rect: { x: 100, y: 90, width: 50, height: 30 },
        },
        { op: "moveElement", target: "p0:o0", by: { dx: 20, dy: 0 } },
      ]);
      assert.equal(receipt.operationCount, 4);
      assert.equal(receipt.revision, 1);
      assert.deepEqual(receipt.changedPages, [0]);
      assert.deepEqual(receipt.createdIds, []);
      assert.deepEqual(receipt.removedIds, []);
      assert.deepEqual(receipt.warnings, []);
      const expected = { x: 120, y: 90, width: 50, height: 30 };
      assert.deepEqual(
        (await session.getElement("p0:o0")).item?.bounds,
        expected,
      );
      await session.undo();
      assert.deepEqual((await session.getElement("p0:o0")).item?.bounds, {
        x: 80,
        y: 80,
        width: 40,
        height: 40,
      });
      assert.equal(session.state.canUndo, false);
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
      } finally {
        await reopened.end();
      }
    } finally {
      await end();
    }
  });
});
