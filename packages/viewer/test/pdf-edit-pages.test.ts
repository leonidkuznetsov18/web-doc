import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PdfOperation } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";
import { sharedFormPdf } from "./fixtures/pdf-forms.js";
import { pdfSession } from "./fixtures/pdf-session.js";

const op = <T extends PdfOperation>(operation: T): T => operation;

for (const [name, operations] of [
  ["page insertion", [op({ op: "insertPage", index: 0 })]],
  [
    "page reorder",
    [
      op({ op: "insertPage", index: 1 }),
      op({ op: "movePage", from: 0, to: 1 }),
    ],
  ],
] as const) {
  it(`keeps ordinary and Form targets through ${name}, checkpoint Undo/Redo and save`, async () => {
    const { session, end } = await pdfSession(sharedFormPdf());
    const control = "p0:o0";
    const first = "p0:o1/2/1";
    const second = "p0:o1/3/1";
    try {
      for (const operation of operations) await session.apply([operation]);
      await session.createCheckpoint("after page structure change");
      await session.apply([
        { op: "setTextStyle", target: control, style: { color: "#00ff00" } },
        { op: "setTextStyle", target: first, style: { color: "#ff0000" } },
      ]);
      await session.undo();
      await session.redo();
      const edited = (await session.getElements({ pageIndex: 1 })).items;
      assert.equal(
        edited.find((element) => element.id === control)?.textStyle?.color,
        "#00ff00",
      );
      assert.equal(
        edited.find((element) => element.id === first)?.textStyle?.color,
        "#ff0000",
      );
      assert.equal(
        edited.find((element) => element.id === second)?.textStyle?.color,
        "#000000",
      );
      assert.ok(
        Math.abs(
          edited.find((element) => element.id === first)!.bounds.x - 72.784,
        ) < 0.01,
      );
      assert.ok(
        Math.abs(
          edited.find((element) => element.id === second)!.bounds.x - 312.784,
        ) < 0.01,
      );

      await session.replaceText({ target: control, text: "Updated control" });
      await session.replaceText({ target: first, text: "Updated inner" });
      const reopened = new PdfEditDocument(
        await fixturePdfium(),
        (await session.save()).bytes,
      );
      try {
        assert.equal(reopened.getElement(control)?.pageIndex, 1);
        assert.equal(reopened.getElement(control)?.text, "Updated control");
        assert.equal(reopened.getElement(first)?.pageIndex, 1);
        assert.equal(reopened.getElement(first)?.text?.trim(), "Updated inner");
        assert.equal(reopened.getElement(second)?.text?.trim(), "Shared inner");
        const incremental = new PdfEditDocument(
          await fixturePdfium(),
          (await session.save({ mode: "incremental" })).bytes,
        );
        try {
          assert.equal(incremental.getElement(control)?.pageIndex, 1);
          assert.equal(
            incremental.getElement(control)?.text,
            "Updated control",
          );
          assert.equal(
            incremental.getElement(first)?.text?.trim(),
            "Updated inner",
          );
          assert.equal(
            incremental.getElement(second)?.text?.trim(),
            "Shared inner",
          );
        } finally {
          incremental.dispose();
        }
      } finally {
        reopened.dispose();
      }
    } finally {
      await end();
    }
  });
}

it("keeps native shape and newly created page targets after checkpoint restoration", async () => {
  const { session, end } = await pdfSession(
    await buildPdf([
      {
        text: "Original",
        rect: { x: 100, y: 300, width: 120, height: 40, fill: [0, 0, 255] },
      },
      "Second",
    ]),
  );
  try {
    await session.insertPage({ index: 0 });
    const created = await session.insertTextBox({
      pageIndex: 0,
      rect: { x: 40, y: 40, width: 200, height: 40 },
      text: "New page",
    });
    const target = created.createdIds[0]!;
    const before = (await session.getElements({ pageIndex: 1 })).items.find(
      (element) => element.id === "p0:o1",
    )!.bounds;
    await session.createCheckpoint("with a new page element");
    await session.apply([
      { op: "moveElement", target: "p0:o1", by: { dx: 9, dy: 13 } },
      { op: "replaceText", target, text: "New page edited" },
    ]);
    await session.undo();
    await session.redo();
    const edited = (await session.getElements({})).items;
    assert.equal(edited.find((element) => element.id === target)?.pageIndex, 0);
    assert.equal(
      edited.find((element) => element.id === target)?.text,
      "New page edited",
    );
    assert.equal(
      edited.find((element) => element.id === "p0:o1")?.pageIndex,
      1,
    );
    assert.equal(
      edited.find((element) => element.id === "p0:o1")?.bounds.x,
      before.x + 9,
    );
    assert.equal(
      edited.find((element) => element.id === "p0:o1")?.bounds.y,
      before.y + 13,
    );
    const reopened = new PdfEditDocument(
      await fixturePdfium(),
      (await session.save()).bytes,
    );
    try {
      assert.equal(reopened.getElement(target)?.text, "New page edited");
      assert.equal(reopened.getElement("p0:o1")?.pageIndex, 1);
      assert.equal(reopened.getElement("p1:o0")?.pageIndex, 2);
      // A fresh session starts its state counter at one. Its new page must
      // not reuse the previously saved q1.0 logical page's identity.
      reopened.apply([{ op: "insertPage", index: 0 }]);
      const added = reopened.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 40, y: 40, width: 200, height: 40 },
          text: "Another new page",
        },
      ]);
      assert.notEqual(added.createdIds[0], target);
      assert.equal(reopened.getElement(target)?.text, "New page edited");
      assert.equal(reopened.getElement(target)?.pageIndex, 1);
      assert.equal(
        reopened.getElement(added.createdIds[0]!)?.text,
        "Another new page",
      );
    } finally {
      reopened.dispose();
    }
  } finally {
    await end();
  }
});

it("rejects malformed or stale page metadata without rebinding native targets", async () => {
  const pdfium = await fixturePdfium();
  const original = await buildPdf(["First", "Second"]);
  const document = pdfium.openDocument(original);
  const numbers = [0, 1].map((index) =>
    pdfium.lib.EPDFDoc_GetPageObjectNumberByIndex(document.handle, index),
  );
  document.close();
  const entries = numbers.map((object, index) => ({
    object,
    key: `p${index}`,
  }));
  const invalid = [
    "invalid json",
    JSON.stringify({ version: 2, pages: entries }),
    JSON.stringify({ version: 1, pages: entries.slice(0, 1) }),
    JSON.stringify({ version: 1, pages: [entries[0], entries[0]] }),
    JSON.stringify({
      version: 1,
      pages: [entries[0], { object: numbers[1], key: "p0" }],
    }),
    JSON.stringify({
      version: 1,
      pages: [entries[0], { object: 999999, key: "p1" }],
    }),
    JSON.stringify({
      version: 1,
      pages: [entries[0], { object: numbers[1], key: "foreign" }],
    }),
  ];
  for (const raw of invalid) {
    const native = pdfium.openDocument(original);
    const pointer = pdfium.writeWideString(raw);
    try {
      assert.ok(
        pdfium.lib.EPDF_SetMetaText(native.handle, "WebDocPageKeys", pointer),
      );
      const model = new PdfEditDocument(pdfium, native.save("full"));
      try {
        assert.equal(model.getElement("p0:o0")?.text, "First");
        assert.equal(model.getElement("p1:o0")?.text, "Second");
      } finally {
        model.dispose();
      }
    } finally {
      pdfium.free(pointer);
      native.close();
    }
  }
});

/** Page sizes and rotations as PDFium reads them back from saved bytes. */
async function pageShapes(bytes: Uint8Array) {
  const pdfium = await fixturePdfium();
  const { lib } = pdfium;
  const document = pdfium.openDocument(bytes);
  try {
    const count = lib.FPDF_GetPageCount(document.handle);
    return Array.from({ length: count }, (_, index) => {
      const page = lib.FPDF_LoadPage(document.handle, index);
      try {
        return [
          lib.FPDF_GetPageWidthF(page),
          lib.FPDF_GetPageHeightF(page),
          lib.FPDFPage_GetRotation(page) * 90,
        ];
      } finally {
        lib.FPDF_ClosePage(page);
      }
    });
  } finally {
    document.close();
  }
}

describe("page operations", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;
  let original: Uint8Array;

  before(async () => {
    pdfium = await fixturePdfium();
    original = await buildPdf([
      "One",
      { text: "Two", image: { x: 300, y: 500, width: 160, height: 80 } },
      { text: "Three", width: 400, height: 500 },
    ]);
  });

  it("inserts, moves, rotates and deletes pages while ids follow their content", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const inserted = model.apply([
        op({ op: "insertPage", index: 1, size: { width: 300, height: 200 } }),
        op({ op: "insertPage", index: 4 }),
      ]);
      assert.equal(inserted.pageCount, 5);
      assert.deepEqual(inserted.changedPages, [1, 2, 3, 4]);
      assert.deepEqual(await pageShapes(model.materialize()), [
        [612, 792, 0],
        [300, 200, 0],
        [612, 792, 0],
        [400, 500, 0],
        [400, 500, 0], // defaults to the page before it
      ]);
      // "Two" kept its id and now sits on page 2.
      const two = model.getElement("p1:o0")!;
      assert.equal(two.pageIndex, 2);
      assert.equal(two.text, "Two");

      const moved = model.apply([op({ op: "movePage", from: 3, to: 0 })]);
      assert.deepEqual(moved.changedPages, [0, 1, 2, 3]);
      assert.equal(await extractPageText(model.materialize(), 0), "Three");
      assert.equal(model.getElement("p2:o0")!.pageIndex, 0);
      assert.equal(model.getElement("p1:o0")!.pageIndex, 3);

      model.apply([op({ op: "rotatePage", pageIndex: 0, rotation: 90 })]);
      assert.deepEqual(
        (await pageShapes(model.materialize()))[0],
        [500, 400, 90],
      );
      // Bounds follow the rotation: the text is listed in the turned frame.
      const turned = model.getElement("p2:o0")!;
      assert.ok(turned.bounds.x > 300, JSON.stringify(turned.bounds));
      assert.equal(turned.rotation, 90);

      const deleted = model.apply([op({ op: "deletePage", pageIndex: 2 })]);
      assert.equal(deleted.pageCount, 4);
      assert.deepEqual(deleted.changedPages, [2, 3]);
      assert.deepEqual(
        model.getElements({}).map((element) => [element.id, element.pageIndex]),
        [
          ["p2:o0", 0],
          ["p0:o0", 1],
          ["p1:o0", 2],
          ["p1:o1", 2],
        ],
      );
      assert.equal(
        model.getElements({ pageIndex: 3 }).length,
        0,
        "the appended blank page is last",
      );
    } finally {
      model.dispose();
    }
  });

  it("keeps elements editable on moved and new pages", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([op({ op: "insertPage", index: 0 })]);
      const change = model.apply([
        op({
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 50, y: 50, width: 200, height: 40 },
          text: "On the new page",
        }),
        op({ op: "moveElement", target: "p0:o0", by: { dx: 5, dy: 5 } }),
      ]);
      assert.deepEqual(change.createdIds, ["q1.0:n2.0.0"]);
      assert.deepEqual(change.changedPages, [0, 1]);
      assert.equal(
        await extractPageText(model.materialize(), 0),
        "On the new page",
      );
      model.apply([op({ op: "movePage", from: 0, to: 3 })]);
      assert.equal(model.getElement("q1.0:n2.0.0")!.pageIndex, 3);
      model.apply([
        op({
          op: "replaceText",
          target: "q1.0:n2.0.0",
          text: "Still editable",
        }),
      ]);
      assert.equal(
        await extractPageText(model.materialize(), 3),
        "Still editable",
      );
    } finally {
      model.dispose();
    }
  });

  it("validates indexes and refuses to delete the last page", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const issues = model
        .validate([
          op({ op: "insertPage", index: 4 }),
          op({ op: "deletePage", pageIndex: 3 }),
          op({ op: "movePage", from: 3, to: 0 }),
          op({ op: "movePage", from: 0, to: 3 }),
          op({ op: "rotatePage", pageIndex: 7, rotation: 180 }),
        ])
        .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(issues, [
        "0/index:range",
        "1/pageIndex:unknown-target",
        "2/from:unknown-target",
        "3/to:range",
        "4/pageIndex:unknown-target",
      ]);
      const single = new PdfEditDocument(pdfium, await buildPdf(["Only"]));
      try {
        assert.equal(
          single.validate([op({ op: "deletePage", pageIndex: 0 })])[0]?.code,
          "last-page",
        );
      } finally {
        single.dispose();
      }
    } finally {
      model.dispose();
    }
  });

  // ACTION-888: "rotate" from a host that cannot see a page's angle. A file
  // saved with page 2 at 90° must turn to 180°, each page on its own, and a
  // replay of the history (undo, redo) turns from the same angle again.
  it("turns a page from the angle it has with `by`", async () => {
    const turned = await buildPdf([
      "Upright",
      { text: "Saved turned", rotation: 1 },
    ]);
    const model = new PdfEditDocument(pdfium, turned);
    try {
      model.apply([op({ op: "rotatePage", pageIndex: 1, by: 90 })]);
      assert.deepEqual(
        (await pageShapes(model.materialize())).map((shape) => shape[2]),
        [0, 180],
      );
      model.apply([op({ op: "rotatePage", pageIndex: 0, by: 90 })]);
      model.apply([op({ op: "rotatePage", pageIndex: 1, by: 270 })]);
      assert.deepEqual(
        (await pageShapes(model.materialize())).map((shape) => shape[2]),
        [90, 90],
      );
      // Undo of the last turn, then redo: the same angles come back.
      model.restore([
        [op({ op: "rotatePage", pageIndex: 1, by: 90 })],
        [op({ op: "rotatePage", pageIndex: 0, by: 90 })],
      ]);
      assert.deepEqual(
        (await pageShapes(model.materialize())).map((shape) => shape[2]),
        [90, 180],
      );
      assert.deepEqual(
        model
          .validate([
            op({ op: "rotatePage", pageIndex: 0 } as unknown as PdfOperation),
            op({
              op: "rotatePage",
              pageIndex: 0,
              rotation: 90,
              by: 90,
            } as unknown as PdfOperation),
          ])
          .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`),
        ["0:one-of", "1:one-of"],
      );
      // A field spread in as undefined is absent, as a structured clone keeps
      // the key.
      assert.deepEqual(
        model.validate([
          op({
            op: "rotatePage",
            pageIndex: 0,
            rotation: 90,
            by: undefined,
          } as unknown as PdfOperation),
          op({
            op: "rotatePage",
            pageIndex: 0,
            rotation: undefined,
            by: 90,
          } as unknown as PdfOperation),
        ]),
        [],
      );
    } finally {
      model.dispose();
    }
  });

  it("replays page structure deterministically", async () => {
    const batches = [
      [
        op({ op: "insertPage", index: 0 }),
        op({ op: "rotatePage", pageIndex: 2, rotation: 270 }),
      ],
      [
        op({ op: "movePage", from: 3, to: 1 }),
        op({ op: "deletePage", pageIndex: 0 }),
      ],
    ];
    const first = new PdfEditDocument(pdfium, original);
    const second = new PdfEditDocument(pdfium, original);
    try {
      for (const batch of batches) first.apply(batch);
      second.restore(batches);
      assert.deepEqual(second.materialize(), first.materialize());
      // [new, One, Two↻, Three] → move Three to 1 → delete the new page.
      assert.deepEqual(await pageShapes(first.materialize()), [
        [400, 500, 0],
        [612, 792, 0],
        [792, 612, 270],
      ]);
      first.restore([]);
      assert.deepEqual(first.materialize(), original);
    } finally {
      first.dispose();
      second.dispose();
    }
  });
});
