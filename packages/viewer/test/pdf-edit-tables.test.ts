import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import { loadPdfEditEngine } from "../src/edit/pdf/provider.js";
import { pdfOperationSchemas } from "../src/edit/pdf/schemas.js";
import { checkOperations } from "../src/edit/operations.js";
import { assertSupportedSchema } from "../src/edit/schema.js";
import type {
  InsertTableOperation,
  PageRect,
  PdfElement,
  PdfOperation,
} from "../src/index.js";
import { defaultResourceLimits } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";

const op = <T extends PdfOperation>(operation: T): T => operation;

const ROWS = [
  ["Item", "Qty", "Price"],
  ["Apples", "3", "1.20"],
  ["Pears", "12", "0.95"],
];

function table(
  fields: Partial<InsertTableOperation> = {},
): InsertTableOperation {
  return {
    op: "insertTable",
    pageIndex: 0,
    at: { x: 72, y: 100 },
    width: 300,
    rows: ROWS,
    ...fields,
  };
}

function near(actual: PageRect, expected: PageRect, slack: number): void {
  for (const key of ["x", "y", "width", "height"] as const)
    assert.ok(
      Math.abs(actual[key] - expected[key]) <= slack,
      `${key}: ${actual[key]} vs ${expected[key]} in ${JSON.stringify(actual)}`,
    );
}

/** Positions of `words` in `text`, which must increase for reading order. */
function order(text: string, words: readonly string[]): number[] {
  return words.map((word) => {
    const at = text.indexOf(word);
    assert.ok(at >= 0, `${word} in ${JSON.stringify(text)}`);
    return at;
  });
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

describe("insertTable", () => {
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

  it("draws the grid and cells, lists one element and reads back in order", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const operation = table({
        columnWidths: [2, 1, 1],
        style: { headerFill: "#eeeeee", fontSize: 12 },
      });
      assert.deepEqual(checkOperations([operation], pdfOperationSchemas), []);
      assert.deepEqual(model.validate([operation]), []);
      const change = model.apply([operation]);
      assert.deepEqual(change, {
        createdIds: ["p0:n1.0.0"],
        removedIds: [],
        changedPages: [0],
        pageCount: 2,
        warnings: [],
      });
      const elements = model.getElements({ pageIndex: 0 });
      assert.deepEqual(
        elements.map((element) => [element.kind, element.id]),
        [
          ["text", "p0:o0"],
          ["table", "p0:n1.0.0"],
        ],
      );
      const grid = elements[1]!;
      assert.deepEqual(grid.table, { rows: ROWS });
      assert.deepEqual(grid.shapeStyle, {
        stroke: { color: "#000000", width: 0.75 },
        fill: { color: "#eeeeee" },
      });
      assert.equal(
        grid.text,
        "Item\tQty\tPrice\nApples\t3\t1.20\nPears\t12\t0.95",
      );
      // Three rows of one 12 pt line each, 4 pt padding, 0.75 pt stroke.
      const rowHeight = 12 * 1.2 + 8;
      near(
        grid.bounds,
        { x: 72, y: 100, width: 300, height: rowHeight * 3 },
        1.6,
      );
      assert.deepEqual(grid.operations, [
        "setTableCell",
        "moveElement",
        "deleteElement",
      ]);

      const saved = model.materialize();
      const text = await extractPageText(saved, 0);
      const positions = order(text, ROWS.flat());
      assert.deepEqual(
        positions,
        [...positions].sort((a, b) => a - b),
      );
      // One text, the header fill, the grid and nine cells.
      assert.equal(await objectCount(saved, 0), 12);
    } finally {
      model.dispose();
    }
  });

  it("wraps long cells, grows rows to the tallest cell and sits upright on rotated pages", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([
        table({
          width: 200,
          rows: [
            ["Short", "A sentence long enough to wrap onto several lines"],
            ["x", ""],
          ],
        }),
        table({
          pageIndex: 1,
          at: { x: 20, y: 30 },
          width: 120,
          rows: [["Turned", "Table"]],
        }),
      ]);
      const wrapped = model.getElement("p0:n1.0.0")!;
      // Two rows: one of three or four lines, one of a single line.
      assert.ok(wrapped.bounds.height > 12 * 4, JSON.stringify(wrapped.bounds));
      assert.ok(wrapped.bounds.height < 12 * 8, JSON.stringify(wrapped.bounds));
      near(
        { ...wrapped.bounds, height: 0 },
        { x: 72, y: 100, width: 200, height: 0 },
        1.6,
      );
      const turned = model.getElement("p1:n1.1.0")!;
      assert.equal(turned.pageIndex, 1);
      assert.equal(turned.rotation, undefined, "upright on the displayed page");
      near(
        { ...turned.bounds, height: 0 },
        { x: 20, y: 30, width: 120, height: 0 },
        1.6,
      );
      assert.equal(
        await extractPageText(model.materialize(), 1),
        "Turned Table",
      );
    } finally {
      model.dispose();
    }
  });

  it("rejects what it cannot draw and changes nothing", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const shape = checkOperations(
        [
          table({ rows: Array.from({ length: 101 }, () => ["x"]) }),
          table({ rows: [Array.from({ length: 21 }, () => "x")] }),
          table({ rows: [] }),
          { ...table(), style: { borderWidth: 50 } },
          op({ op: "setTableCell", target: "t", row: -1, column: 0, text: "" }),
        ],
        pdfOperationSchemas,
      ).map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(shape, [
        "0/rows:max-items",
        "1/rows/0:max-items",
        "2/rows:min-items",
        "3/style/borderWidth:maximum",
        "4/row:minimum",
      ]);
      const engine = model
        .validate([
          table({ pageIndex: 5 }),
          table({ rows: [["a", "b"], ["c"]] }),
          table({ columnWidths: [1, 1] }),
          table({ at: { x: 400, y: 100 } }),
          table({ width: 20 }),
          table({ rows: [["Привіт"]] }),
          table({ style: { fontFamily: "Comic Sans" } }),
          op({
            op: "setTableCell",
            target: "p0:o0",
            row: 0,
            column: 0,
            text: "x",
          }),
        ])
        .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(engine, [
        "0/pageIndex:unknown-target",
        "1/rows/1:range",
        "2/columnWidths:range",
        "3/at:range",
        "4/width:range",
        "5/rows:font-unavailable",
        "6/style/fontFamily:unknown-font",
        "7/target:unsupported-target",
      ]);
      assert.deepEqual(model.materialize(), original);
    } finally {
      model.dispose();
    }
  });

  it("warns when the table runs past the page", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      const change = model.apply([
        table({
          at: { x: 72, y: 760 },
          rows: [["one"], ["two"], ["three"]],
        }),
      ]);
      assert.equal(change.warnings[0]?.code, "fidelity-degraded");
      assert.equal(change.warnings[0]?.details?.elementId, "p0:n1.0.0");
    } finally {
      model.dispose();
    }
  });
});

describe("table edits", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;
  let original: Uint8Array;

  before(async () => {
    pdfium = await fixturePdfium();
    original = await buildPdf(["Existing"]);
  });

  it("rebuilds the table in place for setTableCell and moveElement, and deletes it whole", async () => {
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([
        table(),
        table({ at: { x: 72, y: 300 }, rows: [["Second"]] }),
      ]);
      const ids = () => model.getElements({ pageIndex: 0 }).map((e) => e.id);
      assert.deepEqual(ids(), ["p0:o0", "p0:n1.0.0", "p0:n1.1.0"]);
      const before = model.getElement("p0:n1.0.0")!;

      const changed = model.apply([
        op({
          op: "setTableCell",
          target: "p0:n1.0.0",
          row: 1,
          column: 0,
          text: "Oranges, which is a much longer item name that wraps",
        }),
      ]);
      assert.deepEqual(changed, {
        createdIds: [],
        removedIds: [],
        changedPages: [0],
        pageCount: 1,
        warnings: [],
      });
      assert.deepEqual(ids(), ["p0:o0", "p0:n1.0.0", "p0:n1.1.0"]);
      const after = model.getElement("p0:n1.0.0")!;
      assert.equal(
        after.table?.rows[1]?.[0],
        "Oranges, which is a much longer item name that wraps",
      );
      assert.equal(after.table?.rows[0]?.[0], "Item");
      assert.ok(Math.abs(after.bounds.x - before.bounds.x) < 0.5);
      assert.ok(Math.abs(after.bounds.y - before.bounds.y) < 0.5);
      assert.ok(after.bounds.height > before.bounds.height, "the row grew");

      model.apply([
        op({
          op: "setTableCell",
          target: "p0:n1.0.0",
          row: 2,
          column: 2,
          text: "",
        }),
      ]);
      assert.equal(model.getElement("p0:n1.0.0")!.table?.rows[2]?.[2], "");
      assert.equal(
        (await extractPageText(model.materialize(), 0)).includes("0.95"),
        false,
      );

      model.apply([
        op({ op: "moveElement", target: "p0:n1.0.0", to: { x: 100, y: 400 } }),
      ]);
      const moved = model.getElement("p0:n1.0.0")!;
      near(
        { ...moved.bounds, width: 0, height: 0 },
        { x: 100, y: 400, width: 0, height: 0 },
        0.5,
      );
      assert.deepEqual(moved.table, {
        rows: [
          ROWS[0],
          ["Oranges, which is a much longer item name that wraps", "3", "1.20"],
          ["Pears", "12", ""],
        ],
      });
      assert.deepEqual(ids(), ["p0:o0", "p0:n1.0.0", "p0:n1.1.0"]);

      const refused = model
        .validate([
          op({
            op: "resizeElement",
            target: "p0:n1.0.0",
            rect: { x: 0, y: 0, width: 10, height: 10 },
          }),
          op({
            op: "setTableCell",
            target: "p0:n1.0.0",
            row: 3,
            column: 0,
            text: "x",
          }),
          op({
            op: "setTableCell",
            target: "p0:n1.0.0",
            row: 0,
            column: 3,
            text: "x",
          }),
          op({
            op: "setTableCell",
            target: "p0:n1.0.0",
            row: 0,
            column: 0,
            text: "日本",
          }),
        ])
        .map((issue) => `${issue.operationIndex}${issue.path}:${issue.code}`);
      assert.deepEqual(refused, [
        "0/target:unsupported-target",
        "1/row:range",
        "2/column:range",
        "3/text:font-unavailable",
      ]);

      model.apply([op({ op: "deleteElement", target: "p0:n1.0.0" })]);
      assert.deepEqual(ids(), ["p0:o0", "p0:n1.1.0"]);
      assert.equal(await objectCount(model.materialize(), 0), 3);
    } finally {
      model.dispose();
    }
  });

  it("is one table with its rows after save and reopen", async () => {
    const first = new PdfEditDocument(pdfium, original);
    let saved: Uint8Array;
    try {
      first.apply([table({ style: { headerFill: "#ffcc00" } })]);
      saved = first.materialize();
    } finally {
      first.dispose();
    }
    const reopened = new PdfEditDocument(pdfium, saved);
    try {
      const elements = reopened.getElements({ pageIndex: 0 });
      assert.deepEqual(
        elements.map((element) => [element.kind, element.id]),
        [
          ["text", "p0:o0"],
          ["table", "p0:n1.0.0"],
        ],
      );
      const grid = elements[1] as PdfElement;
      assert.deepEqual(grid.table, { rows: ROWS });
      assert.equal(grid.shapeStyle?.fill?.color, "#ffcc00");
      reopened.apply([
        op({
          op: "setTableCell",
          target: "p0:n1.0.0",
          row: 0,
          column: 0,
          text: "Product",
        }),
      ]);
      assert.equal(
        reopened.getElement("p0:n1.0.0")?.table?.rows[0]?.[0],
        "Product",
      );
      const text = await extractPageText(reopened.materialize(), 0);
      assert.ok(text.includes("Product") && !text.includes("Item"), text);
    } finally {
      reopened.dispose();
    }
  });

  it("treats table members without a head mark as plain objects", async () => {
    const headless = await buildPdf([
      {
        texts: [
          { text: "Plain", x: 72, y: 700 },
          {
            text: "Orphan",
            x: 72,
            y: 650,
            mark: { kind: "table", id: "p0:n9.0.0" },
          },
        ],
      },
    ]);
    const model = new PdfEditDocument(pdfium, headless);
    try {
      assert.deepEqual(
        model.getElements({ pageIndex: 0 }).map((e) => [e.kind, e.id]),
        [
          ["text", "p0:o0"],
          ["text", "p0:o1"],
        ],
      );
    } finally {
      model.dispose();
    }
  });

  it("replays tables deterministically", async () => {
    const batch = [
      table({ style: { headerFill: "#dddddd" } }),
      op({
        op: "setTableCell",
        target: "p0:n1.0.0",
        row: 1,
        column: 1,
        text: "4",
      }),
      op({ op: "moveElement", target: "p0:n1.0.0", by: { dx: 10, dy: 20 } }),
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

describe("tables through the worker", () => {
  const ttf = new Uint8Array(
    readFileSync(
      new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
    ),
  );
  const FALLBACK_URL = "https://fonts.test/noto.ttf";
  const signal = new AbortController().signal;

  it("fetches the fallback font for cells the standard fonts cannot draw", async () => {
    const fetched: string[] = [];
    const pair = loopbackWorker(
      createPdfEditHandler({
        loadPdfium: () => fixturePdfium(),
        decodeImage: async () => {
          throw new Error("no images");
        },
        fetchBytes: async (url) => {
          fetched.push(url);
          if (url === FALLBACK_URL) return ttf;
          throw new Error(`No font at ${url}`);
        },
      }),
    );
    const engine = await loadPdfEditEngine(
      await buildPdf(["Existing"]),
      { format: "pdf", limits: defaultResourceLimits, signal },
      { createWorker: () => pair.worker, fallbackFontUrl: FALLBACK_URL },
    );
    try {
      const change = await engine.apply(
        [table({ rows: [["Товар", "Кількість"]] })],
        signal,
      );
      assert.equal(change.warnings[0]?.code, "font-substitution");
      assert.deepEqual(fetched, [FALLBACK_URL]);
      await engine.apply(
        [
          op({
            op: "setTableCell",
            target: "p0:n1.0.0",
            row: 0,
            column: 1,
            text: "Ціна",
          }),
        ],
        signal,
      );
      assert.equal(
        await extractPageText(await engine.materialize(signal), 0),
        "Existing\r\nТовар Ціна",
      );
    } finally {
      await engine.dispose();
    }
  });
});
