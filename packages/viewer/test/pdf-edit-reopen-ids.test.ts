import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import type { PdfOperation } from "../src/index.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";

const op = <T extends PdfOperation>(operation: T): T => operation;

const savedBox = op({
  op: "insertTextBox",
  pageIndex: 0,
  rect: { x: 72, y: 200, width: 300, height: 40 },
  text: "Saved before",
});

/** What a second session inserts, one of each kind that mints an id. */
const inserts: readonly PdfOperation[] = [
  op({
    op: "insertTextBox",
    pageIndex: 0,
    rect: { x: 72, y: 300, width: 300, height: 40 },
    text: "Inserted after reopen",
  }),
  op({
    op: "insertShape",
    pageIndex: 0,
    shape: "rectangle",
    rect: { x: 72, y: 400, width: 100, height: 40 },
    fill: { color: "#ff8800" },
  }),
  op({
    op: "insertTable",
    pageIndex: 0,
    at: { x: 72, y: 500 },
    width: 200,
    rows: [["Cell"]],
  }),
];

// ACTION-886: a new session numbers its changes from 1 again, so the first
// object it made took the id an earlier session had saved into the file, and
// selecting or deleting it reached that saved object too.
describe("element ids after a document is saved and reopened", () => {
  let pdfium: Awaited<ReturnType<typeof fixturePdfium>>;
  let saved: Uint8Array;
  let savedId: string;

  before(async () => {
    pdfium = await fixturePdfium();
    const first = new PdfEditDocument(pdfium, await buildPdf(["Existing"]));
    try {
      savedId = first.apply([savedBox]).createdIds[0]!;
      saved = first.materialize();
    } finally {
      first.dispose();
    }
  });

  for (const insert of inserts)
    it(`gives a new ${insert.op} an id no saved object has`, async () => {
      const reopened = new PdfEditDocument(pdfium, saved);
      try {
        assert.equal(reopened.getElement(savedId)?.text, "Saved before");
        const before = new Set(
          reopened.getElements({ pageIndex: 0 }).map((element) => element.id),
        );
        const [created] = reopened.apply([insert]).createdIds;
        assert.ok(created, "an id was created");
        assert.ok(!before.has(created), `${created} is already taken`);
        const ids = reopened
          .getElements({ pageIndex: 0 })
          .map((element) => element.id);
        assert.equal(new Set(ids).size, ids.length, `unique: ${ids}`);
        assert.equal(reopened.getElement(savedId)?.text, "Saved before");

        // Deleting what was just made leaves the saved box alone.
        reopened.apply([op({ op: "deleteElement", target: created })]);
        assert.equal(reopened.getElement(created), undefined);
        assert.equal(reopened.getElement(savedId)?.text, "Saved before");
        assert.match(
          await extractPageText(reopened.materialize(), 0),
          /Saved before/,
        );
      } finally {
        reopened.dispose();
      }
    });

  it("hands out the same ids when the history is replayed", () => {
    const reopened = new PdfEditDocument(pdfium, saved);
    try {
      const first = reopened.apply([inserts[0]!]).createdIds;
      reopened.restore([]);
      reopened.restore([[inserts[0]!]]);
      assert.equal(
        reopened.getElement(first[0]!)?.text,
        "Inserted after reopen",
      );
      assert.equal(reopened.getElement(savedId)?.text, "Saved before");
    } finally {
      reopened.dispose();
    }
  });

  it("keeps ids unique over a second save and reopen", () => {
    const second = new PdfEditDocument(pdfium, saved);
    let twice: Uint8Array;
    try {
      second.apply([inserts[0]!]);
      twice = second.materialize();
    } finally {
      second.dispose();
    }
    const third = new PdfEditDocument(pdfium, twice);
    try {
      const [created] = third.apply([inserts[1]!]).createdIds;
      const ids = third.getElements({ pageIndex: 0 }).map((e) => e.id);
      assert.equal(new Set(ids).size, ids.length, `unique: ${ids}`);
      assert.equal(third.getElement(created!)?.kind, "shape");
    } finally {
      third.dispose();
    }
  });
});
