import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { before, describe, it } from "node:test";
import { inflateSync } from "node:zlib";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
import { parseCmap } from "../src/edit/pdf/engine/fonts.js";
import { createPdfEditHandler } from "../src/edit/pdf/engine/handler.js";
import { loadPdfEditEngine } from "../src/edit/pdf/provider.js";
import { PdfSession } from "../src/edit/pdf/session.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../src/edit/session.js";
import type {
  EditElement,
  EditOperation,
  EditSessionBase,
  OperationIssue,
  PdfEditSession,
  PdfOperation,
} from "../src/index.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import {
  buildPdf,
  extractPageText,
  fixturePdfium,
} from "./fixtures/pdf-builder.js";
import { handPdf, signedPdf } from "./fixtures/signed-pdf.js";
import { tinyJpeg } from "./fixtures/tiny-jpeg.js";

const signal = new AbortController().signal;

async function pageCountOf(bytes: Uint8Array): Promise<number> {
  const pdfium = await fixturePdfium();
  const document = pdfium.openDocument(bytes);
  try {
    return pdfium.lib.FPDF_GetPageCount(document.handle);
  } finally {
    document.close();
  }
}

/** A PDF session over the loopback worker, with a host that only counts pages. */
async function pdfSession(
  original: Uint8Array,
): Promise<{ session: PdfEditSession; end(): Promise<void> }> {
  const pair = loopbackWorker(
    createPdfEditHandler({
      loadPdfium: () => fixturePdfium(),
      fetchBytes: async () => {
        throw new Error("no fonts");
      },
      decodeImage: async () => {
        throw new Error("no images");
      },
    }),
  );
  const engine = await loadPdfEditEngine(
    original,
    { format: "pdf", limits: defaultResourceLimits, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "pdf",
    limits: defaultResourceLimits,
    prepareDocument: async (bytes) => ({ pageCount: await pageCountOf(bytes) }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
  };
  const core = new EditSessionController(
    engine,
    host,
    original,
    await pageCountOf(original),
  );
  return { session: new PdfSession(core), end: () => core.end() };
}

function startsWith(bytes: Uint8Array, prefix: Uint8Array): boolean {
  return (
    bytes.length >= prefix.length &&
    prefix.every((byte, index) => bytes[index] === byte)
  );
}

describe("signed PDFs", () => {
  it("edits a signed file, warns on the first change and keeps the signed bytes", async () => {
    const pdfium = await fixturePdfium();
    const original = signedPdf();
    const model = new PdfEditDocument(pdfium, original);
    try {
      assert.equal(model.signatureCount, 1);
      assert.equal(model.getElements({ pageIndex: 0 })[0]?.text, "Signed");
      const first = model.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 72, y: 72, width: 200, height: 40 },
          text: "Added later",
        },
      ]);
      assert.equal(first.warnings[0]?.code, "fidelity-degraded");
      assert.match(first.warnings[0]!.message, /signature/);
      assert.deepEqual(first.warnings[0]?.details, {
        signatures: 1,
        features: [],
      });
      const second = model.apply([
        { op: "replaceText", target: "p0:n1.0.0", text: "Changed" },
      ]);
      assert.deepEqual(second.warnings, []);

      const saved = model.materialize();
      assert.equal(startsWith(saved, original), true, "incremental update");
      const reopened = new PdfEditDocument(pdfium, saved);
      try {
        assert.equal(reopened.signatureCount, 1);
        assert.equal(reopened.getElement("p0:n1.0.0")?.text, "Changed");
      } finally {
        reopened.dispose();
      }
    } finally {
      model.dispose();
    }
  });
});

describe("PDF sessions", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf(["One", "Two"]);
  });

  it("returns the real receipt for a dry run and changes nothing", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const fields = {
        pageIndex: 0,
        rect: { x: 72, y: 72, width: 200, height: 40 },
        text: "Maybe",
      };
      const dry = await session.insertTextBox(fields, { dryRun: true });
      assert.deepEqual(dry, {
        sessionId: session.sessionId,
        revision: 0,
        dryRun: true,
        operationCount: 1,
        createdIds: ["p0:n1.0.0"],
        removedIds: [],
        changedPages: [0],
        pageCount: 2,
        warnings: [],
      });
      assert.equal(session.state.revision, 0);
      assert.equal(session.state.dirty, false);
      assert.equal(
        (await session.getElements({ kinds: ["textBox"] })).items.length,
        0,
      );
      assert.deepEqual((await session.save()).bytes, original);

      const real = await session.insertTextBox(fields);
      assert.deepEqual(real, { ...dry, revision: 1, dryRun: false });
      assert.equal(session.state.revision, 1);
      assert.equal((await session.getElement("p0:n1.0.0")).item?.text, "Maybe");
    } finally {
      await end();
    }
  });

  it("produces identical bytes from two sessions with the same history", async () => {
    const a = await pdfSession(original);
    const b = await pdfSession(original);
    try {
      for (const session of [a.session, b.session]) {
        await session.insertTextBox({
          pageIndex: 0,
          rect: { x: 72, y: 72, width: 200, height: 40 },
          text: "Same",
        });
        await session.insertShape({
          pageIndex: 1,
          shape: "ellipse",
          rect: { x: 10, y: 10, width: 50, height: 30 },
          fill: { color: "#336699" },
        });
        await session.insertTable({
          pageIndex: 0,
          at: { x: 72, y: 200 },
          width: 200,
          rows: [["a", "b"]],
        });
        await session.rotatePage({ pageIndex: 1, rotation: 90 });
        await session.undo();
      }
      const bytesA = (await a.session.save()).bytes;
      const bytesB = (await b.session.save()).bytes;
      assert.deepEqual(bytesA, bytesB);
      assert.ok(bytesA.length > original.length);
    } finally {
      await a.end();
      await b.end();
    }
  });

  it("serves a client that knows only the schemas, with typed issues for bad input", async () => {
    const { session, end } = await pdfSession(original);
    try {
      // Everything crosses as JSON: binary data is base64, nothing is typed.
      const client = session as unknown as EditSessionBase<
        EditOperation,
        EditElement
      >;
      const names = Object.keys(client.schemas.operations).sort();
      assert.equal(names.length, 15);
      const run = async (operation: Record<string, unknown>) => {
        assert.ok(names.includes(operation.op as string), String(operation.op));
        const json = JSON.parse(JSON.stringify(operation)) as EditOperation;
        return client.apply([json]);
      };
      const jpeg = Buffer.from(tinyJpeg()).toString("base64");
      const box = { x: 72, y: 72, width: 200, height: 40 };

      const text = await run({
        op: "insertTextBox",
        pageIndex: 0,
        rect: box,
        text: "Hi",
      });
      const [textId] = text.createdIds;
      await run({ op: "replaceText", target: textId, text: "Hello" });
      await run({
        op: "setTextStyle",
        target: textId,
        style: { color: "#ff0000" },
      });
      const image = await run({
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 300, y: 72, width: 64, height: 32 },
        data: jpeg,
        mimeType: "image/jpeg",
      });
      const shape = await run({
        op: "insertShape",
        pageIndex: 0,
        shape: "rectangle",
        rect: { x: 72, y: 150, width: 100, height: 50 },
        stroke: { color: "#000000", width: 1 },
      });
      const [shapeId] = shape.createdIds;
      await run({
        op: "setShapeStyle",
        target: shapeId,
        fill: { color: "#00ff00" },
      });
      const table = await run({
        op: "insertTable",
        pageIndex: 0,
        at: { x: 72, y: 250 },
        width: 200,
        rows: [["k", "v"]],
      });
      await run({
        op: "setTableCell",
        target: table.createdIds[0],
        row: 0,
        column: 1,
        text: "w",
      });
      await run({ op: "moveElement", target: textId, by: { dx: 5, dy: 5 } });
      await run({
        op: "resizeElement",
        target: shapeId,
        rect: { x: 72, y: 150, width: 120, height: 60 },
      });
      await run({ op: "deleteElement", target: image.createdIds[0] });
      await run({ op: "insertPage", index: 2 });
      await run({ op: "rotatePage", pageIndex: 2, rotation: 90 });
      await run({ op: "movePage", from: 2, to: 0 });
      const last = await run({ op: "deletePage", pageIndex: 0 });
      assert.equal(last.revision, 15);
      assert.equal(last.pageCount, 2);
      const kinds = (await client.getElements({ pageIndex: 0 })).items.map(
        (element) => element.kind,
      );
      assert.deepEqual(kinds, ["text", "textBox", "shape", "table"]);

      const issuesOf = async (operation: Record<string, unknown>) => {
        try {
          await run(operation);
        } catch (error) {
          assert.ok(error instanceof ViewerError);
          assert.equal(error.code, "invalid-operation");
          return error.details?.issues as readonly OperationIssue[];
        }
        assert.fail("expected a rejection");
      };
      const shape2 = await issuesOf({ op: "insertTextBox", pageIndex: 0 });
      assert.ok(shape2.length > 0);
      for (const issue of shape2) {
        assert.equal(issue.operationIndex, 0);
        assert.equal(typeof issue.path, "string");
        assert.equal(typeof issue.code, "string");
        assert.equal(typeof issue.message, "string");
      }
      assert.ok(shape2.some((issue) => issue.code === "required"));
      assert.deepEqual(
        (await issuesOf({ op: "deleteElement", target: "nope" })).map(
          ({ operationIndex, path, code }) => ({ operationIndex, path, code }),
        ),
        [{ operationIndex: 0, path: "/target", code: "unknown-target" }],
      );
      assert.equal(session.state.revision, 15, "rejections change nothing");
    } finally {
      await end();
    }
  });
});

describe("ids and references (revision 2)", () => {
  let original: Uint8Array;

  before(async () => {
    original = await buildPdf(["One", "Two"]);
  });

  const box = (text: string) => ({
    pageIndex: 0,
    rect: { x: 72, y: 72, width: 200, height: 40 },
    text,
  });

  it("never hands out an undone id again", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const first = await session.insertTextBox(box("first"));
      assert.deepEqual(first.createdIds, ["p0:n1.0.0"]);
      await session.undo();
      const second = await session.insertTextBox(box("second"));
      assert.deepEqual(second.createdIds, ["p0:n2.0.0"]);
      assert.equal((await session.getElement("p0:n1.0.0")).item, undefined);
      assert.equal(
        (await session.getElement("p0:n2.0.0")).item?.text,
        "second",
      );
      // A dry run names the ids the real apply then uses.
      const dry = await session.insertTextBox(box("third"), { dryRun: true });
      const real = await session.insertTextBox(box("third"));
      assert.deepEqual(real.createdIds, dry.createdIds);
      assert.deepEqual(real.createdIds, ["p0:n3.0.0"]);
    } finally {
      await end();
    }
  });

  it("reports what a batch, an undo, a redo and a reset removed", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const inserted = await session.apply([
        { op: "insertTextBox", ...box("a") },
        {
          op: "insertShape",
          pageIndex: 0,
          shape: "rectangle",
          rect: { x: 10, y: 10, width: 20, height: 20 },
          fill: { color: "#000000" },
        },
      ]);
      const [boxId, shapeId] = inserted.createdIds as [string, string];
      assert.deepEqual(inserted.removedIds, []);
      const deleted = await session.deleteElement({ target: shapeId });
      assert.deepEqual(deleted.removedIds, [shapeId]);
      const undone = await session.undo();
      assert.deepEqual(undone.removedIds, []);
      assert.equal((await session.getElement(shapeId)).item?.kind, "shape");
      const redone = await session.redo();
      assert.deepEqual(redone.removedIds, [shapeId]);
      const page = await session.deletePage({ pageIndex: 0 });
      assert.deepEqual(page.removedIds, ["p0:o0", boxId]);
      assert.equal(session.state.pageCount, 1);
      const undonePage = await session.undo();
      assert.deepEqual(undonePage.removedIds, []);
      const reset = await session.reset();
      assert.deepEqual(reset.removedIds, [boxId, shapeId]);
    } finally {
      await end();
    }
  });

  it("resolves $n references while applying and rejects the batch when one misses", async () => {
    const { session, end } = await pdfSession(original);
    try {
      const receipt = await session.apply([
        {
          op: "insertTable",
          pageIndex: 0,
          at: { x: 72, y: 200 },
          width: 200,
          rows: [["a", "b"]],
        },
        { op: "setTableCell", target: "$0", row: 0, column: 1, text: "c" },
        { op: "moveElement", target: "$0", by: { dx: 5, dy: 5 } },
      ]);
      assert.deepEqual(receipt.createdIds, ["p0:n1.0.0"]);
      assert.deepEqual(
        (await session.getElement("p0:n1.0.0")).item?.table?.rows,
        [["a", "c"]],
      );

      const issuesOf = async (operations: readonly PdfOperation[]) => {
        try {
          await session.apply(operations);
        } catch (error) {
          assert.ok(error instanceof ViewerError, String(error));
          assert.equal(error.code, "invalid-operation");
          return (error.details?.issues as readonly OperationIssue[]).map(
            (issue) => `${issue.operationIndex}${issue.path}:${issue.code}`,
          );
        }
        assert.fail("expected a rejection");
      };
      // Wrong kind of element: found while applying, document unchanged.
      assert.deepEqual(
        await issuesOf([
          { op: "insertTextBox", ...box("not a table") },
          { op: "setTableCell", target: "$0", row: 0, column: 0, text: "x" },
        ]),
        ["1/target:unsupported-target"],
      );
      assert.equal(session.state.revision, 1);
      assert.equal(
        (await session.getElements({ pageIndex: 0 })).items.length,
        2,
        "nothing from the rejected batch remains",
      );
      // References must point backwards, at an operation that creates something.
      assert.deepEqual(
        await issuesOf([
          { op: "setTableCell", target: "$0", row: 0, column: 0, text: "x" },
        ]),
        ["0/target:unknown-target"],
      );
      assert.deepEqual(
        await issuesOf([
          { op: "rotatePage", pageIndex: 0, rotation: 90 },
          { op: "deleteElement", target: "$0" },
        ]),
        ["1/target:unknown-target"],
      );
    } finally {
      await end();
    }
  });
});

describe("checkpoints (revision 2)", () => {
  it("restores the same ids and content from a checkpoint as from the original", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf(["One", "Two"]);
    const batches = [1, 2, 3, 4].map((index) => ({
      stateId: index,
      operations: [
        {
          op: "insertTextBox" as const,
          pageIndex: 0,
          rect: { x: 72, y: 60 * index, width: 200, height: 40 },
          text: `Box ${index}`,
        },
      ],
    }));
    const straight = new PdfEditDocument(pdfium, original);
    const viaCheckpoint = new PdfEditDocument(pdfium, original);
    try {
      straight.restore({ batches: batches.slice(0, 2) });
      const checkpoint = straight.materialize();
      straight.restore({ batches });
      viaCheckpoint.restore({ base: checkpoint, batches: batches.slice(2) });
      const ids = (model: PdfEditDocument) =>
        model
          .getElements({ pageIndex: 0 })
          .map((element) => [element.id, element.text, element.bounds]);
      assert.deepEqual(ids(viaCheckpoint), ids(straight));
      // Bytes are not compared: PDFium's save also writes objects a
      // regenerated page no longer references, and how many of those a
      // document carries depends on where it was opened from.
      assert.equal(
        await extractPageText(viaCheckpoint.materialize("save", "full"), 0),
        await extractPageText(straight.materialize("save", "full"), 0),
      );
    } finally {
      straight.dispose();
      viaCheckpoint.dispose();
    }
  });
});

describe("save modes (revision 2)", () => {
  /** Every FlateDecode stream of a PDF, inflated, so hidden text can be searched. */
  const streamTexts = (bytes: Uint8Array): string => {
    const text = Array.from(bytes, (byte) => String.fromCharCode(byte)).join(
      "",
    );
    const parts: string[] = [];
    const pattern = /\/Length (\d+)[^>]*>>\s*stream\r?\n/g;
    for (const match of text.matchAll(pattern)) {
      const start = match.index! + match[0].length;
      const slice = bytes.subarray(start, start + Number(match[1]));
      try {
        parts.push(new TextDecoder("latin1").decode(inflateSync(slice)));
      } catch {
        parts.push(new TextDecoder("latin1").decode(slice));
      }
    }
    return parts.join("\n");
  };
  const secret = "SECRET-PHRASE-4711";
  const box = {
    op: "insertTextBox" as const,
    pageIndex: 0,
    rect: { x: 72, y: 72, width: 300, height: 40 },
    text: secret,
  };

  it("drops deleted content from a full save and keeps it in an incremental one", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf(["One", "Two"]);
    const model = new PdfEditDocument(pdfium, original);
    try {
      model.apply([box]);
      assert.ok(
        streamTexts(model.materialize("save", "incremental")).includes(secret),
      );
      model.apply([{ op: "deleteElement", target: "p0:n1.0.0" }]);
      const full = model.materialize("save", "full");
      assert.equal(streamTexts(full).includes(secret), false, "compacted away");
      assert.ok(
        streamTexts(model.materialize("save", "incremental")).includes(secret),
        "an incremental save keeps earlier revisions",
      );
      // The compacted file is a complete, readable PDF.
      const reopened = new PdfEditDocument(pdfium, full);
      try {
        assert.equal(reopened.pageCount, 2);
        assert.equal(await extractPageText(full, 0), "One");
        assert.equal(await extractPageText(full, 1), "Two");
      } finally {
        reopened.dispose();
      }
      // Unsigned files save in full mode by default; the viewer still reopens
      // the incremental form.
      assert.deepEqual(model.materialize("save"), full);
      assert.ok(startsWith(model.materialize("show"), original));
    } finally {
      model.dispose();
    }
  });

  it("produces the same full save with and without prior queries", async () => {
    const pdfium = await fixturePdfium();
    const original = await buildPdf(["One", "Two", "Three"]);
    const quiet = new PdfEditDocument(pdfium, original);
    const curious = new PdfEditDocument(pdfium, original);
    try {
      for (const model of [quiet, curious]) model.apply([box]);
      curious.getElements({});
      curious.findText("Two", {});
      curious.elementsAt(2, { x: 10, y: 10 });
      assert.deepEqual(
        curious.materialize("save", "full"),
        quiet.materialize("save", "full"),
      );
    } finally {
      quiet.dispose();
      curious.dispose();
    }
  });

  it("keeps the incremental default for signed files and returns the original unchanged", async () => {
    const pdfium = await fixturePdfium();
    const signed = new PdfEditDocument(pdfium, signedPdf());
    try {
      assert.deepEqual(signed.materialize("save"), signedPdf());
      signed.apply([box]);
      assert.ok(startsWith(signed.materialize("save"), signedPdf()));
      assert.equal(
        startsWith(signed.materialize("save", "full"), signedPdf()),
        false,
      );
    } finally {
      signed.dispose();
    }
  });

  it("passes the save mode from the session to the engine", async () => {
    const original = await buildPdf(["One"]);
    const { session, end } = await pdfSession(original);
    try {
      await session.insertTextBox(box);
      const full = await session.save();
      const incremental = await session.save({ mode: "incremental" });
      assert.ok(startsWith(incremental.bytes, original));
      assert.equal(startsWith(full.bytes, original), false);
      assert.equal(full.stateToken, incremental.stateToken);
    } finally {
      await end();
    }
  });
});

describe("document features and marks (revision 2)", () => {
  it("names DocMDP, tagging and PDF/A in the first-change warning", async () => {
    const pdfium = await fixturePdfium();
    const model = new PdfEditDocument(
      pdfium,
      handPdf({ signed: true, docMdp: true, tagged: true, pdfa: true }),
    );
    try {
      assert.deepEqual(model.features, ["docmdp", "tagged", "pdfa"]);
      assert.equal(model.pageCount, 1);
      const change = model.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 72, y: 72, width: 200, height: 40 },
          text: "x",
        },
      ]);
      assert.equal(change.warnings[0]?.code, "fidelity-degraded");
      assert.deepEqual(change.warnings[0]?.details, {
        signatures: 1,
        features: ["docmdp", "tagged", "pdfa"],
      });
      assert.match(change.warnings[0]!.message, /DocMDP.*tagged.*PDF\/A/);
    } finally {
      model.dispose();
    }
    const plain = new PdfEditDocument(pdfium, await buildPdf(["One"]));
    try {
      assert.deepEqual(plain.features, []);
      assert.equal(plain.signatureCount, 0);
    } finally {
      plain.dispose();
    }
  });

  it("lists a marked group another tool moved or retyped as plain objects", async () => {
    const pdfium = await fixturePdfium();
    const { lib } = pdfium;
    const original = await buildPdf(["One"]);
    const model = new PdfEditDocument(pdfium, original);
    let saved: Uint8Array;
    try {
      model.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 72, y: 72, width: 120, height: 60 },
          text: "wrapped box text that spans lines",
        },
      ]);
      saved = model.materialize("show");
    } finally {
      model.dispose();
    }
    /** Edits the saved file's marked objects the way a foreign tool would. */
    const tamper = (change: (object: number, index: number) => void) => {
      const document = pdfium.openDocument(saved);
      try {
        const page = lib.FPDF_LoadPage(document.handle, 0);
        try {
          for (
            let index = 1;
            index < lib.FPDFPage_CountObjects(page);
            index += 1
          )
            change(lib.FPDFPage_GetObject(page, index), index);
          lib.FPDFPage_GenerateContent(page);
        } finally {
          lib.FPDF_ClosePage(page);
        }
        return document.save("full");
      } finally {
        document.close();
      }
    };
    const kinds = (bytes: Uint8Array) => {
      const reopened = new PdfEditDocument(pdfium, bytes);
      try {
        return reopened.getElements({ pageIndex: 0 }).map((e) => e.kind);
      } finally {
        reopened.dispose();
      }
    };
    assert.deepEqual(kinds(saved), ["text", "textBox"], "untouched: one box");
    const moved = tamper((object) =>
      lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 0, -150),
    );
    const plain = (bytes: Uint8Array) => {
      const found = kinds(bytes);
      return found.length > 2 && found.every((kind) => kind === "text");
    };
    assert.equal(plain(moved), true, "moved lines are plain text objects");
    const retyped = tamper((object, index) => {
      if (index !== 1) return;
      const wide = pdfium.writeWideString("retyped");
      lib.FPDFText_SetText(object, wide);
      pdfium.free(wide);
    });
    assert.equal(plain(retyped), true, "retyped lines are plain text objects");
  });

  it("returns findText ranges into element text", async () => {
    const pdfium = await fixturePdfium();
    const model = new PdfEditDocument(
      pdfium,
      await buildPdf(["Hello brave world"]),
    );
    try {
      const [hit] = model.findText("brave", {});
      assert.deepEqual(hit?.ranges, [
        {
          start: { elementId: "p0:o0", offset: 6 },
          end: { elementId: "p0:o0", offset: 11 },
        },
      ]);
      const text = model.getElement("p0:o0")!.text!;
      assert.equal(text.slice(6, 11), "brave");

      const boxText = "first line words here\nsecond line target word";
      model.apply([
        {
          op: "insertTextBox",
          pageIndex: 0,
          rect: { x: 72, y: 200, width: 110, height: 100 },
          text: boxText,
        },
        {
          op: "insertTable",
          pageIndex: 0,
          at: { x: 72, y: 400 },
          width: 200,
          rows: [
            ["alpha", "beta"],
            ["gamma", "needle"],
          ],
        },
      ]);
      const [inBox] = model.findText("target", {});
      assert.equal(inBox?.ranges.length, 1);
      const range = inBox!.ranges[0]!;
      assert.equal(range.start.elementId, "p0:n1.0.0");
      assert.equal(
        boxText.slice(range.start.offset, range.end.offset),
        "target",
      );
      const [inTable] = model.findText("needle", {});
      const tableText = model.getElement("p0:n1.1.0")!.text!;
      const cell = inTable!.ranges[0]!;
      assert.equal(cell.start.elementId, "p0:n1.1.0");
      assert.equal(
        tableText.slice(cell.start.offset, cell.end.offset),
        "needle",
      );
    } finally {
      model.dispose();
    }
  });
});

describe("fonts (revision 2)", () => {
  const op = <T extends PdfOperation>(operation: T): T => operation;
  const ttf = new Uint8Array(
    readFileSync(
      new URL("../../fonts/noto-sans-latin-cyrillic.ttf", import.meta.url),
    ),
  );
  const FALLBACK_URL = "https://fonts.test/noto.ttf";

  async function engineWithFallback(original: Uint8Array) {
    const pair = loopbackWorker(
      createPdfEditHandler({
        loadPdfium: () => fixturePdfium(),
        fetchBytes: async (url) => {
          if (url === FALLBACK_URL) return ttf;
          throw new Error(`No font at ${url}`);
        },
        decodeImage: async () => {
          throw new Error("no images");
        },
      }),
    );
    return loadPdfEditEngine(
      original,
      { format: "pdf", limits: defaultResourceLimits, signal },
      { createWorker: () => pair.worker, fallbackFontUrl: FALLBACK_URL },
    );
  }

  /** A copy of the font whose glyph for `character` has no outline, like a careless subset. */
  function withEmptiedGlyph(font: Uint8Array, character: string): Uint8Array {
    const bytes = font.slice();
    const view = new DataView(bytes.buffer);
    const glyph = parseCmap(bytes).glyph(character.codePointAt(0)!);
    assert.ok(glyph, "glyph present");
    const tables = view.getUint16(4);
    let loca = -1;
    let head = -1;
    for (let index = 0; index < tables; index += 1) {
      const record = 12 + index * 16;
      const tag = String.fromCharCode(...bytes.subarray(record, record + 4));
      if (tag === "loca") loca = view.getUint32(record + 8);
      if (tag === "head") head = view.getUint32(record + 8);
    }
    assert.ok(loca > 0 && head > 0);
    const long = view.getInt16(head + 50) === 1;
    // The glyph starts where it ends: zero length, the next glyph untouched.
    if (long)
      view.setUint32(loca + glyph * 4, view.getUint32(loca + glyph * 4 + 4));
    else view.setUint16(loca + glyph * 2, view.getUint16(loca + glyph * 2 + 2));
    return bytes;
  }

  async function pdfWithEmbedded(font: Uint8Array, text: string) {
    const pdfium = await fixturePdfium();
    const { lib } = pdfium;
    const document = pdfium.createDocument();
    try {
      const page = lib.FPDFPage_New(document.handle, 0, 612, 792);
      const data = pdfium.writeBytes(font);
      const handle = lib.FPDFText_LoadFont(
        document.handle,
        data,
        font.length,
        1,
        true,
      );
      const object = lib.FPDFPageObj_CreateTextObj(document.handle, handle, 14);
      const wide = pdfium.writeWideString(text);
      lib.FPDFText_SetText(object, wide);
      pdfium.free(wide);
      lib.FPDFPageObj_Transform(object, 1, 0, 0, 1, 72, 700);
      lib.FPDFPage_InsertObject(page, object);
      lib.FPDFPage_GenerateContent(page);
      lib.FPDF_ClosePage(page);
      const bytes = document.save("full");
      pdfium.free(data);
      return bytes;
    } finally {
      document.close();
    }
  }

  it("rejects an emptied glyph of a subset font for in-place text", async () => {
    const subset = withEmptiedGlyph(ttf, "B");
    assert.equal(parseCmap(subset).has("B".codePointAt(0)!), true);
    assert.equal(parseCmap(subset).drawable("B".codePointAt(0)!), false);
    assert.equal(parseCmap(subset).drawable("A".codePointAt(0)!), true);
    const engine = await engineWithFallback(
      await pdfWithEmbedded(subset, "ABC"),
    );
    try {
      const kept = await engine.apply(
        [op({ op: "replaceText", target: "p0:o0", text: "AC" })],
        signal,
      );
      assert.deepEqual(kept.warnings, []);
      const fallen = await engine.apply(
        [op({ op: "replaceText", target: "p0:o0", text: "BB" })],
        signal,
      );
      assert.equal(fallen.warnings[0]?.code, "font-substitution");
    } finally {
      await engine.dispose();
    }
  });

  it("embeds the fallback font once for several text boxes", async () => {
    const engine = await engineWithFallback(await buildPdf(["One"]));
    try {
      for (const [index, text] of ["Перший", "Другий", "Третій"].entries())
        await engine.apply(
          [
            op({
              op: "insertTextBox",
              pageIndex: 0,
              rect: { x: 72, y: 100 + index * 60, width: 200, height: 40 },
              text,
            }),
          ],
          signal,
        );
      const saved = await engine.materialize("save", {}, signal);
      const text = Array.from(saved, (byte) => String.fromCharCode(byte)).join(
        "",
      );
      // PDFium embeds TrueType as a CID font with one /FontFile stream.
      assert.equal(text.match(/\/FontFile[23]? /g)?.length, 1);
    } finally {
      await engine.dispose();
    }
  });
});
