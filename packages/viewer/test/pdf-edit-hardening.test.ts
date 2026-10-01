import assert from "node:assert/strict";
import { before, describe, it } from "node:test";

import { PdfEditDocument } from "../src/edit/pdf/engine/document.js";
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
} from "../src/index.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";
import { buildPdf, fixturePdfium } from "./fixtures/pdf-builder.js";
import { signedPdf } from "./fixtures/signed-pdf.js";
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
    replaceDocument: (bytes) => pageCountOf(bytes),
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
      assert.deepEqual(first.warnings[0]?.details, { signatures: 1 });
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
        (await session.getElements({ kinds: ["textBox"] })).length,
        0,
      );
      assert.deepEqual(await session.save(), original);

      const real = await session.insertTextBox(fields);
      assert.deepEqual(real, { ...dry, revision: 1, dryRun: false });
      assert.equal(session.state.revision, 1);
      assert.equal((await session.getElement("p0:n1.0.0"))?.text, "Maybe");
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
      const bytesA = await a.session.save();
      const bytesB = await b.session.save();
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
      const kinds = (await client.getElements({ pageIndex: 0 })).map(
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
