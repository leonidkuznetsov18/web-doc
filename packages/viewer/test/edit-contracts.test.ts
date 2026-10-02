import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type {
  DocumentAdapter,
  EditColor,
  EditElement,
  EditReceipt,
  EditSession,
  EditStateChange,
  LayoutChange,
  OperationIssue,
  PageRect,
  ReadItem,
  ReadResult,
  SavedDocument,
  TextRange,
  TextTarget,
  ViewerEventMap,
} from "../src/index.js";
import type { DocumentChange } from "../src/headless.js";
import {
  defaultResourceLimits,
  resolveLimits,
  ViewerClient,
  ViewerError,
} from "../src/index.js";

const pdfBytes = new TextEncoder().encode("%PDF-1.7\nedit-contracts");

function viewOnlyAdapter(): DocumentAdapter<{ readonly pages: number }> {
  return {
    id: "view-only",
    formats: ["pdf"],
    async open() {
      return { pages: 2 };
    },
    async getInfo(handle) {
      return { format: "pdf", unit: "page", pageCount: handle.pages };
    },
    async render() {},
    close() {},
  };
}

function isViewerError(code: string) {
  return (error: unknown) =>
    error instanceof ViewerError && error.code === code;
}

describe("editing contracts", () => {
  it("adds edit limits with defaults and validates them", () => {
    assert.equal(defaultResourceLimits.maxEditOperations, 500);
    assert.equal(defaultResourceLimits.maxEditHistory, 200);
    assert.equal(defaultResourceLimits.maxOutlineNodes, 5000);
    assert.equal(defaultResourceLimits.maxDescribeChars, 200_000);
    assert.equal(resolveLimits({}, { maxEditHistory: 5 }).maxEditHistory, 5);
    assert.throws(
      () => resolveLimits({}, { maxEditOperations: 0 }),
      isViewerError("resource-limit"),
    );
  });

  it("reports documents without an edit engine as not editable", async () => {
    const viewer = ViewerClient.create({
      adapters: [viewOnlyAdapter()],
    }).createViewer();
    await assert.rejects(viewer.edit(), isViewerError("lifecycle-error"));

    await viewer.load(pdfBytes, { fileName: "plain.pdf" });
    assert.equal(viewer.getDocumentInfo().capabilities?.editing, false);
    await assert.rejects(viewer.edit(), (error: unknown) => {
      assert.ok(error instanceof ViewerError);
      assert.equal(error.code, "edit-unsupported");
      assert.equal(error.details?.format, "pdf");
      return true;
    });
    assert.equal(viewer.getEditSession(), undefined);
  });

  it("returns no view geometry for a headless viewer", async () => {
    const viewer = ViewerClient.create({
      adapters: [viewOnlyAdapter()],
    }).createViewer();
    await viewer.load(pdfBytes, { fileName: "plain.pdf" });
    const rect: PageRect = { x: 10, y: 20, width: 30, height: 40 };
    assert.equal(viewer.pageToClient(0, rect), undefined);
    assert.equal(viewer.clientToPage(15, 25), undefined);
    // Malformed input is answered like an unmounted page, never thrown.
    assert.equal(viewer.pageToClient(0.5, rect), undefined);
    assert.equal(
      viewer.pageToClient(0, { ...rect, width: Number.NaN }),
      undefined,
    );
    assert.equal(viewer.clientToPage(Number.POSITIVE_INFINITY, 0), undefined);
    await viewer.close();
    assert.equal(viewer.pageToClient(0, rect), undefined);
    await viewer.destroy();
    assert.throws(
      () => viewer.clientToPage(0, 0),
      isViewerError("lifecycle-error"),
    );
  });

  it("types the new events, receipts and issues", () => {
    // Compile-time checks: these assignments fail to type-check if the
    // public shapes drift from the spec.
    const change: DocumentChange = {
      sessionId: "s",
      revision: 1,
      reason: "apply",
      changedPages: [0],
      pageCount: 2,
    };
    const state: EditStateChange = {
      active: true,
      format: "pdf",
      sessionId: "s",
      revision: 1,
      dirty: true,
      canUndo: true,
      canRedo: false,
      pageCount: 2,
    };
    const receipt: EditReceipt = {
      sessionId: "s",
      revision: 1,
      dryRun: false,
      operationCount: 1,
      createdIds: ["a"],
      removedIds: ["b"],
      changedPages: [0],
      pageCount: 2,
      warnings: [],
    };
    const issue: OperationIssue = {
      operationIndex: 0,
      path: "/style/color",
      code: "pattern",
      message: "Expected #RRGGBB",
    };
    const layout: LayoutChange = { sessionId: "s", revision: 1, pages: [0] };
    const events: Pick<
      ViewerEventMap,
      "documentchange" | "editstatechange" | "layoutchange"
    > = {
      documentchange: change,
      editstatechange: state,
      layoutchange: layout,
    };
    const format: EditSession["format"] = "pdf";
    assert.equal(events.documentchange.revision, receipt.revision);
    assert.equal(issue.operationIndex, 0);
    assert.equal(format, "pdf");
  });

  it("types revision-2 addressing, reads, saves and colours", () => {
    const range: TextRange = {
      start: { elementId: "p0:o1", offset: 0 },
      end: { elementId: "p0:o1", offset: 5 },
    };
    const target: TextTarget = {
      pageIndex: 0,
      text: "Total",
      rects: [{ x: 1, y: 2, width: 3, height: 4 }],
      elementIds: ["p0:o1"],
      ranges: [range],
    };
    const element: EditElement = {
      id: "p0:o1",
      kind: "text",
      pageIndex: 0,
      bounds: { x: 1, y: 2, width: 3, height: 4 },
      frame: {
        x: 1,
        y: 2,
        width: 3,
        height: 4,
        rotation: 0,
        flipH: false,
        flipV: false,
      },
      fragments: [
        { pageIndex: 0, bounds: { x: 1, y: 2, width: 3, height: 4 } },
      ],
      story: { kind: "header", scope: "default" },
      operations: [],
    };
    const items: ReadResult<EditElement> = {
      sessionId: "s",
      revision: 1,
      items: [element],
    };
    const item: ReadItem<EditElement> = {
      sessionId: "s",
      revision: 1,
      item: undefined,
    };
    const saved: SavedDocument = {
      bytes: new Uint8Array(),
      stateToken: "t",
      sessionId: "s",
      revision: 1,
      warnings: [],
    };
    const colours: EditColor[] = [
      "#ff0000",
      "#ff000080",
      "auto",
      { theme: "accent1", mods: { lumMod: 0.5 } },
    ];
    assert.equal(target.ranges?.[0]?.end.offset, 5);
    assert.equal(items.items.length + (item.item ? 1 : 0), 1);
    assert.equal(saved.stateToken, "t");
    assert.equal(colours.length, 4);
    assert.equal(
      defaultResourceLimits.maxEditCheckpointBytes,
      64 * 1024 * 1024,
    );
  });

  it("narrows the session union by format and keeps applyJson on the union", () => {
    // Compile-only: `format` narrows to the PDF session with its typed
    // methods, while the un-narrowed union still accepts plain JSON.
    const narrow = (session: EditSession) => {
      if (session.format === "pdf")
        return session.insertTextBox({
          pageIndex: 0,
          rect: { x: 0, y: 0, width: 1, height: 1 },
          text: "x",
        });
      return session.applyJson([{ op: "unknown" }], {
        expectedSessionId: session.sessionId,
      });
    };
    assert.equal(typeof narrow, "function");
  });
});
