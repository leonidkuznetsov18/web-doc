import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type {
  DocumentAdapter,
  EditReceipt,
  EditSession,
  EditStateChange,
  OperationIssue,
  PageRect,
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
      revision: 1,
      reason: "apply",
      changedPages: [0],
      pageCount: 2,
    };
    const state: EditStateChange = {
      active: true,
      format: "pdf",
      revision: 1,
      dirty: true,
      canUndo: true,
      canRedo: false,
      pageCount: 2,
    };
    const receipt: EditReceipt = {
      revision: 1,
      dryRun: false,
      operationCount: 1,
      createdIds: ["a"],
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
    const events: Pick<ViewerEventMap, "documentchange" | "editstatechange"> = {
      documentchange: change,
      editstatechange: state,
    };
    const format: EditSession["format"] = "pptx";
    assert.equal(events.documentchange.revision, receipt.revision);
    assert.equal(issue.operationIndex, 0);
    assert.equal(format, "pptx");
  });
});
