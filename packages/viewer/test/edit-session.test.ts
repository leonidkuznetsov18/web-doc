import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { EditSessionController } from "../src/edit/session.js";
import type {
  EditOperation,
  EditStateChange,
  OperationIssue,
} from "../src/index.js";
import { ViewerError } from "../src/index.js";
import {
  decodePages,
  encodePages,
  FakeEditEngine,
  FakeHost,
  type FakeEngineOptions,
  type FakeHostOptions,
  type FakeOperation,
} from "./fixtures/fake-edit-engine.js";

const original = encodePages(["one", "two", "three"]);

function session(
  engineOptions: FakeEngineOptions = {},
  hostOptions: FakeHostOptions = {},
) {
  const engine = new FakeEditEngine(original, engineOptions);
  const host = new FakeHost(hostOptions);
  const controller = new EditSessionController(engine, host, original, 3);
  // Loose on purpose: invalid shapes must reach the validator.
  const apply = (
    operations: readonly (FakeOperation | EditOperation)[],
    options = {},
  ) => controller.apply(operations, options);
  return { engine, host, session: controller, apply };
}

function rejectsWith(code: string, check?: (error: ViewerError) => void) {
  return (error: unknown): boolean => {
    assert.ok(error instanceof ViewerError, String(error));
    assert.equal(error.code, code, error.message);
    check?.(error);
    return true;
  };
}

describe("EditSessionController", () => {
  it("starts clean and applies a batch end to end", async () => {
    const { session: edit, host, apply } = session();
    assert.deepEqual(edit.state, {
      sessionId: edit.sessionId,
      revision: 0,
      dirty: false,
      canUndo: false,
      canRedo: false,
      pageCount: 3,
    });
    const receipt = await apply(
      [{ op: "setText", pageIndex: 1, text: "TWO" }],
      { label: "shout" },
    );
    assert.deepEqual(receipt, {
      sessionId: edit.sessionId,
      revision: 1,
      dryRun: false,
      operationCount: 1,
      createdIds: [],
      removedIds: [],
      changedPages: [1],
      pageCount: 3,
      warnings: [],
    });
    assert.equal(Object.isFrozen(receipt), true);
    assert.deepEqual(host.current, ["one", "TWO", "three"]);
    assert.deepEqual(edit.state, {
      sessionId: edit.sessionId,
      revision: 1,
      dirty: true,
      canUndo: true,
      canRedo: false,
      pageCount: 3,
    });
    assert.deepEqual(host.eventTypes, ["editstatechange", "documentchange"]);
    assert.deepEqual(host.events[1]?.event, {
      sessionId: edit.sessionId,
      revision: 1,
      reason: "apply",
      changedPages: [1],
      pageCount: 3,
    });
  });

  it("rejects stale revisions before touching anything", async () => {
    const { session: edit, engine, apply } = session();
    await apply([{ op: "setText", pageIndex: 0, text: "x" }]);
    await assert.rejects(
      apply([{ op: "setText", pageIndex: 0, text: "y" }], {
        expectedRevision: 0,
      }),
      rejectsWith("edit-conflict", (error) =>
        assert.deepEqual(error.details, {
          expectedRevision: 0,
          revision: 1,
          sessionId: edit.sessionId,
        }),
      ),
    );
    await assert.rejects(
      edit.undo({ expectedRevision: 5 }),
      rejectsWith("edit-conflict"),
    );
    assert.equal(engine.calls.filter((call) => call === "apply").length, 1);
    await apply([{ op: "setText", pageIndex: 0, text: "y" }], {
      expectedRevision: 1,
    });
    assert.equal(edit.state.revision, 2);
  });

  it("reports shape issues, then engine issues, without applying", async () => {
    const { session: edit, engine, host, apply } = session();
    await assert.rejects(
      apply([{ op: "setText", pageIndex: -1 }, { op: "nope" }]),
      rejectsWith("invalid-operation", (error) => {
        const issues = error.details?.issues as readonly OperationIssue[];
        assert.deepEqual(
          issues.map(
            (issue) => `${issue.operationIndex}${issue.path}:${issue.code}`,
          ),
          ["0/text:required", "0/pageIndex:minimum", "1/op:unknown-operation"],
        );
      }),
    );
    await assert.rejects(
      apply([
        { op: "setText", pageIndex: 0, text: "ok" },
        { op: "setText", pageIndex: 9, text: "missing" },
      ]),
      rejectsWith("invalid-operation", (error) => {
        const issues = error.details?.issues as readonly OperationIssue[];
        assert.deepEqual(issues, [
          {
            operationIndex: 1,
            path: "/pageIndex",
            code: "unknown-target",
            message: "No page 9",
          },
        ]);
      }),
    );
    await assert.rejects(edit.apply([]), rejectsWith("invalid-operation"));
    await assert.rejects(
      edit.apply(
        Array.from({ length: 501 }, () => ({
          op: "setText",
          pageIndex: 0,
          text: "",
        })),
      ),
      rejectsWith("resource-limit"),
    );
    assert.equal(engine.calls.includes("apply"), false);
    assert.equal(edit.state.revision, 0);
    assert.equal(host.events.length, 0);
  });

  it("rolls back a failing engine, materialize or reopen and reports the stage", async () => {
    const failing = session();
    await failing.apply([{ op: "setText", pageIndex: 2, text: "kept" }]);
    await assert.rejects(
      failing.apply([{ op: "fail" }]),
      rejectsWith("edit-failed", (error) =>
        assert.equal(error.details?.stage, "apply"),
      ),
    );
    assert.deepEqual(failing.engine.pages, ["one", "two", "kept"]);
    // The failed apply, then the rollback replaying the one kept batch.
    assert.deepEqual(failing.engine.calls.slice(-3), [
      "apply",
      "restore:1",
      "apply",
    ]);
    assert.equal(failing.session.state.revision, 1);
    assert.equal(failing.host.shown.length, 1);

    const noDisk = session({ failMaterialize: true });
    await assert.rejects(
      noDisk.apply([{ op: "setText", pageIndex: 0, text: "x" }]),
      rejectsWith("edit-failed", (error) =>
        assert.equal(error.details?.stage, "materialize"),
      ),
    );
    assert.deepEqual(noDisk.engine.pages, ["one", "two", "three"]);

    const noRender = session({}, { failReplace: true });
    await assert.rejects(
      noRender.apply([{ op: "setText", pageIndex: 0, text: "x" }]),
      rejectsWith("edit-failed", (error) =>
        assert.equal(error.details?.stage, "reopen"),
      ),
    );
    assert.deepEqual(noRender.engine.pages, ["one", "two", "three"]);
    assert.equal(noRender.session.state.revision, 0);
    assert.deepEqual(noRender.host.events, []);
    // The session is still usable once the renderer recovers.
    noRender.host.options.failReplace = false;
    await noRender.apply([{ op: "setText", pageIndex: 0, text: "x" }]);
    assert.equal(noRender.session.state.revision, 1);
  });

  it("becomes unusable when a rollback itself fails", async () => {
    const { session: edit, engine, apply } = session();
    engine.options.failRestore = true;
    await assert.rejects(apply([{ op: "fail" }]), rejectsWith("edit-failed"));
    await assert.rejects(
      apply([{ op: "setText", pageIndex: 0, text: "x" }]),
      rejectsWith("edit-failed", (error) =>
        assert.equal(error.details?.recovered, false),
      ),
    );
  });

  it("honours aborts before and during a change", async () => {
    const { session: edit, engine, host, apply } = session();
    const aborted = new AbortController();
    aborted.abort();
    await assert.rejects(
      apply([{ op: "setText", pageIndex: 0, text: "x" }], {
        signal: aborted.signal,
      }),
      rejectsWith("aborted"),
    );
    assert.equal(engine.calls.length, 0);

    const controller = new AbortController();
    const pending = apply(
      [{ op: "setText", pageIndex: 0, text: "x" }, { op: "hang" }],
      {
        signal: controller.signal,
      },
    );
    await new Promise((resolve) => setTimeout(resolve, 0));
    controller.abort();
    await assert.rejects(pending, rejectsWith("aborted"));
    assert.deepEqual(engine.pages, ["one", "two", "three"]);
    assert.equal(edit.state.revision, 0);
    assert.equal(host.events.length, 0);
  });

  it("turns an exceeded maxOperationMs into resource-limit and rolls back", async () => {
    const {
      session: edit,
      engine,
      apply,
    } = session({}, { limits: { maxOperationMs: 20 } });
    await assert.rejects(
      apply([{ op: "hang" }]),
      rejectsWith("resource-limit"),
    );
    assert.deepEqual(engine.pages, ["one", "two", "three"]);
    assert.equal(edit.state.revision, 0);
  });

  it("runs calls one at a time in call order", async () => {
    const { session: edit, host, apply, engine } = session();
    const results = await Promise.all([
      apply([{ op: "setText", pageIndex: 0, text: "a" }]),
      edit.getElements({ pageIndex: 0 }),
      apply([{ op: "setText", pageIndex: 0, text: "b" }]),
      edit.undo(),
    ]);
    assert.equal((results[0] as { revision: number }).revision, 1);
    assert.deepEqual(
      (results[1] as readonly { text?: string }[]).map((e) => e.text),
      ["a"],
    );
    assert.equal((results[2] as { revision: number }).revision, 2);
    assert.equal((results[3] as { revision: number }).revision, 3);
    assert.deepEqual(host.current, ["a", "two", "three"]);
    assert.deepEqual(engine.calls.slice(0, 5), [
      "validate",
      "apply",
      "materialize",
      "getElements",
      "validate",
    ]);
  });

  it("dry runs validate and simulate without changing anything", async () => {
    const { session: edit, host, engine, apply } = session();
    const receipt = await apply([{ op: "insertPage", index: 0, text: "new" }], {
      dryRun: true,
    });
    assert.deepEqual(receipt, {
      sessionId: edit.sessionId,
      revision: 0,
      dryRun: true,
      operationCount: 1,
      createdIds: ["page:new"],
      removedIds: [],
      changedPages: [0, 1, 2, 3],
      pageCount: 4,
      warnings: [],
    });
    assert.deepEqual(engine.pages, ["one", "two", "three"]);
    assert.equal(edit.state.revision, 0);
    assert.equal(edit.state.dirty, false);
    assert.equal(host.events.length, 0);
    assert.equal(host.shown.length, 0);
    await assert.rejects(
      apply([{ op: "deletePage", pageIndex: 7 }], { dryRun: true }),
      rejectsWith("invalid-operation"),
    );
  });

  it("undoes, redoes and resets with no-op receipts at the ends", async () => {
    const { session: edit, host, apply } = session();
    const noop = await edit.undo();
    assert.deepEqual(noop, {
      sessionId: edit.sessionId,
      revision: 0,
      dryRun: false,
      operationCount: 0,
      createdIds: [],
      removedIds: [],
      changedPages: [],
      pageCount: 3,
      warnings: [],
    });
    assert.equal(host.events.length, 0);

    await apply([{ op: "setText", pageIndex: 0, text: "a" }]);
    await apply([
      { op: "insertPage", index: 1, text: "inserted" },
      { op: "setText", pageIndex: 3, text: "c" },
    ]);
    assert.deepEqual(host.current, ["a", "inserted", "two", "c"]);
    assert.equal(edit.state.pageCount, 4);

    const undone = await edit.undo();
    assert.deepEqual(host.current, ["a", "two", "three"]);
    assert.equal(undone.revision, 3);
    assert.equal(undone.operationCount, 2);
    // The page count changed, so every page of the result is reported.
    assert.deepEqual(undone.changedPages, [0, 1, 2]);
    assert.deepEqual(edit.state, {
      sessionId: edit.sessionId,
      revision: 3,
      dirty: true,
      canUndo: true,
      canRedo: true,
      pageCount: 3,
    });

    const redone = await edit.redo();
    assert.deepEqual(host.current, ["a", "inserted", "two", "c"]);
    assert.equal(redone.pageCount, 4);
    assert.equal(edit.state.canRedo, false);

    await edit.undo();
    await apply([{ op: "setText", pageIndex: 2, text: "tail" }]);
    assert.equal(edit.state.canRedo, false);
    assert.equal((await edit.redo()).operationCount, 0);

    const reset = await edit.reset();
    assert.deepEqual(host.current, ["one", "two", "three"]);
    assert.equal(reset.operationCount, 2);
    assert.deepEqual(reset.changedPages, [0, 1, 2]);
    assert.deepEqual(edit.state, {
      sessionId: edit.sessionId,
      revision: 7,
      dirty: false,
      canUndo: false,
      canRedo: false,
      pageCount: 3,
    });
    assert.deepEqual(
      host.events
        .filter((entry) => entry.type === "documentchange")
        .map((entry) => (entry.event as { reason: string }).reason),
      ["apply", "apply", "undo", "redo", "undo", "apply", "reset"],
    );
    assert.equal((await edit.reset()).operationCount, 0);
  });

  it("folds history beyond maxEditHistory and keeps reset reachable", async () => {
    const {
      session: edit,
      host,
      apply,
    } = session({}, { limits: { maxEditHistory: 2 } });
    for (const text of ["a", "b", "c"])
      await apply([{ op: "setText", pageIndex: 0, text }]);
    await edit.undo();
    await edit.undo();
    assert.equal(edit.state.canUndo, false);
    assert.deepEqual(host.current, ["a", "two", "three"]);
    await edit.reset();
    assert.deepEqual(host.current, ["one", "two", "three"]);
  });

  it("tracks dirty across save, undo and reset", async () => {
    const { session: edit, host, apply } = session();
    assert.deepEqual(await edit.save(), original);
    assert.equal(edit.state.dirty, false);

    await apply([{ op: "setText", pageIndex: 0, text: "a" }]);
    assert.equal(edit.state.dirty, true);
    const saved = await edit.save();
    assert.deepEqual(decodePages(saved), ["a", "two", "three"]);
    assert.equal(edit.state.dirty, false);
    assert.equal(edit.state.revision, 1);
    const stateEvents = host.events.filter(
      (entry) => entry.type === "editstatechange",
    );
    assert.equal(stateEvents.length, 2);
    assert.equal((stateEvents.at(-1)?.event as EditStateChange).dirty, false);

    await edit.undo();
    assert.equal(edit.state.dirty, true);
    await edit.redo();
    assert.equal(edit.state.dirty, false);
    await edit.reset();
    assert.equal(edit.state.dirty, true);
    await edit.save();
    assert.equal(edit.state.dirty, false);
    // apply, save, undo, redo, reset, save; the first save changed nothing.
    assert.equal(
      host.events.filter((entry) => entry.type === "editstatechange").length,
      6,
    );
  });

  it("materializes identical bytes for the same history and after undoing to zero", async () => {
    const first = session();
    const second = session();
    for (const edit of [first, second]) {
      await edit.apply([{ op: "setText", pageIndex: 1, text: "same" }]);
      await edit.apply([{ op: "insertPage", index: 0, text: "front" }]);
    }
    assert.deepEqual(await first.session.save(), await second.session.save());
    await first.session.undo();
    await first.session.undo();
    assert.deepEqual(await first.session.save(), original);
  });

  it("answers element and text queries through the engine", async () => {
    const { session: edit, apply } = session();
    await apply([{ op: "setText", pageIndex: 0, text: "alpha beta" }]);
    assert.deepEqual(
      (await edit.getElements({ pageIndex: 0 })).map((element) => element.text),
      ["alpha", "beta"],
    );
    assert.equal((await edit.getElement("p0w1"))?.text, "beta");
    assert.equal(await edit.getElement("missing"), undefined);
    assert.deepEqual(
      (await edit.elementsAt(0, { x: 15, y: 1 })).map((element) => element.id),
      ["p0w1"],
    );
    assert.deepEqual(
      (await edit.findText("BETA")).map((target) => target.pageIndex),
      [0],
    );
    assert.deepEqual(await edit.findText("BETA", { caseSensitive: true }), []);
  });

  it("ends: pending calls abort, later calls fail, the engine is disposed", async () => {
    const { session: edit, engine, host, apply } = session();
    const pending = apply([{ op: "hang" }]);
    await new Promise((resolve) => setTimeout(resolve, 0));
    await edit.end();
    await assert.rejects(pending, rejectsWith("aborted"));
    assert.equal(engine.disposed, true);
    assert.equal(edit.ended, true);
    await assert.rejects(
      apply([{ op: "setText", pageIndex: 0, text: "x" }]),
      rejectsWith("lifecycle-error"),
    );
    await assert.rejects(edit.save(), rejectsWith("lifecycle-error"));
    const last = host.events.at(-1)?.event as EditStateChange;
    assert.equal(last.active, false);
    assert.equal(last.format, "pdf");
    await edit.end();
  });
});

describe("session identity (revision 2)", () => {
  it("gives every session a unique id and stamps it on state, receipts and events", async () => {
    const first = session();
    const second = session();
    assert.match(first.session.sessionId, /^[A-Za-z0-9_-]{22}$/);
    assert.notEqual(first.session.sessionId, second.session.sessionId);
    assert.equal(first.session.state.sessionId, first.session.sessionId);
    const receipt = await first.apply([
      { op: "setText", pageIndex: 0, text: "x" },
    ]);
    assert.equal(receipt.sessionId, first.session.sessionId);
    assert.deepEqual(receipt.removedIds, []);
    const change = first.host.events.find(
      (entry) => entry.type === "documentchange",
    )?.event as { sessionId: string };
    assert.equal(change.sessionId, first.session.sessionId);
  });

  it("rejects calls that name another session, before the revision check", async () => {
    const { session: edit, engine, apply } = session();
    const other = session().session.sessionId;
    await assert.rejects(
      apply([{ op: "setText", pageIndex: 0, text: "x" }], {
        expectedSessionId: other,
        expectedRevision: 0,
      }),
      rejectsWith("edit-conflict", (error) =>
        assert.deepEqual(error.details, {
          expectedRevision: 0,
          revision: 0,
          expectedSessionId: other,
          sessionId: edit.sessionId,
        }),
      ),
    );
    assert.deepEqual(engine.calls, []);
    // Plain JSON, as an AI client would hold it.
    const json = JSON.parse('{"op":"setText","pageIndex":0,"text":"x"}');
    const ok = await edit.applyJson([json as EditOperation], {
      expectedSessionId: edit.sessionId,
    });
    assert.equal(ok.revision, 1);
    await assert.rejects(
      edit.undo({ expectedSessionId: other }),
      rejectsWith("edit-conflict"),
    );
  });
});
