import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { EditSessionController } from "../src/edit/session.js";
import type { EditOperation } from "../src/index.js";
import { ViewerError } from "../src/index.js";
import {
  encodePages,
  FakeEditEngine,
  FakeHost,
  type FakeEngineOptions,
  type FakeHostOptions,
  type FakeOperation,
} from "./fixtures/fake-edit-engine.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * Task 61 of the ai-edit module: named checkpoints pin a state the host
 * names, restore it as one undoable history entry, survive undo, redo and
 * reset, hold their bytes against eviction and replay from the original
 * when the budget cannot hold them.
 */

const original = encodePages(["one", "two", "three"]);

function session(
  engineOptions: FakeEngineOptions = {},
  hostOptions: FakeHostOptions = {},
) {
  const engine = new FakeEditEngine(original, engineOptions);
  const host = new FakeHost(hostOptions);
  const controller = new EditSessionController(engine, host, original, 3);
  const apply = (
    operations: readonly (FakeOperation | EditOperation)[],
    options = {},
  ) => controller.apply(operations as readonly EditOperation[], options);
  return { engine, host, session: controller, apply };
}

const text = (index: number) => ({
  op: "setText" as const,
  pageIndex: 0,
  text: `s${index}`,
});

/** The batch count of the engine's latest restore. */
const lastRestore = (calls: readonly string[]) =>
  calls.filter((call) => call.startsWith("restore:")).at(-1);

function rejectsWith(code: string) {
  return (error: unknown): boolean => {
    assert.ok(error instanceof ViewerError, String(error));
    assert.equal(error.code, code, error.message);
    return true;
  };
}

describe("named checkpoints", () => {
  it("creates, lists and drops checkpoints within the limit", async () => {
    const { session: edit, apply } = session(
      {},
      { limits: { maxEditCheckpoints: 2 } },
    );
    await apply([text(1)]);
    const first = await edit.createCheckpoint("before the turn");
    assert.match(first.id, /^[A-Za-z0-9_-]{22}$/);
    assert.equal(first.label, "before the turn");
    assert.equal(first.revision, 1);
    assert.ok(!Number.isNaN(Date.parse(first.createdAt)));
    assert.equal(Object.isFrozen(first), true);
    const second = await edit.createCheckpoint();
    assert.equal("label" in second, false);
    assert.notEqual(second.id, first.id);
    assert.deepEqual(edit.listCheckpoints(), [first, second]);
    await assert.rejects(
      edit.createCheckpoint("third"),
      rejectsWith("resource-limit"),
    );
    edit.dropCheckpoint("nope");
    edit.dropCheckpoint(first.id);
    assert.deepEqual(edit.listCheckpoints(), [second]);
    const third = await edit.createCheckpoint("third");
    assert.deepEqual(
      edit.listCheckpoints().map((item) => item.id),
      [second.id, third.id],
    );
    await edit.end();
    assert.deepEqual(
      edit.listCheckpoints().map((item) => item.id),
      [second.id, third.id],
    );
    await assert.rejects(
      edit.createCheckpoint(),
      rejectsWith("lifecycle-error"),
    );
  });

  it("restores a checkpoint as one undoable entry with events and a receipt", async () => {
    const { session: edit, host, apply } = session();
    await apply([text(1)]);
    await apply([{ op: "insertPage", index: 1, text: "extra" }]);
    const checkpoint = await edit.createCheckpoint("two changes");
    await apply([{ op: "deletePage", pageIndex: 1 }]);
    await apply([{ op: "insertPage", index: 0, text: "front" }]);
    assert.deepEqual(host.current, ["front", "s1", "two", "three"]);
    host.events.length = 0;
    const receipt = await edit.restoreCheckpoint(checkpoint.id);
    assert.deepEqual(host.current, ["s1", "extra", "two", "three"]);
    assert.deepEqual(receipt, {
      sessionId: edit.sessionId,
      revision: 5,
      dryRun: false,
      operationCount: 0,
      createdIds: [],
      removedIds: ["page:front"],
      changedPages: [0, 1, 2, 3],
      pageCount: 4,
      warnings: [],
    });
    assert.deepEqual(
      host.events.map((event) => [
        event.type,
        (event.event as { reason?: string }).reason,
      ]),
      [
        ["editstatechange", undefined],
        ["documentchange", "restore"],
      ],
    );
    assert.equal(edit.state.revision, 5);
    assert.equal(edit.state.canUndo, true);
    assert.equal(edit.state.pageCount, 4);
    // The restore is one undo step; redo brings the checkpoint back.
    await edit.undo();
    assert.deepEqual(host.current, ["front", "s1", "two", "three"]);
    await edit.redo();
    assert.deepEqual(host.current, ["s1", "extra", "two", "three"]);
    // Restoring the state the session is in changes nothing.
    const noop = await edit.restoreCheckpoint(checkpoint.id);
    assert.equal(noop.revision, 7);
    assert.equal(edit.state.revision, 7);
    assert.deepEqual(noop.changedPages, []);
    // A later change after the restore builds on the checkpoint's content.
    await apply([text(8)]);
    assert.deepEqual(host.current, ["s8", "extra", "two", "three"]);
    await edit.undo();
    await edit.undo();
    assert.deepEqual(host.current, ["front", "s1", "two", "three"]);
    await edit.end();
  });

  it("refuses unknown ids and stale revisions", async () => {
    const { session: edit, apply } = session();
    const checkpoint = await edit.createCheckpoint();
    await apply([text(1)]);
    await assert.rejects(
      edit.restoreCheckpoint("missing"),
      rejectsWith("invalid-operation"),
    );
    await assert.rejects(
      edit.restoreCheckpoint(checkpoint.id, { expectedRevision: 0 }),
      rejectsWith("edit-conflict"),
    );
    assert.equal(edit.state.revision, 1);
    const receipt = await edit.restoreCheckpoint(checkpoint.id, {
      expectedRevision: 1,
    });
    assert.equal(receipt.revision, 2);
    await edit.end();
  });

  it("survives undo, redo and reset, and keeps dirty exact", async () => {
    const { session: edit, host, apply } = session();
    await apply([text(1)]);
    await apply([text(2)]);
    const checkpoint = await edit.createCheckpoint();
    const saved = await edit.save();
    edit.markSaved(saved.stateToken);
    assert.equal(edit.state.dirty, false);
    await edit.undo();
    await apply([text(9)]);
    assert.equal(edit.state.dirty, true);
    assert.deepEqual(host.current, ["s9", "two", "three"]);
    // The checkpoint sits on the branch the new change dropped.
    await edit.restoreCheckpoint(checkpoint.id);
    assert.deepEqual(host.current, ["s2", "two", "three"]);
    assert.equal(edit.state.dirty, false, "the saved content again");
    await edit.reset();
    assert.deepEqual(host.current, ["one", "two", "three"]);
    assert.deepEqual(edit.listCheckpoints(), [checkpoint]);
    await edit.restoreCheckpoint(checkpoint.id);
    assert.deepEqual(host.current, ["s2", "two", "three"]);
    assert.equal(edit.state.dirty, false);
    await edit.undo();
    assert.deepEqual(host.current, ["one", "two", "three"]);
    await edit.end();
  });

  it("restores the original through a checkpoint of revision 0", async () => {
    const { session: edit, host, engine, apply } = session();
    const start = await edit.createCheckpoint("start");
    await apply([text(1)]);
    await apply([text(2)]);
    const receipt = await edit.restoreCheckpoint(start.id);
    assert.deepEqual(host.current, ["one", "two", "three"]);
    assert.equal(engine.restoreBases.at(-1), undefined);
    assert.equal(lastRestore(engine.calls), "restore:0");
    assert.equal(receipt.revision, 3);
    const saved = await edit.save();
    assert.deepEqual(saved.bytes, original);
    assert.equal(edit.state.dirty, false);
    await edit.redo(); // nothing to redo
    await edit.undo();
    assert.deepEqual(host.current, ["s2", "two", "three"]);
    await edit.end();
  });

  it("pins the checkpoint's bytes against eviction", async () => {
    // maxEditHistory 8 retains every second state; the budget holds one.
    const {
      session: edit,
      engine,
      apply,
    } = session(
      {},
      {
        limits: {
          maxEditHistory: 8,
          maxEditCheckpointBytes: encodePages(["s2", "two", "three"])
            .byteLength,
        },
      },
    );
    await apply([text(1)]);
    await apply([text(2)]);
    const checkpoint = await edit.createCheckpoint();
    for (let index = 3; index <= 7; index += 1) await apply([text(index)]);
    // Later retained states could not evict the pinned one: the restore
    // starts from its bytes, and an undo to state 6 replays from them too.
    await edit.restoreCheckpoint(checkpoint.id);
    assert.deepEqual(engine.restoreBases.at(-1), ["s2", "two", "three"]);
    assert.equal(lastRestore(engine.calls), "restore:0");
    await edit.undo();
    assert.deepEqual(engine.restoreBases.at(-1), ["s2", "two", "three"]);
    assert.equal(lastRestore(engine.calls), "restore:5");
    // Dropping the checkpoint frees its bytes for the ordinary retention.
    edit.dropCheckpoint(checkpoint.id);
    await apply([text(8)]);
    await edit.undo();
    assert.notDeepEqual(engine.restoreBases.at(-1), ["s2", "two", "three"]);
    await edit.end();
  });

  it("replays from the original when the budget cannot hold the bytes", async () => {
    const {
      session: edit,
      host,
      engine,
      apply,
    } = session({}, { limits: { maxEditCheckpointBytes: 1 } });
    await apply([text(1)]);
    await apply([text(2)]);
    const checkpoint = await edit.createCheckpoint();
    await apply([text(3)]);
    await edit.restoreCheckpoint(checkpoint.id);
    assert.deepEqual(host.current, ["s2", "two", "three"]);
    assert.equal(engine.restoreBases.at(-1), undefined);
    assert.equal(lastRestore(engine.calls), "restore:2");
    // A checkpoint taken after a restore replays through it.
    await apply([text(4)]);
    const later = await edit.createCheckpoint();
    await edit.reset();
    await edit.restoreCheckpoint(later.id);
    assert.deepEqual(host.current, ["s4", "two", "three"]);
    assert.equal(engine.restoreBases.at(-1), undefined);
    assert.equal(lastRestore(engine.calls), "restore:3");
    await edit.undo();
    assert.deepEqual(host.current, ["one", "two", "three"]);
    await edit.end();
  });

  it("works on a PDF session with real element ids", async () => {
    const { session: edit, end } = await pdfSession(await buildPdf(["Hello"]));
    try {
      const checkpoint = await edit.createCheckpoint("before the agent");
      const inserted = await edit.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 100, width: 200, height: 40 },
        text: "Agent text",
      });
      const boxId = inserted.createdIds[0]!;
      assert.equal((await edit.getElements()).items.length, 2);
      const receipt = await edit.restoreCheckpoint(checkpoint.id);
      assert.deepEqual(receipt.removedIds, [boxId]);
      assert.equal(receipt.revision, 2);
      assert.equal((await edit.getElements()).items.length, 1);
      // An undo reports what it removed, as every undo does; the box is back.
      const back = await edit.undo();
      assert.deepEqual(back.removedIds, []);
      assert.equal(back.revision, 3);
      assert.deepEqual(
        (await edit.getElements()).items.map((item) => item.id),
        [(await edit.getElements()).items[0]!.id, boxId],
      );
      assert.equal(edit.listCheckpoints().length, 1);
    } finally {
      await end();
    }
  });
});
