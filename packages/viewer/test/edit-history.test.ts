import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { EditHistory } from "../src/edit/history.js";

function entry(label: string, pageCountAfter = 1) {
  return {
    operations: [{ op: label }],
    label,
    createdIds: [],
    removedIds: [],
    changedPages: [0],
    pageCountBefore: 1,
    pageCountAfter,
  };
}

describe("EditHistory", () => {
  it("tracks position, redo tail and content ids", () => {
    const history = new EditHistory(10, 1);
    assert.equal(history.stateId, 0);
    assert.equal(history.isPristine, true);
    history.push(entry("a"));
    history.push(entry("b"));
    assert.deepEqual(
      history.applied().map((batch) => batch.operations[0]?.op),
      ["a", "b"],
    );
    history.undo();
    assert.equal(history.canRedo, true);
    assert.equal(history.stateId, 1);
    history.push(entry("c"));
    assert.equal(history.canRedo, false);
    assert.deepEqual(
      history.applied().map((batch) => batch.operations[0]?.op),
      ["a", "c"],
    );
    assert.equal(history.stateId, 3);
    history.undo();
    history.undo();
    assert.equal(history.stateId, 0);
    assert.equal(history.canUndo, false);
    assert.equal(history.isPristine, false);
    history.undo();
    assert.equal(history.position, 0);
  });

  it("folds entries beyond the limit into the starting point", () => {
    const history = new EditHistory(2, 1);
    history.push(entry("a"));
    history.push(entry("b"));
    history.push(entry("c"));
    assert.equal(history.position, 2);
    history.undo();
    history.undo();
    assert.equal(history.canUndo, false);
    // "a" stays applied: it can no longer be undone.
    assert.deepEqual(
      history.applied().map((batch) => batch.operations[0]?.op),
      ["a"],
    );
    assert.notEqual(history.stateId, 0);
    history.clear();
    assert.equal(history.stateId, 0);
    assert.equal(history.isPristine, true);
  });

  it("hands out state ids that are never reused, even after an undo", () => {
    const history = new EditHistory(10, 1);
    assert.equal(history.nextStateId, 1);
    const a = history.push(entry("a"));
    history.undo();
    const b = history.push(entry("b"));
    assert.notEqual(a.stateId, b.stateId);
    assert.deepEqual(
      history.applied().map((batch) => batch.stateId),
      [b.stateId],
    );
    assert.equal(history.nextStateId, b.stateId + 1);
  });

  it("reports the page count of the current state", () => {
    const history = new EditHistory(10, 4);
    assert.equal(history.pageCount, 4);
    history.push(entry("a", 5));
    assert.equal(history.pageCount, 5);
    history.undo();
    assert.equal(history.pageCount, 4);
  });
});
