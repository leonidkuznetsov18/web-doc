import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type {
  EditElement,
  EditSession,
  EditSessionBase,
  EditStateChange,
  ViewerApi,
} from "../src/index.js";
import { ViewerClient, ViewerError } from "../src/index.js";
import {
  encodePages,
  fakeEditableAdapter,
  fakeProvider,
  type FakeAdapterOptions,
  type FakeEditEngine,
  type FakeEngineOptions,
  type FakeOperation,
} from "./fixtures/fake-edit-engine.js";

const original = encodePages(["one two", "three", "four"]);

function viewerWith(
  adapterOptions: Omit<FakeAdapterOptions, "edit"> = {},
  engineOptions: FakeEngineOptions = {},
) {
  const engines: FakeEditEngine[] = [];
  const provider = fakeProvider(engineOptions, (engine) =>
    engines.push(engine),
  );
  // Kept as one object: the adapter reads flags such as failOpen at call time.
  const options: FakeAdapterOptions = { ...adapterOptions, edit: provider };
  const adapter = fakeEditableAdapter(options);
  const viewer = ViewerClient.create({ adapters: [adapter] }).createViewer();
  const events: { readonly type: string; readonly event: unknown }[] = [];
  for (const type of [
    "editstatechange",
    "documentchange",
    "searchchange",
    "selectionchange",
    "statechange",
  ] as const)
    viewer.on(type, (event) => events.push({ type, event }));
  return { viewer, adapter, options, provider, engines, events };
}

/** Typed batches: literals would fail excess-property checks against EditOperation. */
function ops(...operations: FakeOperation[]): FakeOperation[] {
  return operations;
}

type FakeSession = EditSessionBase<FakeOperation, EditElement>;

/** The fake format's session, seen with its own operation type. */
async function editFake(viewer: ViewerApi): Promise<FakeSession> {
  return (await viewer.edit()) as unknown as FakeSession;
}

function rejectsWith(code: string) {
  return (error: unknown): boolean => {
    assert.ok(error instanceof ViewerError, String(error));
    assert.equal(error.code, code, error.message);
    return true;
  };
}

describe("viewer editing integration", () => {
  it("advertises editing only when the adapter has an engine for the format", async () => {
    const { viewer, provider } = viewerWith();
    await viewer.load(original, { fileName: "doc.pdf" });
    assert.equal(viewer.getDocumentInfo().capabilities?.editing, true);
    assert.equal(provider.loads, 0);
    assert.equal(viewer.getEditSession(), undefined);

    const plain = ViewerClient.create({
      adapters: [fakeEditableAdapter()],
    }).createViewer();
    await plain.load(original, { fileName: "doc.pdf" });
    assert.equal(plain.getDocumentInfo().capabilities?.editing, false);
    await assert.rejects(plain.edit(), rejectsWith("edit-unsupported"));
  });

  it("starts the engine lazily, once, and shares it between calls", async () => {
    const { viewer, provider, events } = viewerWith();
    await viewer.load(original, { fileName: "doc.pdf" });
    const [first, second] = await Promise.all([viewer.edit(), viewer.edit()]);
    assert.equal(first, second);
    assert.equal(provider.loads, 1);
    assert.equal(await viewer.edit(), first);
    assert.equal(viewer.getEditSession(), first);
    assert.equal(first.format, "pdf");
    const started = events.find((entry) => entry.type === "editstatechange")
      ?.event as EditStateChange;
    assert.deepEqual(started, {
      active: true,
      format: "pdf",
      revision: 0,
      dirty: false,
      canUndo: false,
      canRedo: false,
      pageCount: 3,
    });
  });

  it("makes every read API reflect a change as soon as apply resolves", async () => {
    const { viewer, adapter, events } = viewerWith();
    await viewer.load(original, { fileName: "doc.pdf" });
    await viewer.search("three");
    await viewer.selectText({
      startPageIndex: 0,
      startOffset: 0,
      endPageIndex: 0,
      endOffset: 3,
    });
    viewer.goToPage(2);
    events.length = 0;

    const session = await editFake(viewer);
    const operations: FakeOperation[] = [
      { op: "setText", pageIndex: 1, text: "THREE changed" },
      { op: "deletePage", pageIndex: 2 },
    ];
    const receipt = await session.apply(operations);
    assert.equal(receipt.pageCount, 2);
    assert.equal(viewer.getDocumentInfo().pageCount, 2);
    assert.equal(viewer.getDocumentInfo().capabilities?.editing, true);
    assert.equal(viewer.state.pageCount, 2);
    assert.equal(viewer.state.pageIndex, 1);
    assert.equal(await viewer.getPageText(1), "THREE changed");
    assert.equal((await viewer.search("changed")).matches.length, 1);
    assert.equal(viewer.getSelection(), null);
    assert.deepEqual(adapter.closed, [1]);
    assert.deepEqual(
      events
        .map((entry) => entry.type)
        .filter((type) => type !== "statechange"),
      [
        "editstatechange",
        "searchchange",
        "selectionchange",
        "editstatechange",
        "documentchange",
        "searchchange",
      ],
    );
    // The original bytes stay available while editing.
    assert.deepEqual(viewer.getOriginalBytes(), original);
    assert.equal(session.state.dirty, true);
  });

  it("reopens through the adapter's reopen when it has one", async () => {
    const { viewer, adapter } = viewerWith({ reopen: true });
    await viewer.load(original, { fileName: "doc.pdf" });
    const session = await editFake(viewer);
    await session.apply(ops({ op: "setText", pageIndex: 0, text: "x" }));
    await session.undo();
    assert.deepEqual(adapter.reopened, [1, 2]);
    assert.deepEqual(adapter.closed, [1, 2]);
    assert.equal(await viewer.getPageText(0), "one two");
  });

  it("leaves the document untouched when the reopen fails", async () => {
    const { viewer, adapter, options, engines } = viewerWith();
    await viewer.load(original, { fileName: "doc.pdf" });
    const session = await editFake(viewer);
    options.failOpen = true;
    await assert.rejects(
      session.apply(ops({ op: "setText", pageIndex: 0, text: "x" })),
      (error: unknown) => {
        assert.ok(error instanceof ViewerError);
        assert.equal(error.code, "edit-failed");
        assert.equal(error.details?.stage, "reopen");
        return true;
      },
    );
    assert.equal(await viewer.getPageText(0), "one two");
    assert.equal(session.state.revision, 0);
    assert.deepEqual(engines[0]?.pages, ["one two", "three", "four"]);
    assert.deepEqual(adapter.closed, []);
  });

  it("ends the session when the document is replaced, closed or destroyed", async () => {
    const { viewer, engines, events } = viewerWith();
    await viewer.load(original, { fileName: "doc.pdf" });
    const session = await editFake(viewer);
    const pending = session.apply(ops({ op: "hang" }));
    await new Promise((resolve) => setTimeout(resolve, 0));
    await viewer.load(original, { fileName: "again.pdf" });
    await assert.rejects(pending, rejectsWith("aborted"));
    assert.equal(engines[0]?.disposed, true);
    assert.equal(viewer.getEditSession(), undefined);
    await assert.rejects(session.save(), rejectsWith("lifecycle-error"));
    const ended = events
      .filter((entry) => entry.type === "editstatechange")
      .map((entry) => (entry.event as EditStateChange).active);
    assert.deepEqual(ended, [true, false]);

    const next = await viewer.edit();
    assert.notEqual(next, session);
    await viewer.close();
    assert.equal(engines[1]?.disposed, true);
    await assert.rejects(viewer.edit(), rejectsWith("lifecycle-error"));

    await viewer.load(original, { fileName: "doc.pdf" });
    const last = await viewer.edit();
    await viewer.destroy();
    assert.equal(engines[2]?.disposed, true);
    await assert.rejects(last.undo(), rejectsWith("lifecycle-error"));
  });

  it("reports an engine that fails to start and allows a retry", async () => {
    let fail = true;
    const provider = {
      loads: 0,
      formats: ["pdf"] as const,
      async load(bytes: Uint8Array) {
        provider.loads += 1;
        if (fail) throw new Error("wasm missing");
        return fakeProvider().load(bytes, {
          format: "pdf",
          limits: ViewerClient.create().limits,
          signal: new AbortController().signal,
        });
      },
      createSession: (core: unknown) => core as EditSession,
    };
    const viewer = ViewerClient.create({
      adapters: [fakeEditableAdapter({ edit: provider })],
    }).createViewer();
    await viewer.load(original, { fileName: "doc.pdf" });
    await assert.rejects(viewer.edit(), (error: unknown) => {
      assert.ok(error instanceof ViewerError);
      assert.equal(error.code, "edit-failed");
      assert.equal(error.details?.stage, "load");
      return true;
    });
    assert.equal(viewer.getEditSession(), undefined);
    fail = false;
    const session = await viewer.edit();
    assert.equal(provider.loads, 2);
    assert.equal(session.state.pageCount, 3);
  });
});
