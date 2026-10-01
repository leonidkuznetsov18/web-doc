import { linkedAbortController } from "../abort.js";
import type { ResourceLimits, ViewerEventMap } from "../contracts.js";
import { abortError, ViewerError } from "../errors.js";
import type { EditEngine, EngineChange } from "./engine.js";
import { EditHistory, type HistoryEntry } from "./history.js";
import {
  assertBatchSize,
  checkOperations,
  freezeOperations,
  invalidOperationError,
} from "./operations.js";
import type {
  ApplyOptions,
  DocumentChangeReason,
  EditableFormat,
  EditElement,
  EditFindOptions,
  EditOperation,
  EditReceipt,
  EditSessionBase,
  EditState,
  ElementQuery,
  HistoryOptions,
  OperationSchemaSet,
  PagePoint,
  SaveOptions,
  TextTarget,
} from "./types.js";

/** What the viewer provides to a session: a place to show bytes, and events. */
export interface EditSessionHost {
  readonly format: EditableFormat;
  readonly limits: ResourceLimits;
  /** Shows `bytes` as the current document and returns its page count. */
  replaceDocument(bytes: Uint8Array, signal: AbortSignal): Promise<number>;
  emit<K extends "editstatechange" | "documentchange">(
    type: K,
    event: ViewerEventMap[K],
  ): void;
}

type FailureStage = "apply" | "materialize" | "reopen";

/**
 * The format-independent editing session: validation, history, revisions,
 * saving and the viewer refresh. Format modules wrap it to add typed methods.
 * Calls run one at a time in call order.
 */
export class EditSessionController implements EditSessionBase<
  EditOperation,
  EditElement
> {
  readonly format: EditableFormat;
  readonly schemas: OperationSchemaSet;
  readonly #engine: EditEngine;
  readonly #host: EditSessionHost;
  readonly #history: EditHistory;
  readonly #original: Uint8Array;
  readonly #originalPageCount: number;
  readonly #ending = new AbortController();
  readonly sessionId = newSessionId();
  #state: EditState;
  #queue: Promise<unknown> = Promise.resolve();
  #revision = 0;
  #savedStateId = 0;
  #ended = false;
  /** Set when a rollback failed: the engine no longer matches the history. */
  #broken = false;

  constructor(
    engine: EditEngine,
    host: EditSessionHost,
    original: Uint8Array,
    originalPageCount: number,
  ) {
    this.format = host.format;
    this.schemas = engine.schemas;
    this.#engine = engine;
    this.#host = host;
    this.#original = original;
    this.#originalPageCount = originalPageCount;
    this.#history = new EditHistory(
      host.limits.maxEditHistory,
      originalPageCount,
    );
    this.#state = this.#snapshot();
  }

  get state(): EditState {
    return this.#state;
  }

  applyJson(
    operations: readonly EditOperation[],
    options: ApplyOptions = {},
  ): Promise<EditReceipt> {
    return this.apply(operations, options);
  }

  apply(
    operations: readonly EditOperation[],
    options: ApplyOptions = {},
  ): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      assertBatchSize(operations, this.#host.limits.maxEditOperations);
      const shapeIssues = checkOperations(operations, this.schemas);
      if (shapeIssues.length > 0) throw invalidOperationError(shapeIssues);
      const batch = freezeOperations(operations);
      const engineIssues = await this.#engine.validate(batch, signal);
      throwIfAborted(signal);
      if (engineIssues.length > 0) throw invalidOperationError(engineIssues);

      if (options.dryRun) {
        const change = await this.#transaction(signal, "apply", async () => {
          const result = await this.#engine.apply(batch, signal);
          // Nothing moves, so the working copy goes back to the current state.
          await this.#engine.restore(this.#history.applied(), signal);
          return result;
        });
        return this.#receipt(true, batch.length, change.createdIds, change);
      }

      const before = this.#history.pageCount;
      const change = await this.#transaction(signal, "apply", async () => {
        const result = await this.#engine.apply(batch, signal);
        await this.#show(signal);
        return result;
      });
      this.#history.push({
        operations: batch,
        ...(options.label === undefined ? {} : { label: options.label }),
        changedPages: change.changedPages,
        pageCountBefore: before,
        pageCountAfter: change.pageCount,
      });
      this.#commit("apply", change.changedPages, change.pageCount);
      return this.#receipt(false, batch.length, change.createdIds, change);
    });
  }

  undo(options: HistoryOptions = {}): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const entry = this.#history.undoEntry;
      if (!entry) return this.#noop();
      await this.#moveTo(this.#history.position - 1, signal);
      this.#history.undo();
      const pageCount = entry.pageCountBefore;
      const changedPages = pagesTouched(entry, pageCount);
      this.#commit("undo", changedPages, pageCount);
      return this.#receipt(false, entry.operations.length, [], {
        changedPages,
        pageCount,
        warnings: [],
      });
    });
  }

  redo(options: HistoryOptions = {}): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const entry = this.#history.redoEntry;
      if (!entry) return this.#noop();
      await this.#moveTo(this.#history.position + 1, signal);
      this.#history.redo();
      const pageCount = entry.pageCountAfter;
      const changedPages = pagesTouched(entry, pageCount);
      this.#commit("redo", changedPages, pageCount);
      return this.#receipt(false, entry.operations.length, [], {
        changedPages,
        pageCount,
        warnings: [],
      });
    });
  }

  reset(options: HistoryOptions = {}): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const applied = this.#history.applied();
      if (this.#history.stateId === 0 && this.#history.isPristine)
        return this.#noop();
      await this.#transaction(signal, "apply", async () => {
        await this.#engine.restore([], signal);
        await this.#show(signal);
      });
      this.#history.clear();
      const pageCount = this.#originalPageCount;
      const changedPages = allPages(pageCount);
      this.#commit("reset", changedPages, pageCount);
      return this.#receipt(
        false,
        applied.reduce((count, batch) => count + batch.length, 0),
        [],
        { changedPages, pageCount, warnings: [] },
      );
    });
  }

  save(options: SaveOptions = {}): Promise<Uint8Array> {
    return this.#enqueue(options.signal, async (signal) => {
      const bytes =
        this.#history.stateId === 0
          ? this.#original.slice()
          : await this.#engine.materialize(signal);
      throwIfAborted(signal);
      const wasDirty = this.#state.dirty;
      this.#savedStateId = this.#history.stateId;
      this.#state = this.#snapshot();
      if (wasDirty !== this.#state.dirty) this.#emitState();
      return bytes;
    });
  }

  getElements(query: ElementQuery = {}): Promise<readonly EditElement[]> {
    return this.#enqueue(undefined, (signal) =>
      this.#engine.getElements(query, signal),
    );
  }

  getElement(id: string): Promise<EditElement | undefined> {
    return this.#enqueue(undefined, async (signal) =>
      this.#engine.getElement
        ? this.#engine.getElement(id, signal)
        : (await this.#engine.getElements({}, signal)).find(
            (element) => element.id === id,
          ),
    );
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
  ): Promise<readonly EditElement[]> {
    return this.#enqueue(undefined, (signal) =>
      this.#engine.elementsAt(pageIndex, point, signal),
    );
  }

  findText(
    query: string,
    options: EditFindOptions = {},
  ): Promise<readonly TextTarget[]> {
    return this.#enqueue(undefined, (signal) =>
      this.#engine.findText(query, options, signal),
    );
  }

  /**
   * Ends the session: pending calls reject with `aborted`, later calls with
   * `lifecycle-error`, the engine is disposed and `active: false` is emitted.
   */
  async end(): Promise<void> {
    if (this.#ended) return;
    this.#ended = true;
    this.#ending.abort(abortError());
    await this.#queue.catch(() => undefined);
    try {
      await this.#engine.dispose();
    } finally {
      this.#emitState(false);
    }
  }

  get ended(): boolean {
    return this.#ended;
  }

  /** False once the session ended or a failed recovery left it inconsistent. */
  get usable(): boolean {
    return !this.#ended && !this.#broken;
  }

  /** Runs `task` after every earlier call, with a signal bounded by `maxOperationMs`. */
  #enqueue<T>(
    signal: AbortSignal | undefined,
    task: (signal: AbortSignal) => Promise<T>,
  ): Promise<T> {
    const run = async (): Promise<T> => {
      this.#assertAlive();
      if (signal?.aborted) throw abortError();
      const controller = linkedAbortController(signal, this.#ending.signal);
      let timedOut = false;
      const timer = setTimeout(() => {
        timedOut = true;
        controller.abort(abortError());
      }, this.#host.limits.maxOperationMs);
      try {
        return await task(controller.signal);
      } catch (error) {
        // A dead engine worker cannot be recovered; the next edit() starts anew.
        if (error instanceof ViewerError && error.code === "worker-crashed")
          this.#broken = true;
        if (timedOut)
          throw new ViewerError(
            "resource-limit",
            "Edit operation exceeded maxOperationMs",
            { details: { timeoutMs: this.#host.limits.maxOperationMs } },
          );
        throw error;
      } finally {
        clearTimeout(timer);
      }
    };
    const result = this.#queue.then(run, run);
    this.#queue = result.catch(() => undefined);
    return result;
  }

  /**
   * Runs `work` against the engine and the viewer. On any failure the engine
   * goes back to the current history state and the viewer is left untouched,
   * so a failed call never leaks a partial change.
   */
  async #transaction<T>(
    signal: AbortSignal,
    stage: FailureStage,
    work: () => Promise<T>,
  ): Promise<T> {
    let failedStage: FailureStage = stage;
    try {
      return await work();
    } catch (error) {
      if (error instanceof ViewerError && error.code === "edit-failed")
        failedStage = (error.details?.stage as FailureStage) ?? stage;
      await this.#rollback();
      if (signal.aborted) throw abortError();
      if (error instanceof ViewerError && error.code === "edit-failed")
        throw error;
      throw new ViewerError(
        "edit-failed",
        `Editing failed while ${describe(failedStage)}; the document is unchanged`,
        { cause: error, details: { stage: failedStage } },
      );
    }
  }

  /** Materializes the working copy and shows it in the viewer. */
  async #show(signal: AbortSignal): Promise<void> {
    let bytes: Uint8Array;
    try {
      bytes = await this.#engine.materialize(signal);
    } catch (error) {
      throw stageError("materialize", error);
    }
    throwIfAborted(signal);
    try {
      await this.#host.replaceDocument(bytes, signal);
    } catch (error) {
      throw stageError("reopen", error);
    }
    throwIfAborted(signal);
  }

  async #moveTo(position: number, signal: AbortSignal): Promise<void> {
    await this.#transaction(signal, "apply", async () => {
      await this.#engine.restore(this.#history.batchesAt(position), signal);
      await this.#show(signal);
    });
  }

  async #rollback(): Promise<void> {
    // Recovery must finish even when the caller's signal is gone, so it gets
    // a fresh budget of its own.
    const controller = new AbortController();
    const timer = setTimeout(
      () => controller.abort(abortError()),
      this.#host.limits.maxOperationMs,
    );
    try {
      await this.#engine.restore(this.#history.applied(), controller.signal);
    } catch {
      this.#broken = true;
    } finally {
      clearTimeout(timer);
    }
  }

  #commit(
    reason: DocumentChangeReason,
    changedPages: readonly number[],
    pageCount: number,
  ): void {
    this.#revision += 1;
    this.#state = this.#snapshot();
    this.#emitState();
    this.#host.emit("documentchange", {
      sessionId: this.sessionId,
      revision: this.#revision,
      reason,
      changedPages: Object.freeze([...changedPages]),
      pageCount,
    });
  }

  #snapshot(): EditState {
    return Object.freeze({
      sessionId: this.sessionId,
      revision: this.#revision,
      dirty: this.#history.stateId !== this.#savedStateId,
      canUndo: this.#history.canUndo,
      canRedo: this.#history.canRedo,
      pageCount: this.#history.pageCount,
    });
  }

  #emitState(active = true): void {
    this.#host.emit("editstatechange", {
      ...this.#state,
      active,
      format: this.format,
    });
  }

  #receipt(
    dryRun: boolean,
    operationCount: number,
    createdIds: readonly string[],
    change: Pick<EngineChange, "changedPages" | "pageCount" | "warnings"> &
      Partial<Pick<EngineChange, "removedIds" | "remappedIds">>,
  ): EditReceipt {
    return Object.freeze({
      sessionId: this.sessionId,
      revision: this.#revision,
      dryRun,
      operationCount,
      createdIds: Object.freeze([...createdIds]),
      removedIds: Object.freeze([...(change.removedIds ?? [])]),
      ...(change.remappedIds
        ? { remappedIds: Object.freeze({ ...change.remappedIds }) }
        : {}),
      changedPages: Object.freeze([...change.changedPages]),
      pageCount: change.pageCount,
      warnings: Object.freeze(
        change.warnings.map((warning) => Object.freeze({ ...warning })),
      ),
    });
  }

  #noop(): EditReceipt {
    return this.#receipt(false, 0, [], {
      changedPages: [],
      pageCount: this.#history.pageCount,
      warnings: [],
    });
  }

  #assertRevision(options: {
    readonly expectedRevision?: number;
    readonly expectedSessionId?: string;
  }): void {
    const { expectedRevision, expectedSessionId } = options;
    const staleSession =
      expectedSessionId !== undefined && expectedSessionId !== this.sessionId;
    const staleRevision =
      expectedRevision !== undefined && expectedRevision !== this.#revision;
    if (staleSession || staleRevision)
      throw new ViewerError(
        "edit-conflict",
        staleSession
          ? "The session the caller read from has ended"
          : "The document changed since it was read",
        {
          details: {
            ...(expectedRevision === undefined ? {} : { expectedRevision }),
            revision: this.#revision,
            ...(expectedSessionId === undefined ? {} : { expectedSessionId }),
            sessionId: this.sessionId,
          },
        },
      );
  }

  #assertAlive(): void {
    if (this.#ended)
      throw new ViewerError("lifecycle-error", "The edit session has ended");
    if (this.#broken)
      throw new ViewerError(
        "edit-failed",
        "The edit session is unusable after a failed recovery",
        { details: { stage: "apply", recovered: false } },
      );
  }
}

function stageError(stage: FailureStage, cause: unknown): ViewerError {
  if (cause instanceof ViewerError && cause.code === "aborted") return cause;
  return new ViewerError(
    "edit-failed",
    `Editing failed while ${describe(stage)}; the document is unchanged`,
    { cause, details: { stage } },
  );
}

function describe(stage: FailureStage): string {
  switch (stage) {
    case "apply":
      return "applying the change";
    case "materialize":
      return "producing the edited file";
    case "reopen":
      return "reopening the edited file";
  }
}

function throwIfAborted(signal: AbortSignal): void {
  if (signal.aborted) throw abortError();
}

/** Pages to re-render when an entry is undone or redone. */
function pagesTouched(entry: HistoryEntry, pageCount: number): number[] {
  return entry.pageCountBefore === entry.pageCountAfter
    ? [...entry.changedPages]
    : allPages(pageCount);
}

function allPages(pageCount: number): number[] {
  return Array.from({ length: pageCount }, (_, index) => index);
}

/** 128 random bits as URL-safe base64; unique across sessions and reloads. */
function newSessionId(): string {
  const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/, "");
}
