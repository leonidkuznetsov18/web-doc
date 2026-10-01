import { linkedAbortController } from "../abort.js";
import type {
  ResourceLimits,
  ViewerEventMap,
  ViewerWarning,
} from "../contracts.js";
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
  SavedDocument,
  SaveOptions,
  TextTarget,
} from "./types.js";

/** What the viewer provides to a session: a place to show bytes, and events. */
export interface EditSessionHost {
  readonly format: EditableFormat;
  readonly limits: ResourceLimits;
  /** Opens `bytes` next to the current document; may fail or be aborted. */
  prepareDocument(
    bytes: Uint8Array,
    signal: AbortSignal,
  ): Promise<PreparedDocument>;
  /**
   * Shows a prepared document: synchronous, cannot fail, and returns the page
   * count the renderer reports. `changedPages` lets the viewer repaint only
   * what changed.
   */
  commitDocument(
    prepared: PreparedDocument,
    changedPages: readonly number[],
  ): number;
  /** Releases a preparation that will not be shown. */
  discardDocument(prepared: PreparedDocument): void;
  emit<K extends "editstatechange" | "documentchange">(
    type: K,
    event: ViewerEventMap[K],
  ): void;
}

/** What `prepareDocument` hands back; hosts attach their own handle to it. */
export interface PreparedDocument {
  readonly pageCount: number;
}

/** What a transaction showed: the bytes and the renderer's page count. */
interface Shown {
  readonly bytes: Uint8Array;
  readonly pageCount: number;
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
  /** Bytes of the last committed state; what the viewer shows and what a broken session saves. */
  #committedBytes: Uint8Array;
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
    this.#committedBytes = original;
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
    // Copied before the call is queued, so what the engine sees is what the
    // caller passed, whatever it does to its objects while waiting.
    const batch = Array.isArray(operations)
      ? freezeOperations(operations)
      : operations;
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      assertBatchSize(batch, this.#host.limits.maxEditOperations);
      const shapeIssues = checkOperations(batch, this.schemas);
      if (shapeIssues.length > 0) throw invalidOperationError(shapeIssues);
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
        return this.#receipt(true, batch.length, change.createdIds, {
          ...change,
          pageCount: change.pageCount ?? this.#history.pageCount,
        });
      }

      const before = this.#history.pageCount;
      const { change, shown } = await this.#transaction(
        signal,
        "apply",
        async () => {
          const result = await this.#engine.apply(batch, signal);
          return {
            change: result,
            shown: await this.#show(signal, result.changedPages),
          };
        },
      );
      this.#history.push({
        operations: batch,
        ...(options.label === undefined ? {} : { label: options.label }),
        changedPages: change.changedPages,
        pageCountBefore: before,
        pageCountAfter: shown.pageCount,
      });
      this.#commit("apply", change.changedPages, shown);
      return this.#receipt(false, batch.length, change.createdIds, {
        ...change,
        pageCount: shown.pageCount,
        warnings: [...change.warnings, ...pageCountWarning(change, shown)],
      });
    });
  }

  undo(options: HistoryOptions = {}): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const entry = this.#history.undoEntry;
      if (!entry) return this.#noop();
      const changedPages = pagesTouched(entry, entry.pageCountBefore);
      const shown = await this.#moveTo(
        this.#history.position - 1,
        changedPages,
        signal,
      );
      this.#history.undo();
      this.#commit("undo", changedPages, shown);
      return this.#receipt(false, entry.operations.length, [], {
        changedPages,
        pageCount: shown.pageCount,
        warnings: [],
      });
    });
  }

  redo(options: HistoryOptions = {}): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const entry = this.#history.redoEntry;
      if (!entry) return this.#noop();
      const changedPages = pagesTouched(entry, entry.pageCountAfter);
      const shown = await this.#moveTo(
        this.#history.position + 1,
        changedPages,
        signal,
      );
      this.#history.redo();
      this.#commit("redo", changedPages, shown);
      return this.#receipt(false, entry.operations.length, [], {
        changedPages,
        pageCount: shown.pageCount,
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
      const changedPages = allPages(
        Math.max(this.#originalPageCount, this.#history.pageCount),
      );
      const shown = await this.#transaction(signal, "apply", async () => {
        await this.#engine.restore([], signal);
        return this.#show(signal, changedPages);
      });
      this.#history.clear();
      this.#commit("reset", changedPages, shown);
      return this.#receipt(
        false,
        applied.reduce((count, batch) => count + batch.length, 0),
        [],
        { changedPages, pageCount: shown.pageCount, warnings: [] },
      );
    });
  }

  save(options: SaveOptions = {}): Promise<SavedDocument> {
    return this.#enqueue(
      options.signal,
      async (signal) => {
        // A session whose recovery failed still hands out what it last
        // showed; the history state matches those bytes.
        const bytes = this.#broken
          ? this.#committedBytes.slice()
          : this.#history.stateId === 0
            ? this.#original.slice()
            : await this.#engine.materialize(signal);
        throwIfAborted(signal);
        return Object.freeze({
          bytes,
          stateToken: this.#stateToken(this.#history.stateId),
          sessionId: this.sessionId,
          revision: this.#revision,
        });
      },
      { allowBroken: true },
    );
  }

  /** Records that the host persisted the state a `save()` token names. */
  markSaved(stateToken: string): void {
    const stateId = this.#parseStateToken(stateToken);
    if (stateId === undefined) {
      reportError(
        new ViewerError(
          "edit-conflict",
          "The save token belongs to another session",
          { details: { stateToken, sessionId: this.sessionId } },
        ),
      );
      return;
    }
    const wasDirty = this.#state.dirty;
    this.#savedStateId = stateId;
    this.#state = this.#snapshot();
    if (wasDirty !== this.#state.dirty) this.#emitState();
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
    options: { readonly allowBroken?: boolean } = {},
  ): Promise<T> {
    // Calls made after the session ended fail at once; calls still queued
    // when it ends are cancelled, like any other pending work.
    if (this.#ended)
      return Promise.reject(
        new ViewerError("lifecycle-error", "The edit session has ended"),
      );
    const run = async (): Promise<T> => {
      if (this.#ending.signal.aborted) throw abortError();
      if (!options.allowBroken) this.#assertAlive();
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

  /**
   * Materializes the working copy and shows it in the viewer in two phases:
   * the preparation may fail or be aborted and is then discarded; the commit
   * is synchronous and cannot fail, so once it ran the call completes
   * whatever its signal says.
   */
  async #show(
    signal: AbortSignal,
    changedPages: readonly number[],
  ): Promise<Shown> {
    let bytes: Uint8Array;
    try {
      bytes = await this.#engine.materialize(signal);
    } catch (error) {
      throw stageError("materialize", error);
    }
    throwIfAborted(signal);
    let prepared: PreparedDocument;
    try {
      prepared = await this.#host.prepareDocument(bytes, signal);
    } catch (error) {
      throw stageError("reopen", error);
    }
    if (signal.aborted) {
      this.#host.discardDocument(prepared);
      throw abortError();
    }
    return {
      bytes,
      pageCount: this.#host.commitDocument(prepared, changedPages),
    };
  }

  async #moveTo(
    position: number,
    changedPages: readonly number[],
    signal: AbortSignal,
  ): Promise<Shown> {
    return this.#transaction(signal, "apply", async () => {
      await this.#engine.restore(this.#history.batchesAt(position), signal);
      return this.#show(signal, changedPages);
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
    shown: Shown,
  ): void {
    this.#committedBytes = shown.bytes;
    this.#revision += 1;
    this.#state = this.#snapshot();
    this.#emitState();
    this.#emit("documentchange", {
      sessionId: this.sessionId,
      revision: this.#revision,
      reason,
      changedPages: Object.freeze([...changedPages]),
      pageCount: shown.pageCount,
    });
  }

  /** Hands an event to the host; a throwing listener is reported, never propagated. */
  #emit<K extends "editstatechange" | "documentchange">(
    type: K,
    event: ViewerEventMap[K],
  ): void {
    try {
      this.#host.emit(type, event);
    } catch (error) {
      reportError(error);
    }
  }

  #stateToken(stateId: number): string {
    return `${this.sessionId}:${stateId}`;
  }

  #parseStateToken(token: string): number | undefined {
    const prefix = `${this.sessionId}:`;
    if (!token.startsWith(prefix)) return undefined;
    const stateId = Number(token.slice(prefix.length));
    return Number.isSafeInteger(stateId) && stateId >= 0 ? stateId : undefined;
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
    this.#emit("editstatechange", {
      ...this.#state,
      active,
      format: this.format,
    });
  }

  #receipt(
    dryRun: boolean,
    operationCount: number,
    createdIds: readonly string[],
    change: Pick<EngineChange, "changedPages" | "warnings"> &
      Partial<Pick<EngineChange, "removedIds" | "remappedIds">> & {
        readonly pageCount: number;
      },
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

/** The renderer owns the page count; an engine that disagrees is reported, not trusted. */
function pageCountWarning(
  change: EngineChange,
  shown: Shown,
): readonly ViewerWarning[] {
  if (change.pageCount === undefined || change.pageCount === shown.pageCount)
    return [];
  return [
    {
      code: "fidelity-degraded",
      message: `The engine reports ${change.pageCount} pages but the renderer shows ${shown.pageCount}`,
      details: { engine: change.pageCount, renderer: shown.pageCount },
    },
  ];
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

/** Surfaces a listener's exception without failing the call that emitted the event. */
function reportError(error: unknown): void {
  const report = (globalThis as { reportError?: (error: unknown) => void })
    .reportError;
  if (report) report(error);
  else console.error(error);
}
