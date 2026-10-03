import { linkedAbortController } from "../abort.js";
import type {
  ResourceLimits,
  ViewerEventMap,
  ViewerWarning,
} from "../contracts.js";
import { abortError, ViewerError } from "../errors.js";
import { readDescription, readOutline } from "./ai/outline.js";
import { resolveTargets } from "./ai/targets.js";
import { buildToolSet, callTool as runTool } from "./ai/tools.js";
import type {
  DescribeOptions,
  DocumentDescription,
  EditCheckpoint,
  OutlineOptions,
  OutlineResult,
  TargetCandidate,
  TargetQuery,
  ToolCall,
  ToolCallOptions,
  ToolResult,
  ToolSet,
} from "./ai/types.js";
import {
  assetIdOf,
  AssetStore,
  binaryFields,
  isAssetReference,
} from "./assets.js";
import type {
  EditEngine,
  EditSessionCore,
  EngineBatch,
  EngineChange,
  MaterializedDocument,
  MaterializeOptions,
  RestoreTarget,
} from "./engine.js";
import { batchOf, EditHistory, modeOf, type HistoryEntry } from "./history.js";
import {
  assertBatchSize,
  checkOperations,
  freezeOperations,
  invalidOperationError,
  parseReference,
} from "./operations.js";
import type {
  AssetOptions,
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
  OperationIssue,
  ReadItem,
  ReadOptions,
  ReadResult,
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
  /**
   * Flow formats: the first page whose text runs carry `paragraphId`, when
   * the host has laid it out; the core repaints from it to the end.
   */
  pageOf?(paragraphId: string): number | undefined;
  emit<K extends "editstatechange" | "documentchange">(
    type: K,
    event: ViewerEventMap[K],
  ): void;
}

/** What `prepareDocument` hands back; hosts attach their own handle to it. */
export interface PreparedDocument {
  readonly pageCount: number;
}

/** What a transaction showed: the bytes, the renderer's page count and the pages repainted. */
interface Shown {
  readonly bytes: Uint8Array;
  readonly pageCount: number;
  readonly changedPages: readonly number[];
}

/** The pages to repaint, known once the renderer has counted the shown document. */
type ChangedPages =
  readonly number[] | ((pageCount: number) => readonly number[]);

type FailureStage = "apply" | "materialize" | "reopen";

/**
 * The format-independent editing session: validation, history, revisions,
 * saving and the viewer refresh. Format modules wrap it to add typed methods.
 * Calls run one at a time in call order.
 */
export class EditSessionController implements EditSessionCore {
  readonly format: EditableFormat;
  readonly schemas: OperationSchemaSet;
  readonly #engine: EditEngine;
  readonly #host: EditSessionHost;
  readonly #history: EditHistory;
  readonly #original: Uint8Array;
  readonly #originalPageCount: number;
  readonly #ending = new AbortController();
  readonly sessionId = randomId();
  /** Bytes of the last committed state; what the viewer shows and what a broken session saves. */
  #committedBytes: Uint8Array;
  /** Materialized bytes of some committed states, by state id, so restores replay less. */
  readonly #checkpoints = new Map<number, Uint8Array>();
  #checkpointBytes = 0;
  /** Named checkpoints by id, in creation order. */
  readonly #named = new Map<string, NamedCheckpoint>();
  /** State ids named checkpoints pin, with how many name each; never evicted. */
  readonly #pinned = new Map<number, number>();
  /** Binary payloads of this session's batches, by content id. */
  readonly #assets = new AssetStore();
  #state: EditState;
  #tools: ToolSet | undefined;
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

  get limits(): ResourceLimits {
    return this.#host.limits;
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
      const modeIssues = checkChangeMode(this.format, options);
      if (modeIssues.length > 0) throw invalidOperationError(modeIssues);
      const shapeIssues = checkOperations(batch, this.schemas);
      if (shapeIssues.length > 0) throw invalidOperationError(shapeIssues);
      const referenceIssues = checkBatchReferences(batch);
      if (referenceIssues.length > 0)
        throw invalidOperationError(referenceIssues);
      const interned = await this.#intern(batch, signal);
      const mode = modeOf(options);
      const engineIssues = await this.#engine.validate(interned, signal, mode);
      throwIfAborted(signal);
      if (engineIssues.length > 0) throw invalidOperationError(engineIssues);
      // The id the history will give this state; a dry run uses the same one,
      // so its receipt names the ids a real apply would.
      const engineBatch: EngineBatch = {
        stateId: this.#history.nextStateId,
        operations: interned,
        ...mode,
      };

      if (options.dryRun) {
        const change = await this.#transaction(signal, "apply", async () => {
          const result = await this.#engine.apply(engineBatch, signal);
          // Producing the bytes catches what only saving would; then the
          // working copy goes back to the current state.
          await this.#engine.materialize("save", {}, signal);
          await this.#engine.restore(this.#restoreTarget(), signal);
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
          const result = await this.#engine.apply(engineBatch, signal);
          return {
            change: result,
            shown: await this.#show(
              signal,
              this.#pagesOf(result.changedPages, result.reflowFrom),
            ),
          };
        },
      );
      this.#history.push({
        operations: interned,
        ...mode,
        ...(options.label === undefined ? {} : { label: options.label }),
        createdIds: change.createdIds,
        removedIds: change.removedIds,
        changedPages: shown.changedPages,
        ...(change.reflowFrom === undefined
          ? {}
          : { reflowFrom: change.reflowFrom }),
        pageCountBefore: before,
        pageCountAfter: shown.pageCount,
      });
      this.#commit("apply", shown.changedPages, shown);
      return this.#receipt(false, batch.length, change.createdIds, {
        ...change,
        changedPages: shown.changedPages,
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
      const shown = await this.#moveTo(
        this.#history.position - 1,
        this.#pagesOf(
          pagesTouched(entry, entry.pageCountBefore),
          entry.reflowFrom,
        ),
        signal,
      );
      this.#history.undo();
      this.#commit("undo", shown.changedPages, shown);
      return this.#receipt(false, entry.operations.length, [], {
        removedIds: entry.createdIds,
        changedPages: shown.changedPages,
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
      const shown = await this.#moveTo(
        this.#history.position + 1,
        this.#pagesOf(
          pagesTouched(entry, entry.pageCountAfter),
          entry.reflowFrom,
        ),
        signal,
      );
      this.#history.redo();
      this.#commit("redo", shown.changedPages, shown);
      return this.#receipt(false, entry.operations.length, entry.createdIds, {
        removedIds: entry.removedIds,
        changedPages: shown.changedPages,
        pageCount: shown.pageCount,
        warnings: [],
      });
    });
  }

  reset(options: HistoryOptions = {}): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const applied = this.#history.entriesAt(this.#history.position);
      if (this.#history.stateId === 0 && this.#history.isPristine)
        return this.#noop();
      const changedPages = allPages(
        Math.max(this.#originalPageCount, this.#history.pageCount),
      );
      const shown = await this.#transaction(signal, "apply", async () => {
        await this.#engine.restore({ batches: [] }, signal);
        return this.#show(signal, changedPages);
      });
      this.#history.clear();
      this.#commit("reset", changedPages, shown);
      return this.#receipt(
        false,
        applied.reduce((count, entry) => count + entry.operations.length, 0),
        [],
        {
          removedIds: applied.flatMap((entry) => entry.createdIds),
          changedPages,
          pageCount: shown.pageCount,
          warnings: [],
        },
      );
    });
  }

  save(options: SaveOptions = {}): Promise<SavedDocument> {
    // Format fields (the PDF save mode) travel to the engine; the signal stays.
    const { signal: own, ...format } = options as SaveOptions &
      Record<string, unknown>;
    return this.#enqueue(
      own,
      async (signal) => {
        // A session whose recovery failed still hands out what it last
        // showed; the history state matches those bytes.
        const { bytes, warnings } = this.#broken
          ? { bytes: this.#committedBytes.slice(), warnings: [] }
          : this.#history.stateId === 0
            ? { bytes: this.#original.slice(), warnings: [] }
            : await this.#materialize("save", format, signal);
        throwIfAborted(signal);
        return Object.freeze({
          bytes,
          stateToken: this.#stateToken(this.#history.stateId),
          sessionId: this.sessionId,
          revision: this.#revision,
          warnings: Object.freeze([...warnings]),
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

  getElements(
    query: ElementQuery = {},
    options: ReadOptions = {},
  ): Promise<ReadResult<EditElement>> {
    return this.#enqueue(options.signal, async (signal) =>
      this.#items(await this.#engine.getElements(query, signal)),
    );
  }

  getElement(
    id: string,
    options: ReadOptions = {},
  ): Promise<ReadItem<EditElement>> {
    return this.#enqueue(options.signal, async (signal) => {
      const item = this.#engine.getElement
        ? await this.#engine.getElement(id, signal)
        : (await this.#engine.getElements({}, signal)).find(
            (element) => element.id === id,
          );
      return Object.freeze({
        sessionId: this.sessionId,
        revision: this.#revision,
        item,
      });
    });
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
    options: ReadOptions = {},
  ): Promise<ReadResult<EditElement>> {
    return this.#enqueue(options.signal, async (signal) =>
      this.#items(await this.#engine.elementsAt(pageIndex, point, signal)),
    );
  }

  findText(
    query: string,
    options: EditFindOptions = {},
  ): Promise<ReadResult<TextTarget>> {
    // The signal stays on this side; the rest may cross to a worker.
    const { signal: own, ...engineOptions } = options;
    return this.#enqueue(own, async (signal) =>
      this.#items(await this.#engine.findText(query, engineOptions, signal)),
    );
  }

  getOutline(options?: OutlineOptions): Promise<OutlineResult> {
    return readOutline(this, this.#host.limits, options);
  }

  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>> {
    return readDescription(this, this.#host.limits, options);
  }

  resolveTargets(
    query: TargetQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<TargetCandidate>> {
    return resolveTargets(this, query, options);
  }

  createCheckpoint(label?: string): Promise<EditCheckpoint> {
    return this.#enqueue(undefined, async () => {
      const limit = this.#host.limits.maxEditCheckpoints;
      if (this.#named.size >= limit)
        throw new ViewerError(
          "resource-limit",
          "Too many edit checkpoints; drop one first",
          { details: { limit } },
        );
      const stateId = this.#history.stateId;
      const checkpoint: EditCheckpoint = Object.freeze({
        id: randomId(),
        ...(label === undefined ? {} : { label }),
        revision: this.#revision,
        createdAt: new Date().toISOString(),
      });
      const entries = this.#history.entriesAt(this.#history.position);
      this.#named.set(checkpoint.id, {
        checkpoint,
        stateId,
        pageCount: this.#history.pageCount,
        entries,
        batches: batchesFromOriginal(entries),
      });
      this.#pin(stateId, this.#committedBytes);
      return checkpoint;
    });
  }

  listCheckpoints(): readonly EditCheckpoint[] {
    return Object.freeze(
      [...this.#named.values()].map((named) => named.checkpoint),
    );
  }

  restoreCheckpoint(
    id: string,
    options: HistoryOptions = {},
  ): Promise<EditReceipt> {
    return this.#enqueue(options.signal, async (signal) => {
      this.#assertRevision(options);
      const named = this.#named.get(id);
      if (!named)
        throw new ViewerError(
          "invalid-operation",
          `Unknown edit checkpoint ${id}`,
          { details: { checkpointId: id } },
        );
      if (named.stateId === this.#history.stateId) return this.#noop();
      const before = this.#history.pageCount;
      const changedPages = allPages(Math.max(before, named.pageCount));
      const shown = await this.#transaction(signal, "apply", async () => {
        await this.#engine.restore(this.#targetFor(named.entries), signal);
        return this.#show(signal, changedPages);
      });
      const diff = entryDiff(
        this.#history.entriesAt(this.#history.position),
        named.entries,
      );
      this.#history.push(
        {
          operations: [],
          createdIds: diff.created,
          removedIds: diff.removed,
          changedPages: shown.changedPages,
          pageCountBefore: before,
          pageCountAfter: shown.pageCount,
          base: { stateId: named.stateId, batches: named.batches },
        },
        named.stateId,
      );
      this.#commit("restore", shown.changedPages, shown);
      return this.#receipt(false, 0, diff.created, {
        removedIds: diff.removed,
        changedPages: shown.changedPages,
        pageCount: shown.pageCount,
        warnings: [],
      });
    });
  }

  dropCheckpoint(id: string): void {
    const named = this.#named.get(id);
    if (!named) return;
    this.#named.delete(id);
    this.#unpin(named.stateId);
  }

  get tools(): ToolSet {
    return (this.#tools ??= buildToolSet(this.format, this.schemas));
  }

  callTool(call: ToolCall, options?: ToolCallOptions): Promise<ToolResult> {
    return runTool(this, this.tools, call, options);
  }

  readItem<T>(
    options: ReadOptions | undefined,
    task: (engine: EditEngine, signal: AbortSignal) => Promise<T | undefined>,
  ): Promise<ReadItem<T>> {
    return this.#enqueue(options?.signal, async (signal) =>
      Object.freeze({
        sessionId: this.sessionId,
        revision: this.#revision,
        item: await task(this.#engine, signal),
      }),
    );
  }

  readItems<T>(
    options: ReadOptions | undefined,
    task: (engine: EditEngine, signal: AbortSignal) => Promise<readonly T[]>,
  ): Promise<ReadResult<T>> {
    return this.#enqueue(options?.signal, async (signal) =>
      this.#items(await task(this.#engine, signal)),
    );
  }

  /** Stamps a read with the state it describes: the revision at its queue position. */
  #items<T>(items: readonly T[]): ReadResult<T> {
    return Object.freeze({
      sessionId: this.sessionId,
      revision: this.#revision,
      items: Object.freeze([...items]),
    });
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
      if (
        error instanceof ViewerError &&
        (error.code === "edit-failed" || error.code === "invalid-operation")
      )
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
  /** The engine's bytes with their warnings; an engine without the richer form warns of nothing. */
  async #materialize(
    purpose: "show" | "save",
    options: MaterializeOptions,
    signal: AbortSignal,
  ): Promise<MaterializedDocument> {
    if (this.#engine.materializeDocument)
      return this.#engine.materializeDocument(purpose, options, signal);
    return {
      bytes: await this.#engine.materialize(purpose, options, signal),
      warnings: [],
    };
  }

  /**
   * The pages a change repaints: the engine's list, or for a reflow every
   * page from the paragraph's first page (asked of the host before the
   * document is replaced) to the end of the shown document.
   */
  #pagesOf(
    changedPages: readonly number[],
    reflowFrom: string | undefined,
  ): ChangedPages {
    if (reflowFrom === undefined) return changedPages;
    // Everything from the reflowed paragraph's page on, and the pages
    // the batch names before it (an undo that moves a paragraph back).
    const first = Math.min(
      this.#host.pageOf?.(reflowFrom) ?? 0,
      changedPages[0] ?? Number.POSITIVE_INFINITY,
    );
    return (pageCount) =>
      Array.from(
        { length: Math.max(0, pageCount - first) },
        (_, index) => first + index,
      );
  }

  async #show(signal: AbortSignal, changedPages: ChangedPages): Promise<Shown> {
    let bytes: Uint8Array;
    try {
      ({ bytes } = await this.#materialize("show", {}, signal));
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
    const pages = Object.freeze(
      typeof changedPages === "function"
        ? [...changedPages(prepared.pageCount)]
        : [...changedPages],
    );
    return {
      bytes,
      pageCount: this.#host.commitDocument(prepared, pages),
      changedPages: pages,
    };
  }

  async #moveTo(
    position: number,
    changedPages: ChangedPages,
    signal: AbortSignal,
  ): Promise<Shown> {
    return this.#transaction(signal, "apply", async () => {
      await this.#engine.restore(this.#restoreTarget(position), signal);
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
      await this.#engine.restore(this.#restoreTarget(), controller.signal);
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
    if (reason === "apply" || reason === "restore")
      this.#keepCheckpoint(shown.bytes);
    else if (reason === "reset") this.#dropCheckpoints();
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

  addAsset(data: Uint8Array, options: AssetOptions = {}): Promise<string> {
    return this.#enqueue(options.signal, async (signal) => {
      if (data.byteLength > this.#host.limits.maxInputBytes)
        throw new ViewerError("resource-limit", "Asset exceeds maxInputBytes", {
          details: {
            actual: data.byteLength,
            limit: this.#host.limits.maxInputBytes,
          },
        });
      return this.#register(data.slice(), signal);
    });
  }

  /** Stores bytes under their content id and hands them to the engine once. */
  async #register(bytes: Uint8Array, signal: AbortSignal): Promise<string> {
    const id = await assetIdOf(bytes);
    if (!this.#assets.has(id)) {
      this.#assets.set(id, bytes);
      await this.#engine.putAsset(id, bytes, signal);
    }
    return id;
  }

  /**
   * Replaces inline binary payloads by asset references, so the history
   * holds references only; unknown references are reported like any issue.
   */
  async #intern(
    batch: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly EditOperation[]> {
    const issues: OperationIssue[] = [];
    const result: EditOperation[] = [];
    for (const [operationIndex, operation] of batch.entries()) {
      const fields = binaryFields(this.schemas.operations[operation.op]);
      if (fields.length === 0) {
        result.push(operation);
        continue;
      }
      const patched: Record<string, unknown> = { ...operation };
      for (const field of fields) {
        const value = patched[field];
        if (value === undefined) continue;
        if (isAssetReference(value)) {
          if (!this.#assets.has(value))
            issues.push({
              operationIndex,
              path: `/${field}`,
              code: "unknown-asset",
              message: `Unknown asset ${value}`,
            });
          continue;
        }
        const bytes =
          value instanceof Uint8Array
            ? value
            : Uint8Array.from(atob(value as string), (c) => c.charCodeAt(0));
        patched[field] = await this.#register(bytes, signal);
      }
      result.push(Object.freeze(patched) as unknown as EditOperation);
    }
    if (issues.length > 0) throw invalidOperationError(issues);
    return Object.freeze(result);
  }

  /** Keeps every stride-th committed state's bytes within the memory budget. */
  #keepCheckpoint(bytes: Uint8Array): void {
    const stride = Math.max(
      1,
      Math.floor(this.#host.limits.maxEditHistory / 4),
    );
    const stateId = this.#history.stateId;
    // Entries dropped by a new change after an undo can never be restored;
    // a pinned state stays whatever the history does.
    const reachable = this.#reachable();
    for (const [id, kept] of this.#checkpoints)
      if (!reachable.has(id) && !this.#pinned.has(id))
        this.#forgetCheckpoint(id, kept);
    if (
      stateId === 0 ||
      this.#checkpoints.has(stateId) ||
      stateId % stride !== 0
    )
      return;
    this.#retain(stateId, bytes);
  }

  /** Stores a state's bytes when the budget, less the pinned states, can hold them. */
  #retain(stateId: number, bytes: Uint8Array): void {
    const budget = this.#host.limits.maxEditCheckpointBytes;
    if (bytes.byteLength > budget) return;
    const evictable = [...this.#checkpoints.keys()]
      .filter((id) => !this.#pinned.has(id))
      .sort((a, b) => a - b);
    while (
      this.#checkpointBytes + bytes.byteLength > budget &&
      evictable.length > 0
    ) {
      const id = evictable.shift()!;
      this.#forgetCheckpoint(id, this.#checkpoints.get(id)!);
    }
    if (this.#checkpointBytes + bytes.byteLength > budget) return;
    this.#checkpoints.set(stateId, bytes);
    this.#checkpointBytes += bytes.byteLength;
  }

  /**
   * Pins a state for a named checkpoint: its bytes are kept while any
   * checkpoint names it, or rebuilt by replay when the budget cannot hold
   * them. The original (state 0) needs no bytes.
   */
  #pin(stateId: number, bytes: Uint8Array): void {
    if (stateId === 0) return;
    this.#pinned.set(stateId, (this.#pinned.get(stateId) ?? 0) + 1);
    if (!this.#checkpoints.has(stateId)) this.#retain(stateId, bytes);
  }

  #unpin(stateId: number): void {
    const count = this.#pinned.get(stateId);
    if (!count) return;
    if (count > 1) {
      this.#pinned.set(stateId, count - 1);
      return;
    }
    this.#pinned.delete(stateId);
    const kept = this.#checkpoints.get(stateId);
    if (kept && !this.#reachable().has(stateId))
      this.#forgetCheckpoint(stateId, kept);
  }

  /** State ids a replay may start from: every entry's, and the base of every restore. */
  #reachable(): Set<number> {
    const reachable = new Set<number>();
    for (const entry of this.#history.allEntries) {
      reachable.add(entry.stateId);
      if (entry.base) reachable.add(entry.base.stateId);
    }
    return reachable;
  }

  #forgetCheckpoint(id: number, bytes: Uint8Array): void {
    this.#checkpoints.delete(id);
    this.#checkpointBytes -= bytes.byteLength;
  }

  /** After a reset nothing in the history is reachable; pinned states stay. */
  #dropCheckpoints(): void {
    for (const [id, kept] of this.#checkpoints)
      if (!this.#pinned.has(id)) this.#forgetCheckpoint(id, kept);
  }

  /**
   * The cheapest way to rebuild the state at `position`: the newest
   * retained bytes at or before it, plus the batches after them.
   */
  #restoreTarget(position = this.#history.position): RestoreTarget {
    return this.#targetFor(this.#history.entriesAt(position));
  }

  /**
   * Walks `entries` from the end: the first with retained bytes is the base;
   * a restore entry without them starts from the original plus the batches
   * that built its checkpoint.
   */
  #targetFor(entries: readonly HistoryEntry[]): RestoreTarget {
    for (let index = entries.length - 1; index >= 0; index -= 1) {
      const entry = entries[index]!;
      const after = (): readonly EngineBatch[] =>
        batchesOf(entries.slice(index + 1));
      const retained = this.#checkpoints.get(entry.stateId);
      if (retained) return { base: retained, batches: after() };
      // A restore entry carries its checkpoint's state id, so its bytes
      // were just looked up; without them the checkpoint's batches rebuild
      // it from the original.
      if (entry.base) return { batches: [...entry.base.batches, ...after()] };
    }
    return { batches: batchesOf(entries) };
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
      Partial<
        Pick<
          EngineChange,
          "removedIds" | "remappedIds" | "textAnchorMigrations"
        >
      > & {
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
      ...(change.textAnchorMigrations?.length
        ? {
            textAnchorMigrations: Object.freeze(
              change.textAnchorMigrations.map((migration) =>
                Object.freeze({ ...migration }),
              ),
            ),
          }
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

/**
 * Same-batch references: `"$<n>"` names the first element created by
 * operation `n` of the batch, which must come earlier. The engine resolves
 * them while applying; here only the form is checked.
 */
function checkBatchReferences(
  operations: readonly EditOperation[],
): OperationIssue[] {
  const issues: OperationIssue[] = [];
  operations.forEach((operation, operationIndex) => {
    const target = (operation as { readonly target?: unknown }).target;
    const reference =
      typeof target === "string" ? parseReference(target) : undefined;
    if (reference === undefined) return;
    if (reference >= operationIndex)
      issues.push({
        operationIndex,
        path: "/target",
        code: "unknown-target",
        message: `"$${reference}" must name an earlier operation of the batch`,
      });
  });
  return issues;
}

/**
 * Tracked changes exist in DOCX only and always name their author (decision
 * 8 of the ai-edit module): the batch is refused before any engine work.
 */
function checkChangeMode(
  format: EditableFormat,
  options: ApplyOptions,
): OperationIssue[] {
  const issues: OperationIssue[] = [];
  // A timestamp is written into the file where a format records one, so
  // it must be a date-time the file can hold whatever the mode.
  if (
    options.timestamp !== undefined &&
    (!DATE_TIME.test(options.timestamp) ||
      Number.isNaN(Date.parse(options.timestamp)))
  )
    issues.push({
      operationIndex: -1,
      path: "/timestamp",
      code: "invalid-value",
      message: "The timestamp must be an ISO 8601 date-time",
    });
  if (options.changeMode !== "tracked") return issues;
  if (format !== "docx")
    issues.push({
      operationIndex: -1,
      path: "",
      code: "unsupported-change-mode",
      message: `${format.toUpperCase()} has no tracked changes; apply directly and review with checkpoints`,
    });
  else if (!options.author || options.author.trim().length === 0)
    issues.push({
      operationIndex: -1,
      path: "/author",
      code: "required",
      message: "Tracked changes name their author; pass ApplyOptions.author",
    });
  return issues;
}

/** ISO 8601 as `xsd:dateTime` takes it. */
const DATE_TIME =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?$/;

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

/** A state the host named: what names it, where it sits and how to rebuild it. */
interface NamedCheckpoint {
  readonly checkpoint: EditCheckpoint;
  readonly stateId: number;
  readonly pageCount: number;
  /** The entries applied in that state, folded ones first; for receipts. */
  readonly entries: readonly HistoryEntry[];
  /** The batches that build the state from the original, restores flattened. */
  readonly batches: readonly EngineBatch[];
}

/** The engine batches of entries that carry operations; restore entries carry none. */
function batchesOf(entries: readonly HistoryEntry[]): readonly EngineBatch[] {
  return entries.filter((entry) => entry.operations.length > 0).map(batchOf);
}

/** The batches from the original to the state after `entries`, through the last restore. */
function batchesFromOriginal(
  entries: readonly HistoryEntry[],
): readonly EngineBatch[] {
  for (let index = entries.length - 1; index >= 0; index -= 1) {
    const base = entries[index]!.base;
    if (base) return [...base.batches, ...batchesOf(entries.slice(index + 1))];
  }
  return batchesOf(entries);
}

/**
 * The ids a move from the state after `current` to the state after `target`
 * removes and creates: the entries past the common prefix are undone in
 * reverse, then the target's are redone, each id netted out.
 */
function entryDiff(
  current: readonly HistoryEntry[],
  target: readonly HistoryEntry[],
): { readonly created: string[]; readonly removed: string[] } {
  let common = 0;
  while (
    common < current.length &&
    common < target.length &&
    current[common]!.stateId === target[common]!.stateId
  )
    common += 1;
  const created = new Set<string>();
  const removed = new Set<string>();
  const create = (id: string): void => {
    if (removed.has(id)) removed.delete(id);
    else created.add(id);
  };
  const remove = (id: string): void => {
    if (created.has(id)) created.delete(id);
    else removed.add(id);
  };
  for (const entry of current.slice(common).reverse()) {
    entry.createdIds.forEach(remove);
    entry.removedIds.forEach(create);
  }
  for (const entry of target.slice(common)) {
    entry.removedIds.forEach(remove);
    entry.createdIds.forEach(create);
  }
  return { created: [...created], removed: [...removed] };
}

/** 128 random bits as 22 URL-safe base64 characters; unique across sessions and reloads. */
function randomId(): string {
  const bytes = globalThis.crypto.getRandomValues(new Uint8Array(16));
  let binary = "";
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary)
    .replaceAll("+", "-")
    .replaceAll("/", "_")
    .replace(/=+$/, "");
}

/** Surfaces a listener's exception without failing the call that emitted the event. */
export function reportError(error: unknown): void {
  const report = (globalThis as { reportError?: (error: unknown) => void })
    .reportError;
  if (report) report(error);
  else console.error(error);
}
