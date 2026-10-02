import type { BatchMode, EngineBatch } from "./engine.js";
import type { EditOperation } from "./types.js";

export interface HistoryEntry extends BatchMode {
  readonly operations: readonly EditOperation[];
  readonly label?: string;
  /** Ids the batch created; an undo removes them again. */
  readonly createdIds: readonly string[];
  /** Ids the batch removed; a redo removes them again. */
  readonly removedIds: readonly string[];
  /** Page indexes the batch changed, in the document after it. */
  readonly changedPages: readonly number[];
  /** Flow formats: the paragraph the document reflows from, for undo and redo. */
  readonly reflowFrom?: string;
  readonly pageCountBefore: number;
  readonly pageCountAfter: number;
  /** Identifies the content after this batch; equal ids mean equal content. */
  readonly stateId: number;
  /**
   * A restore to a named checkpoint: the content after this entry is the
   * checkpoint's, so a replay starts from its retained bytes or from the
   * original plus `batches`, never from the entries before.
   */
  readonly base?: HistoryBase;
}

/** Where a restore entry's content comes from. */
export interface HistoryBase {
  /** The checkpoint's state id; its bytes may be retained under it. */
  readonly stateId: number;
  /** The batches that build the checkpoint's state from the original. */
  readonly batches: readonly EngineBatch[];
}

/**
 * Linear undo history over an immutable original. Entries beyond the limit are
 * folded into the starting point: they stay applied but can no longer be
 * undone. Every reachable content state has a numeric id, so "unchanged since
 * the last save" is an id comparison.
 */
export class EditHistory {
  readonly #limit: number;
  readonly #originalPageCount: number;
  /** Folded batches, always applied before the undoable entries. */
  #folded: HistoryEntry[] = [];
  #entries: HistoryEntry[] = [];
  #position = 0;
  #nextStateId = 1;

  constructor(limit: number, originalPageCount: number) {
    this.#limit = limit;
    this.#originalPageCount = originalPageCount;
  }

  get canUndo(): boolean {
    return this.#position > 0;
  }

  get canRedo(): boolean {
    return this.#position < this.#entries.length;
  }

  get isPristine(): boolean {
    return this.#folded.length === 0 && this.#entries.length === 0;
  }

  /** The id the next pushed entry will get; engines derive created ids from it. */
  get nextStateId(): number {
    return this.#nextStateId;
  }

  /** Content id of the current state; 0 is the original document. */
  get stateId(): number {
    return this.#position > 0
      ? this.#entries[this.#position - 1]!.stateId
      : (this.#folded.at(-1)?.stateId ?? 0);
  }

  get pageCount(): number {
    return this.#position > 0
      ? this.#entries[this.#position - 1]!.pageCountAfter
      : (this.#folded.at(-1)?.pageCountAfter ?? this.#originalPageCount);
  }

  /** State ids of every entry still in the history, folded and redo tail included. */
  get stateIds(): readonly number[] {
    return this.allEntries.map((entry) => entry.stateId);
  }

  /** Every entry still in the history: folded ones, then the undoable ones and the redo tail. */
  get allEntries(): readonly HistoryEntry[] {
    return [...this.#folded, ...this.#entries];
  }

  /** The entry `undo()` would revert, if any. */
  get undoEntry(): HistoryEntry | undefined {
    return this.#position > 0 ? this.#entries[this.#position - 1] : undefined;
  }

  /** The entry `redo()` would re-apply, if any. */
  get redoEntry(): HistoryEntry | undefined {
    return this.#entries[this.#position];
  }

  /** Batches applied to the original in the current state. */
  applied(): readonly EngineBatch[] {
    return this.batchesAt(this.#position);
  }

  /** Batches applied to the original when `position` undoable entries are applied. */
  batchesAt(position: number): readonly EngineBatch[] {
    return this.entriesAt(position).map(batchOf);
  }

  /** Entries applied when `position` undoable entries are applied, folded ones first. */
  entriesAt(position: number): readonly HistoryEntry[] {
    return [...this.#folded, ...this.#entries.slice(0, position)];
  }

  get position(): number {
    return this.#position;
  }

  /**
   * Appends an entry with a fresh state id, or with `stateId` when the entry
   * reproduces a known state (a restore to a checkpoint): the same id means
   * the same content, so `dirty` stays exact.
   */
  push(entry: Omit<HistoryEntry, "stateId">, stateId?: number): HistoryEntry {
    const stored: HistoryEntry = Object.freeze({
      ...entry,
      stateId: stateId ?? this.#nextStateId++,
    });
    // A new change after an undo drops the redo tail.
    this.#entries.length = this.#position;
    this.#entries.push(stored);
    this.#position = this.#entries.length;
    while (this.#entries.length > this.#limit) {
      this.#folded.push(this.#entries.shift()!);
      this.#position -= 1;
    }
    return stored;
  }

  undo(): void {
    if (this.canUndo) this.#position -= 1;
  }

  redo(): void {
    if (this.canRedo) this.#position += 1;
  }

  clear(): void {
    this.#folded = [];
    this.#entries = [];
    this.#position = 0;
  }
}

/** The engine batch an entry replays as: its operations and how they were written. */
export function batchOf(entry: HistoryEntry): EngineBatch {
  return {
    stateId: entry.stateId,
    operations: entry.operations,
    ...modeOf(entry),
  };
}

/** The write mode fields of a batch or an entry, only those set. */
export function modeOf(mode: BatchMode): BatchMode {
  return {
    ...(mode.changeMode === undefined ? {} : { changeMode: mode.changeMode }),
    ...(mode.author === undefined ? {} : { author: mode.author }),
    ...(mode.timestamp === undefined ? {} : { timestamp: mode.timestamp }),
  };
}
