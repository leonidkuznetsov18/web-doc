import type { EditOperation } from "./types.js";

export interface HistoryEntry {
  readonly operations: readonly EditOperation[];
  readonly label?: string;
  /** Page indexes the batch changed, in the document after it. */
  readonly changedPages: readonly number[];
  readonly pageCountBefore: number;
  readonly pageCountAfter: number;
  /** Identifies the content after this batch; equal ids mean equal content. */
  readonly stateId: number;
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

  /** The entry `undo()` would revert, if any. */
  get undoEntry(): HistoryEntry | undefined {
    return this.#position > 0 ? this.#entries[this.#position - 1] : undefined;
  }

  /** The entry `redo()` would re-apply, if any. */
  get redoEntry(): HistoryEntry | undefined {
    return this.#entries[this.#position];
  }

  /** Batches applied to the original in the current state. */
  applied(): readonly (readonly EditOperation[])[] {
    return this.batchesAt(this.#position);
  }

  /** Batches applied to the original when `position` undoable entries are applied. */
  batchesAt(position: number): readonly (readonly EditOperation[])[] {
    return [...this.#folded, ...this.#entries.slice(0, position)].map(
      (entry) => entry.operations,
    );
  }

  get position(): number {
    return this.#position;
  }

  push(entry: Omit<HistoryEntry, "stateId">): HistoryEntry {
    const stored: HistoryEntry = Object.freeze({
      ...entry,
      stateId: this.#nextStateId++,
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
