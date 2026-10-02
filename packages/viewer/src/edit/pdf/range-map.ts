import type {
  EditOperation,
  EditReceipt,
  TextPosition,
  TextRange,
} from "../types.js";

/*
 * Where a text range goes after the session's later mutations. The PDF
 * session records every committed call — a batch, an undo, a redo, a reset —
 * with the operations involved and the receipt; a range taken at an earlier
 * revision is carried through those records: `replaceText` shifts offsets,
 * deletions and page removals make the element gone, undos bring it back,
 * and renamed ids are followed.
 */

export interface MutationRecord {
  /** The revision the call produced. */
  readonly revision: number;
  readonly kind: "apply" | "undo" | "redo" | "reset";
  /**
   * The operations the call applied (apply, redo) or took back (undo, reset);
   * for a reset every batch of the session, in order.
   */
  readonly operations: readonly EditOperation[];
  readonly receipt: EditReceipt;
}

interface ReplaceLike {
  readonly op: "replaceText";
  readonly target: string;
  readonly text: string;
  readonly range?: TextRange;
}

interface Carried {
  position: TextPosition;
  gone: boolean;
}

/** The range after the records, or `undefined` when its element is gone. */
export function mapRangeThrough(
  range: TextRange,
  records: readonly MutationRecord[],
): TextRange | undefined {
  const start = carry(range.start, records);
  const end = carry(range.end, records);
  if (start.gone || end.gone) return undefined;
  if (
    start.position.elementId === end.position.elementId &&
    end.position.offset < start.position.offset
  )
    return { start: start.position, end: start.position };
  return { start: start.position, end: end.position };
}

function carry(
  position: TextPosition,
  records: readonly MutationRecord[],
): Carried {
  const carried: Carried = { position, gone: false };
  for (const record of records) {
    const forward = record.kind === "apply" || record.kind === "redo";
    const operations = forward
      ? record.operations
      : [...record.operations].reverse();
    for (const operation of operations) {
      if (forward) applyForward(carried, operation);
      else applyBackward(carried, operation);
    }
    const { receipt } = record;
    const id = carried.position.elementId;
    if (receipt.removedIds.includes(id)) carried.gone = true;
    if (receipt.createdIds.includes(id)) carried.gone = false;
    const renamed = receipt.remappedIds?.[id];
    if (renamed) carried.position = { ...carried.position, elementId: renamed };
  }
  return carried;
}

function applyForward(carried: Carried, operation: EditOperation): void {
  const id = carried.position.elementId;
  if (operation.op === "deleteElement" && targetOf(operation) === id) {
    carried.gone = true;
    return;
  }
  const replace = asReplace(operation, id);
  if (!replace) return;
  const { offset } = carried.position;
  const span = spanOf(replace);
  // A whole-text replacement: offsets stay and the live element bounds them.
  if (!span) return;
  const delta = replace.text.length - (span.end - span.start);
  carried.position = {
    elementId: id,
    offset:
      offset <= span.start
        ? offset
        : offset >= span.end
          ? offset + delta
          : Math.min(offset, span.start + replace.text.length),
  };
}

function applyBackward(carried: Carried, operation: EditOperation): void {
  const id = carried.position.elementId;
  if (operation.op === "deleteElement" && targetOf(operation) === id) {
    carried.gone = false;
    return;
  }
  const replace = asReplace(operation, id);
  if (!replace) return;
  const span = spanOf(replace);
  // A whole-text replacement taken back: the old length is unknown here.
  if (!span) return;
  const { offset } = carried.position;
  const newEnd = span.start + replace.text.length;
  const delta = span.end - span.start - replace.text.length;
  carried.position = {
    elementId: id,
    offset:
      offset <= span.start
        ? offset
        : offset >= newEnd
          ? offset + delta
          : Math.min(offset, span.end),
  };
}

function asReplace(
  operation: EditOperation,
  elementId: string,
): ReplaceLike | undefined {
  if (operation.op !== "replaceText") return undefined;
  const candidate = operation as Partial<ReplaceLike>;
  if (candidate.target !== elementId || typeof candidate.text !== "string")
    return undefined;
  return candidate as ReplaceLike;
}

function spanOf(
  replace: ReplaceLike,
): { readonly start: number; readonly end: number } | undefined {
  const { range } = replace;
  if (
    !range ||
    range.start.elementId !== replace.target ||
    range.end.elementId !== replace.target
  )
    return undefined;
  return {
    start: Math.min(range.start.offset, range.end.offset),
    end: Math.max(range.start.offset, range.end.offset),
  };
}

function targetOf(operation: EditOperation): unknown {
  return (operation as { readonly target?: unknown }).target;
}
