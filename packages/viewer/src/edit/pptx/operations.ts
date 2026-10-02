import type { ResourceLimits, ViewerWarning } from "../../contracts.js";
import type { AssetSource } from "../assets.js";
import type { OoxmlPackage } from "../ooxml/package.js";
import type { CommittedChange } from "../ooxml/transaction.js";
import type { OperationIssue } from "../types.js";
import type { ShapeRecord, SlideElements } from "./elements.js";
import type { DeckModel } from "./model.js";
import type { PptxOperation } from "./types.js";

/*
 * How an operation sees the deck while validating and applying. Each
 * operation runs in its own package transaction; the engine snapshots the
 * package before a batch and restores it when any operation fails.
 */

export interface PptxOperationContext {
  readonly pkg: OoxmlPackage;
  readonly model: DeckModel;
  readonly limits: ResourceLimits;
  /** Bytes behind `asset:` references. */
  readonly assets: AssetSource;
  /** Slides after the operations before this one in the batch; the model's count when applying. */
  readonly pageCount: number;
  /** The batch's state id, for ids an operation derives. */
  readonly stateId: number;
  readonly operationIndex: number;
  /** Every element of a slide, read at the current revision. */
  elements(pageIndex: number): Promise<SlideElements>;
  /** The element an id names, at the current revision. */
  locate(id: string): Promise<ShapeRecord | undefined>;
  /**
   * A `p:cNvPr` id for an element created on a slide: above every id the
   * slide holds and above every id the session already issued there, so
   * an id is never reused after a deletion.
   */
  allocateShapeId(elements: SlideElements): number;
}

export interface PptxOperationResult {
  readonly createdIds: readonly string[];
  readonly removedIds?: readonly string[];
  readonly changedPages: readonly number[];
  readonly warnings: readonly ViewerWarning[];
  /** Parts the operation's transaction changed, added and removed; absent means "everything may have". */
  readonly parts?: {
    readonly changed: readonly string[];
    readonly added: readonly string[];
    readonly removed: readonly string[];
  };
}

export type Issue = (path: string, code: string, message: string) => void;

export interface PptxOperationHandler<T extends PptxOperation = PptxOperation> {
  validate(
    operation: T,
    context: PptxOperationContext,
    issue: Issue,
  ): Promise<void>;
  apply(
    operation: T,
    context: PptxOperationContext,
  ): Promise<PptxOperationResult>;
}

export function issueCollector(
  operationIndex: number,
  issues: OperationIssue[],
): Issue {
  return (path, code, message) =>
    issues.push({ operationIndex, path, code, message });
}

/** The parts a transaction touched, in the shape an operation result carries. */
export function committedParts(
  change: CommittedChange,
): NonNullable<PptxOperationResult["parts"]> {
  return {
    changed: change.changedParts,
    added: change.addedParts,
    removed: change.removedParts,
  };
}
