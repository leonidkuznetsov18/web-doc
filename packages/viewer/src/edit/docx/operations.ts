import type { ResourceLimits, ViewerWarning } from "../../contracts.js";
import type { AssetSource } from "../assets.js";
import type { OoxmlPackage } from "../ooxml/package.js";
import type { OperationIssue } from "../types.js";
import type { DocxModel } from "./model.js";
import type { DocxOperation } from "./types.js";

/*
 * How an operation sees the document while validating and applying. Each
 * operation runs in its own package transaction; the engine snapshots the
 * package before a batch and restores it when any operation fails.
 */

export interface DocxOperationContext {
  readonly pkg: OoxmlPackage;
  readonly model: DocxModel;
  readonly limits: ResourceLimits;
  /** Bytes behind `asset:` references. */
  readonly assets: AssetSource;
  /** The batch's state id, for ids an operation derives. */
  readonly stateId: number;
  readonly operationIndex: number;
}

export interface DocxOperationResult {
  readonly createdIds: readonly string[];
  readonly removedIds?: readonly string[];
  readonly warnings: readonly ViewerWarning[];
}

export type Issue = (path: string, code: string, message: string) => void;

export interface DocxOperationHandler<T extends DocxOperation = DocxOperation> {
  validate(
    operation: T,
    context: DocxOperationContext,
    issue: Issue,
  ): Promise<void>;
  apply(
    operation: T,
    context: DocxOperationContext,
  ): Promise<DocxOperationResult>;
}

export function issueCollector(
  operationIndex: number,
  issues: OperationIssue[],
): Issue {
  return (path, code, message) =>
    issues.push({ operationIndex, path, code, message });
}
