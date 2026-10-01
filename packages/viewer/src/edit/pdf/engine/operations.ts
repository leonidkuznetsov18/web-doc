import type { ViewerWarning } from "../../../contracts.js";
import type { OperationIssue } from "../../types.js";
import type { PdfOperation } from "../types.js";
import type { PageGeometry } from "./geometry.js";
import type { ObjectRecord } from "./elements.js";
import type { Pdfium } from "./pdfium.js";
import type { TextMeasurer } from "./fonts.js";

/** What an operation sees of the document while validating or applying. */
export interface OperationContext {
  readonly pdfium: Pdfium;
  readonly document: number;
  readonly measurer: TextMeasurer;
  readonly pageCount: number;
  /** Geometry of a page; loads it if needed. */
  geometry(pageIndex: number): PageGeometry;
  /** Element id for something this operation creates. */
  newId(pageIndex: number, suffix?: string): string;
  /** Loads a page for writing; its content is regenerated afterwards. */
  withPage<T>(pageIndex: number, use: (page: number) => T): T;
  /** Records objects appended to a page by this operation. */
  appendObjects(pageIndex: number, records: readonly ObjectRecord[]): void;
}

export interface OperationResult {
  readonly createdIds: readonly string[];
  readonly changedPages: readonly number[];
  readonly warnings: readonly ViewerWarning[];
}

export interface OperationHandler<T extends PdfOperation = PdfOperation> {
  validate(
    operation: T,
    context: OperationContext,
    issue: (path: string, code: string, message: string) => void,
  ): void;
  apply(operation: T, context: OperationContext): OperationResult;
}

export type Issue = (path: string, code: string, message: string) => void;

export function issueCollector(
  operationIndex: number,
  issues: OperationIssue[],
): Issue {
  return (path, code, message) =>
    issues.push({ operationIndex, path, code, message });
}
