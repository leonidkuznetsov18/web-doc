import type { ResourceLimits, ViewerWarning } from "../../../contracts.js";
import type { ImageCache } from "./images.js";
import type { OperationIssue } from "../../types.js";
import type { PdfElement, PdfOperation } from "../types.js";
import type { PageGeometry } from "./geometry.js";
import type { ObjectRecord } from "./elements.js";
import type { Pdfium } from "./pdfium.js";
import type { FontLibrary, TextMeasurer } from "./fonts.js";

/** What an operation sees of the document while validating or applying. */
export interface OperationContext {
  readonly pdfium: Pdfium;
  readonly document: number;
  readonly measurer: TextMeasurer;
  readonly fonts: FontLibrary;
  readonly images: ImageCache;
  readonly limits: ResourceLimits;
  readonly pageCount: number;
  /** Geometry of a page; loads it if needed. */
  geometry(pageIndex: number): PageGeometry;
  /** Element id for something this operation creates. */
  newId(pageIndex: number, suffix?: string): string;
  /** Loads a page for writing; its content is regenerated afterwards. */
  withPage<T>(pageIndex: number, use: (page: number) => T): T;
  /** Records objects appended to a page by this operation. */
  appendObjects(pageIndex: number, records: readonly ObjectRecord[]): void;
  /** Where an element's objects sit: their page and their indexes in drawing order. */
  locate(id: string): ElementLocation | undefined;
  /** The element as a query would return it. */
  element(id: string): PdfElement | undefined;
  /** Ids of every element on a page, for what a page deletion removes. */
  pageElementIds(pageIndex: number): readonly string[];
  /** Displayed size of a page in points. */
  pageSize(pageIndex: number): {
    readonly width: number;
    readonly height: number;
  };
  /** Records a page created at `index`; returns its stable key. */
  insertPageRecord(index: number): string;
  removePageRecord(index: number): void;
  movePageRecord(from: number, to: number): void;
  /** Drops what the model cached about a page after its geometry changed. */
  invalidatePage(index: number): void;
  /** Replaces `count` object records from `start` with `records`. */
  spliceObjects(
    pageIndex: number,
    start: number,
    count: number,
    records: readonly ObjectRecord[],
  ): void;
}

export interface ElementLocation {
  readonly pageIndex: number;
  /** Indexes of the element's objects in the page's drawing order, ascending. */
  readonly indexes: readonly number[];
  readonly record: ObjectRecord;
}

export interface OperationResult {
  readonly createdIds: readonly string[];
  /** Ids the operation made disappear. */
  readonly removedIds?: readonly string[];
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
