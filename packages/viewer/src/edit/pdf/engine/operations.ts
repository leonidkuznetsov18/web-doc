import type { ResourceLimits, ViewerWarning } from "../../../contracts.js";
import type { AssetSource } from "../../assets.js";
import type { ImageCache } from "./images.js";
import type { OperationIssue, TextAnchorMigration } from "../../types.js";
import type { PdfElement, PdfOperation } from "../types.js";
import type { PageGeometry } from "./geometry.js";
import type { ObjectRecord } from "./elements.js";
import type { Pdfium } from "./pdfium.js";
import type { FontLibrary, TextMeasurer } from "./fonts.js";
import type { ParagraphTarget } from "./paragraph.js";

/** What an operation sees of the document while validating or applying. */
export interface OperationContext {
  readonly pdfium: Pdfium;
  readonly document: number;
  readonly measurer: TextMeasurer;
  readonly fonts: FontLibrary;
  readonly images: ImageCache;
  readonly limits: ResourceLimits;
  /** Bytes behind `asset:` references. */
  readonly assets: AssetSource;
  readonly pageCount: number;
  /** Geometry of a page; loads it if needed. */
  geometry(pageIndex: number): PageGeometry;
  /** Element id for something this operation creates. */
  newId(pageIndex: number, suffix?: string): string;
  /** Loads a page for writing; its content is regenerated afterwards. */
  withPage<T>(pageIndex: number, use: (page: number) => T): T;
  /** Reads native resources without regenerating the content stream. */
  readPage<T>(pageIndex: number, use: (page: number) => T): T;
  /**
   * Loads what holds an element's objects for writing: its page, or for an
   * element inside forms a stand-in page with the innermost form's objects,
   * whose forms are rewritten afterwards (see forms.ts). Indexes in
   * `location.indexes` address objects of the holder.
   */
  withHolder<T>(location: ElementLocation, use: (holder: number) => T): T;
  /** Reads an object of an element's holder, with the page's text, changing nothing. */
  readObject<T>(
    location: ElementLocation,
    index: number,
    use: (object: number, textPage: number) => T,
  ): T;
  /** The page's geometry for objects of the element's holder; see `PageGeometry.matrix`. */
  holderGeometry(location: ElementLocation): PageGeometry;
  /**
   * Whether the forms that hold an element can be rewritten without the
   * page looking any different; always true on the page itself.
   */
  rewritable(location: ElementLocation): boolean;
  /** Records objects appended to a page by this operation. */
  appendObjects(pageIndex: number, records: readonly ObjectRecord[]): void;
  /** Where an element's objects sit: their page and their indexes in drawing order. */
  locate(id: string): ElementLocation | undefined;
  /** The element as a query would return it. */
  element(id: string): PdfElement | undefined;
  paragraph(id: string): ParagraphTarget | undefined;
  pageElements(pageIndex: number): readonly PdfElement[];
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
  /** Replaces `count` object records from `start` with `records`, inside `forms` if given. */
  spliceObjects(
    pageIndex: number,
    start: number,
    count: number,
    records: readonly ObjectRecord[],
    forms?: readonly number[],
  ): void;
}

export interface ElementLocation {
  readonly pageIndex: number;
  /** Indexes of the forms that hold the element, from the page inwards; empty on the page. */
  readonly forms: readonly number[];
  /** Indexes of the element's objects in its holder's drawing order, ascending. */
  readonly indexes: readonly number[];
  readonly record: ObjectRecord;
}

export interface OperationResult {
  readonly createdIds: readonly string[];
  /** Ids the operation made disappear. */
  readonly removedIds?: readonly string[];
  readonly textAnchorMigrations?: readonly Omit<
    TextAnchorMigration,
    "operationIndex"
  >[];
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

/**
 * Text inside forms is written by rewriting its forms, which must leave the
 * page looking as it did; reports the target when it would not.
 */
export function rewritable(
  location: ElementLocation,
  context: OperationContext,
  issue: Issue,
): boolean {
  if (context.rewritable(location)) return true;
  issue(
    "/target",
    "unsupported-target",
    "This text sits in a form web-doc cannot rewrite without changing how the page looks",
  );
  return false;
}
