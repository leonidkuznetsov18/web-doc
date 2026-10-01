import type { ViewerWarning } from "../contracts.js";

/** Formats an edit session can be started for. */
export type EditableFormat = "pdf" | "pptx" | "docx";

export interface EditOptions {
  readonly signal?: AbortSignal;
}

/** A plain JSON object; each format defines its concrete operation union. */
export interface EditOperation {
  readonly op: string;
}

/** Binary payload: bytes, or a base64 string for pure-JSON transports. */
export type BinaryData = Uint8Array | string;

export type JsonSchema = Readonly<Record<string, unknown>>;

export interface OperationSchemaSet {
  readonly format: EditableFormat;
  /** Raised whenever an operation's shape changes incompatibly. */
  readonly version: number;
  /** One JSON Schema (draft 2020-12) per operation name. */
  readonly operations: Readonly<Record<string, JsonSchema>>;
}

/**
 * A point in page space: the units of `DocumentInfo.pageSizes` at zoom 1
 * (points for PDF, CSS pixels for DOCX and PPTX), origin at the top-left of
 * the page as displayed, `y` growing downwards, page rotation applied.
 */
export interface PagePoint {
  readonly x: number;
  readonly y: number;
}

/** A rectangle in page space; see `PagePoint`. */
export interface PageRect {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface PageHit {
  readonly pageIndex: number;
  readonly point: PagePoint;
}

/** Rectangle in client (CSS pixel) coordinates, like `getBoundingClientRect()`. */
export interface ViewportRect {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

export interface EditElement {
  /** Opaque, stable for the lifetime of the session. */
  readonly id: string;
  /** Format-specific kind, for example "text", "image", "shape", "table". */
  readonly kind: string;
  readonly pageIndex: number;
  /** Axis-aligned bounds in page space. */
  readonly bounds: PageRect;
  /** Clockwise rotation in degrees, when the element is rotated. */
  readonly rotation?: number;
  readonly text?: string;
  readonly parentId?: string;
  /** Names of the operations that accept this element as their target. */
  readonly operations: readonly string[];
}

export interface ElementQuery {
  readonly pageIndex?: number;
  readonly kinds?: readonly string[];
  /** Only elements whose bounds intersect this rectangle; requires `pageIndex`. */
  readonly intersects?: PageRect;
}

export interface EditFindOptions {
  readonly caseSensitive?: boolean;
  /** Inclusive 0-based page range. */
  readonly pageRange?: readonly [number, number];
  readonly maxResults?: number;
}

export interface TextTarget {
  readonly pageIndex: number;
  readonly text: string;
  readonly rects: readonly PageRect[];
  /** Elements that contain the matched text, in reading order. */
  readonly elementIds: readonly string[];
}

export interface EditState {
  /** Starts at 0 and grows by one with every applied change (apply, undo, redo, reset). */
  readonly revision: number;
  /** True when the content differs from the last `save()` result, or from the original. */
  readonly dirty: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly pageCount: number;
}

export interface ApplyOptions {
  /** Rejects the call with `edit-conflict` unless `state.revision` equals this value. */
  readonly expectedRevision?: number;
  /** Validates and simulates the batch without changing anything. */
  readonly dryRun?: boolean;
  /** Free text stored with the history entry, for host UIs and audit logs. */
  readonly label?: string;
  readonly signal?: AbortSignal;
}

export interface HistoryOptions {
  readonly expectedRevision?: number;
  readonly signal?: AbortSignal;
}

export interface SaveOptions {
  readonly signal?: AbortSignal;
}

export interface EditReceipt {
  /** `state.revision` after the call; unchanged for a dry run or a no-op. */
  readonly revision: number;
  readonly dryRun: boolean;
  readonly operationCount: number;
  /** Ids of the elements the batch created, in operation order. */
  readonly createdIds: readonly string[];
  /** Page indexes in the resulting document whose content changed. */
  readonly changedPages: readonly number[];
  readonly pageCount: number;
  readonly warnings: readonly ViewerWarning[];
}

/** One problem found while validating a batch. */
export interface OperationIssue {
  readonly operationIndex: number;
  /** JSON Pointer into the operation, for example "/style/color". */
  readonly path: string;
  /** Stable machine code, for example "required", "type", "range", "unknown-operation", "unknown-target". */
  readonly code: string;
  readonly message: string;
}

export interface EditSessionBase<
  TOperation extends EditOperation,
  TElement extends EditElement,
> {
  readonly format: EditableFormat;
  /** Frozen snapshot, replaced (never mutated) on every state change. */
  readonly state: EditState;
  /** JSON Schemas of every operation this session accepts. */
  readonly schemas: OperationSchemaSet;

  apply(
    operations: readonly TOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  undo(options?: HistoryOptions): Promise<EditReceipt>;
  redo(options?: HistoryOptions): Promise<EditReceipt>;
  /** Drops every change and returns to the original bytes. Cannot be undone. */
  reset(options?: HistoryOptions): Promise<EditReceipt>;
  /** Bytes of the current state; marks this state as saved. */
  save(options?: SaveOptions): Promise<Uint8Array>;

  getElements(query?: ElementQuery): Promise<readonly TElement[]>;
  getElement(id: string): Promise<TElement | undefined>;
  elementsAt(pageIndex: number, point: PagePoint): Promise<readonly TElement[]>;
  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<readonly TextTarget[]>;
}

/**
 * The session of the loaded document. Format modules narrow this to a union of
 * format sessions discriminated by `format`.
 */
export type EditSession = EditSessionBase<EditOperation, EditElement>;

export interface EditStateChange extends EditState {
  readonly active: boolean;
  readonly format?: EditableFormat;
}

export type DocumentChangeReason = "apply" | "undo" | "redo" | "reset";

export interface DocumentChange {
  readonly revision: number;
  readonly reason: DocumentChangeReason;
  readonly changedPages: readonly number[];
  readonly pageCount: number;
}
