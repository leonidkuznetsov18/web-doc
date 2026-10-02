import type { ViewerWarning } from "../contracts.js";
import type { EditSessionReads } from "./ai/types.js";

/** Formats an edit session can be started for. */
export type EditableFormat = "pdf" | "pptx" | "docx";

export interface EditOptions {
  readonly signal?: AbortSignal;
  /** Reserved: recorded with changes where a format keeps authorship (tracked changes). */
  readonly author?: string;
}

/** A plain JSON object; each format defines its concrete operation union. */
export interface EditOperation {
  readonly op: string;
}

/**
 * Binary payload: bytes, a base64 string for pure-JSON transports, or an
 * `asset:` reference returned by `addAsset()`.
 */
export type BinaryData = Uint8Array | string;

/** A colour as a format accepts it; PDF takes the string form only. */
export type EditColor = ColorString | ThemeColor;

/** `#RRGGBB` or `#RRGGBBAA`, or "auto" where a format has automatic colours. */
export type ColorString = string;

/** A theme slot with optional modifiers, for OOXML; keeps the theme link. */
export interface ThemeColor {
  readonly theme: string;
  readonly mods?: Readonly<Record<string, number>>;
}

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
  /** Opaque, stable for the lifetime of the session, never reused for another element. */
  readonly id: string;
  /** Format-specific kind, for example "text", "image", "shape", "table". */
  readonly kind: string;
  /** Page of the element, or of its first fragment. */
  readonly pageIndex: number;
  /** Axis-aligned bounds in page space, or those of the first fragment. */
  readonly bounds: PageRect;
  /** Clockwise rotation in degrees, when the element is rotated. */
  readonly rotation?: number;
  /** Untransformed frame for formats that keep one (OOXML shapes); `bounds` stays the AABB. */
  readonly frame?: ElementFrame;
  /** Every piece of an element that spans pages; absent when it has one. */
  readonly fragments?: readonly ElementFragment[];
  /** Where the element lives in a flow document; absent for the page body. */
  readonly story?: ElementStory;
  readonly text?: string;
  readonly parentId?: string;
  /** Names of the operations that accept this element as their target. */
  readonly operations: readonly string[];
}

export interface ElementFragment {
  readonly pageIndex: number;
  readonly bounds: PageRect;
}

export interface ElementFrame extends PageRect {
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

export type ElementStory =
  | { readonly kind: "body" }
  | { readonly kind: "header" | "footer"; readonly scope: string }
  | { readonly kind: "footnote" | "endnote" | "comment"; readonly id: string }
  | { readonly kind: "notes" | "layout" | "master"; readonly id: string };

/**
 * A place in an element's text, in UTF-16 code units of `EditElement.text`.
 * A tab, a break, an inline image and a field each count as one placeholder
 * character.
 */
export interface TextPosition {
  readonly elementId: string;
  readonly offset: number;
}

/** Half-open; may span elements in reading order. */
export interface TextRange {
  readonly start: TextPosition;
  readonly end: TextPosition;
}

export interface ElementQuery {
  readonly pageIndex?: number;
  readonly kinds?: readonly string[];
  /** Only elements with a fragment whose bounds intersect this rectangle; requires `pageIndex`. */
  readonly intersects?: PageRect;
}

export interface ReadOptions {
  readonly signal?: AbortSignal;
}

export interface EditFindOptions extends ReadOptions {
  readonly caseSensitive?: boolean;
  /** Inclusive 0-based page range. */
  readonly pageRange?: readonly [number, number];
  readonly maxResults?: number;
}

export interface TextTarget {
  /** Page of the first rectangle. */
  readonly pageIndex: number;
  readonly text: string;
  readonly rects: readonly PageRect[];
  /** Elements that contain the matched text, in reading order. */
  readonly elementIds: readonly string[];
  /** The match as text ranges, one per element it touches, in reading order. */
  readonly ranges: readonly TextRange[];
}

/** Every read says which state it describes. */
export interface ReadEnvelope {
  readonly sessionId: string;
  readonly revision: number;
}

export interface ReadResult<T> extends ReadEnvelope {
  readonly items: readonly T[];
}

export interface ReadItem<T> extends ReadEnvelope {
  readonly item: T | undefined;
}

export interface EditState {
  /** Unique per session; stamped on receipts, read results and events. */
  readonly sessionId: string;
  /** Starts at 0 and grows by one with every applied change (apply, undo, redo, reset). */
  readonly revision: number;
  /** True while the content differs from the state last given to `markSaved()`, or from the original. */
  readonly dirty: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly pageCount: number;
}

export interface ApplyOptions {
  /** Rejects the call with `edit-conflict` unless `state.revision` equals this value. */
  readonly expectedRevision?: number;
  /** Rejects with `edit-conflict` unless it equals `sessionId`; for callers that outlive a session. */
  readonly expectedSessionId?: string;
  /** Validates and simulates the batch without changing anything. */
  readonly dryRun?: boolean;
  /** Free text stored with the history entry, for host UIs and audit logs. */
  readonly label?: string;
  /** Reserved: ISO 8601 time a format records with the change; never generated by web-doc. */
  readonly timestamp?: string;
  readonly signal?: AbortSignal;
}

export interface HistoryOptions {
  readonly expectedRevision?: number;
  readonly expectedSessionId?: string;
  readonly signal?: AbortSignal;
}

export interface SaveOptions {
  readonly signal?: AbortSignal;
}

/** What `save()` returns: the bytes and the token that names their state. */
export interface SavedDocument {
  readonly bytes: Uint8Array;
  /** Pass it to `markSaved()` once the bytes are persisted. */
  readonly stateToken: string;
  readonly sessionId: string;
  readonly revision: number;
  /**
   * What the bytes do not guarantee: `privacy-not-guaranteed` when a PDF full
   * save could not be compacted and deleted content may remain recoverable.
   * Empty for a save that holds every guarantee.
   */
  readonly warnings: readonly ViewerWarning[];
}

export interface AssetOptions {
  readonly mimeType?: string;
  readonly signal?: AbortSignal;
}

export interface EditReceipt {
  readonly sessionId: string;
  /** `state.revision` after the call; unchanged for a dry run or a no-op. */
  readonly revision: number;
  readonly dryRun: boolean;
  readonly operationCount: number;
  /** Ids of the elements the batch created, in operation order. */
  readonly createdIds: readonly string[];
  /** Ids that no longer exist after the call, including every element of a deleted page. */
  readonly removedIds: readonly string[];
  /** Old id → new id, when a format has to rename an element; absent for formats that never do. */
  readonly remappedIds?: Readonly<Record<string, string>>;
  /** Page indexes in the resulting document whose content may have changed; a superset. */
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
> extends EditSessionReads {
  readonly format: EditableFormat;
  /** Unique per session; stamped on state, receipts, read results and events. */
  readonly sessionId: string;
  /** Frozen snapshot, replaced (never mutated) on every state change. */
  readonly state: EditState;
  /** JSON Schemas of every operation this session accepts. */
  readonly schemas: OperationSchemaSet;

  apply(
    operations: readonly TOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** `apply()` for callers that hold operations as plain JSON; identical behaviour. */
  applyJson(
    operations: readonly EditOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  undo(options?: HistoryOptions): Promise<EditReceipt>;
  redo(options?: HistoryOptions): Promise<EditReceipt>;
  /** Drops every change and returns to the original bytes. Cannot be undone. */
  reset(options?: HistoryOptions): Promise<EditReceipt>;
  /** Bytes of the current state. Pure: changes neither the state nor `dirty`. */
  save(options?: SaveOptions): Promise<SavedDocument>;
  /** Tells the session the host has persisted the state named by `stateToken`. */
  markSaved(stateToken: string): void;
  /** Registers binary data once; returns an `asset:` reference usable in operations. */
  addAsset(data: Uint8Array, options?: AssetOptions): Promise<string>;

  getElements(
    query?: ElementQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<TElement>>;
  getElement(id: string, options?: ReadOptions): Promise<ReadItem<TElement>>;
  elementsAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadResult<TElement>>;
  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<ReadResult<TextTarget>>;
}

export interface EditStateChange extends EditState {
  readonly active: boolean;
  readonly format?: EditableFormat;
}

export type DocumentChangeReason = "apply" | "undo" | "redo" | "reset";

export interface DocumentChange {
  readonly sessionId: string;
  readonly revision: number;
  readonly reason: DocumentChangeReason;
  /** May over-approximate; for flow formats every page from the first affected one. */
  readonly changedPages: readonly number[];
  readonly pageCount: number;
}

/** The viewport has laid out and painted the pages of `revision`; geometry helpers are current. */
export interface LayoutChange {
  readonly sessionId: string;
  readonly revision: number;
  /** Pages painted since the previous `layoutchange`. */
  readonly pages: readonly number[];
}
