import type {
  DocumentFormat,
  PageSize,
  RegisteredFont,
  ResourceLimits,
  TextRun,
  ViewerWarning,
} from "../contracts.js";
import type { EditSession } from "./sessions.js";
import type {
  ChangeMode,
  EditableFormat,
  EditElement,
  EditFindOptions,
  EditOperation,
  EditSessionBase,
  ElementQuery,
  OperationIssue,
  OperationSchemaSet,
  PagePoint,
  ReadItem,
  ReadOptions,
  ReadResult,
  TextTarget,
  TextAnchorMigration,
} from "./types.js";

/*
 * The interface between the editing core and a format engine. It is not
 * exported from the package entry points: built-in adapters provide engines,
 * and the shape may change between minor releases.
 */

export interface EditEngineContext {
  readonly format: EditableFormat;
  readonly fileName?: string;
  readonly limits: ResourceLimits;
  readonly assetBaseUrl?: URL;
  /** Fonts the host registered with the client, for text the engine writes. */
  readonly fonts?: readonly RegisteredFont[];
  readonly signal: AbortSignal;
}

/**
 * The format-independent session a provider wraps into its typed session.
 * The read hooks let a typed session add engine reads that queue behind
 * earlier calls and carry the same envelope as the core's own reads.
 */
export interface EditSessionCore extends EditSessionBase<
  EditOperation,
  EditElement
> {
  /** The host's limits, for reads a typed session builds over the core. */
  readonly limits: ResourceLimits;
  readItem<T>(
    options: ReadOptions | undefined,
    task: (engine: EditEngine, signal: AbortSignal) => Promise<T | undefined>,
  ): Promise<ReadItem<T>>;
  readItems<T>(
    options: ReadOptions | undefined,
    task: (engine: EditEngine, signal: AbortSignal) => Promise<readonly T[]>,
  ): Promise<ReadResult<T>>;
}

/**
 * What the viewer lends a typed session beyond the core: the text runs of
 * the shown document, for formats whose engine never lays out (DOCX joins
 * its elements with the renderer's runs on the main thread).
 */
export interface EditSessionAccess {
  /** Text runs of a page of the shown document; computed on demand and cached by the viewer. */
  getTextRuns(
    pageIndex: number,
    signal?: AbortSignal,
  ): Promise<readonly TextRun[]>;
  /** Pages whose text runs are already known, in no particular order. */
  cachedPages(): readonly number[];
  /** Renders a temporary document without changing the viewer's shown handle. */
  previewDocument?(
    bytes: Uint8Array,
    options: DocumentPreviewRenderOptions,
    signal: AbortSignal,
    /** DOCX paragraph ids whose preceding runs are needed for unambiguous offsets. */
    paragraphIds?: readonly string[],
  ): Promise<DocumentPreviewRead>;
}

export interface DocumentPreviewRenderOptions {
  /** Per-page decoded raster pixels; a positive safe integer, capped by runtime maxDecodedPixels. */
  readonly maxPixelsPerPage?: number;
  /** Detached targets owned by the caller; publish only after this read resolves. */
  readonly pages: readonly {
    readonly pageIndex: number;
    readonly target: HTMLCanvasElement | OffscreenCanvas;
  }[];
  readonly zoom?: number;
  readonly devicePixelRatio?: number;
}

export interface DocumentPreviewPages {
  readonly pageCount: number;
  readonly pageSizes?: readonly PageSize[];
  readonly pages: readonly {
    readonly pageIndex: number;
    readonly runs: readonly TextRun[];
  }[];
}

/** Internal text alignment input from the same temporary handle, never extra raster pages. */
export interface DocumentPreviewRead extends DocumentPreviewPages {
  readonly alignmentPages?: DocumentPreviewPages["pages"];
}

/** Advertised by a `DocumentAdapter` that can edit some of its formats. */
export interface EditEngineProvider {
  /** Document formats the provider edits; variants map to a session format (pptm → pptx). */
  readonly formats: readonly DocumentFormat[];
  load(original: Uint8Array, context: EditEngineContext): Promise<EditEngine>;
  /** Adds the format's typed methods on top of the core session. */
  createSession(core: EditSessionCore, access: EditSessionAccess): EditSession;
}

/** Bytes of a state, with what the engine could not guarantee about them. */
export interface MaterializedDocument {
  readonly bytes: Uint8Array;
  /** For example `privacy-not-guaranteed` when a PDF full save could not be compacted. */
  readonly warnings: readonly ViewerWarning[];
}

/** How a batch is written; travels with the batch to the engine and into the history, so a replay writes it the same way. */
export interface BatchMode {
  /** Default `direct`; `tracked` writes revisions where the format has them. */
  readonly changeMode?: ChangeMode;
  /** The author of tracked changes. */
  readonly author?: string;
  /** ISO 8601, the date of tracked changes. */
  readonly timestamp?: string;
}

/** A batch with the identity the core assigned to the state after it. */
export interface EngineBatch extends BatchMode {
  /** Unique within the session and never reused; engines derive created ids from it. */
  readonly stateId: number;
  readonly operations: readonly EditOperation[];
}

/** Format-specific save fields, as the session's `save()` received them without the signal. */
export type MaterializeOptions = Readonly<Record<string, unknown>>;

/** What `restore` rebuilds: a base document (a checkpoint, else the original) plus batches. */
export interface RestoreTarget {
  readonly base?: Uint8Array;
  readonly batches: readonly EngineBatch[];
}

export interface EngineChange {
  readonly createdIds: readonly string[];
  /** Ids that no longer exist after the batch, including every element of a deleted page. */
  readonly removedIds: readonly string[];
  /** Old id → new id, when the format had to rename an element. */
  readonly remappedIds?: Readonly<Record<string, string>>;
  readonly textAnchorMigrations?: readonly TextAnchorMigration[];
  /** A superset of the pages whose content changed. */
  readonly changedPages: readonly number[];
  /**
   * Flow formats: the `paragraphId` of the text runs where the document
   * reflows; the host turns it into every page from that paragraph's first
   * page to the end, and `changedPages` is ignored.
   */
  readonly reflowFrom?: string;
  /** Optional: the core takes the count from the renderer and only cross-checks this. */
  readonly pageCount?: number;
  readonly warnings: readonly ViewerWarning[];
}

/**
 * A mutable working copy of one document. The core serializes every call, so
 * an engine never sees two calls at once.
 */
export interface EditEngine {
  readonly schemas: OperationSchemaSet;
  /** Validates against the current state; an empty result means the batch can be applied. */
  validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
    mode?: BatchMode,
  ): Promise<readonly OperationIssue[]>;
  /**
   * Applies an already validated batch to the working copy. Same-batch
   * references (`"$<n>"` targets) are resolved here; one that resolves to an
   * element the operation cannot act on rejects with `invalid-operation`.
   */
  apply(batch: EngineBatch, signal: AbortSignal): Promise<EngineChange>;
  /**
   * Bytes of the current state; the original bytes when nothing changed.
   * `show` asks for what the viewer reopens, `save` for what the host keeps;
   * formats read their own fields from `options` (the PDF save mode).
   */
  materialize(
    purpose: "show" | "save",
    options: MaterializeOptions,
    signal: AbortSignal,
  ): Promise<Uint8Array>;
  /**
   * `materialize` with what the bytes do not guarantee; an engine without it
   * is taken to guarantee everything.
   */
  materializeDocument?(
    purpose: "show" | "save",
    options: MaterializeOptions,
    signal: AbortSignal,
  ): Promise<MaterializedDocument>;
  /**
   * Rebuilds a state: the base document (a checkpoint, else the original)
   * with the batches applied in order. Used by undo, redo, reset, dry runs
   * and recovery.
   */
  restore(target: RestoreTarget, signal: AbortSignal): Promise<void>;
  /** Keeps asset bytes for the session; operations refer to them by id. */
  putAsset(id: string, data: Uint8Array, signal: AbortSignal): Promise<void>;
  getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]>;
  /** Direct lookup; without it the core scans `getElements`. */
  getElement?(
    id: string,
    signal: AbortSignal,
  ): Promise<EditElement | undefined>;
  elementsAt(
    pageIndex: number,
    point: PagePoint,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]>;
  findText(
    query: string,
    options: EditFindOptions,
    signal: AbortSignal,
  ): Promise<readonly TextTarget[]>;
  dispose(): Promise<void>;
}
