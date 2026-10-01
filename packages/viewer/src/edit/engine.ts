import type { ResourceLimits, ViewerWarning } from "../contracts.js";
import type { EditSession } from "./sessions.js";
import type {
  EditableFormat,
  EditElement,
  EditFindOptions,
  EditOperation,
  EditSessionBase,
  ElementQuery,
  OperationIssue,
  OperationSchemaSet,
  PagePoint,
  TextTarget,
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
  readonly signal: AbortSignal;
}

/** The format-independent session a provider wraps into its typed session. */
export type EditSessionCore = EditSessionBase<EditOperation, EditElement>;

/** Advertised by a `DocumentAdapter` that can edit some of its formats. */
export interface EditEngineProvider {
  readonly formats: readonly EditableFormat[];
  load(original: Uint8Array, context: EditEngineContext): Promise<EditEngine>;
  /** Adds the format's typed methods on top of the core session. */
  createSession(core: EditSessionCore): EditSession;
}

export interface EngineChange {
  readonly createdIds: readonly string[];
  /** Page indexes in the resulting document whose content changed. */
  readonly changedPages: readonly number[];
  readonly pageCount: number;
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
  ): Promise<readonly OperationIssue[]>;
  /** Applies an already validated batch to the working copy. */
  apply(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange>;
  /** Bytes of the current state; the original bytes when nothing changed. */
  materialize(signal: AbortSignal): Promise<Uint8Array>;
  /**
   * Rebuilds the state for a history prefix: the original document with these
   * batches applied in order. Used by undo, redo, reset, dry runs and recovery.
   */
  restore(
    batches: readonly (readonly EditOperation[])[],
    signal: AbortSignal,
  ): Promise<void>;
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
