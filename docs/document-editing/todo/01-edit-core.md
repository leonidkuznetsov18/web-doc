# Module 01. `edit-core` — editing contract and viewer integration

**Status:** Approved 2026-10-01; implementation not started

## Goal

Define and implement the format-independent part of editing: the public
session API, JSON operations with schemas, validation, history and revisions,
saving, events, element inspection and geometry, and the viewer refresh that
makes every read API show the edited file. After this module a format module
only has to supply an engine; it never touches the viewer.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md), fixed decisions 1–4 and the success
  criteria.
- Research of 2026-10-01, the patterns this module reproduces:
  - GenOffice — one flat operation vocabulary shared by UI, CLI and AI, with
    patch semantics; the whole batch is validated before anything is applied;
    `dryRun`; writes based on stale indexes are refused until the client re-reads.
  - SuperDoc Document API — stable addresses, `expectedRevision`, structured
    receipts.
  - Docmentis — a single JSON command entry point, so every edit is one undo step
    and is easy to automate and test.

## Dependencies

None. The module builds on the existing viewer (`DocumentViewer`, adapters,
viewport, worker RPC).

## In scope

- Public types and methods listed under [Public API](#public-api).
- The internal engine interface that format modules implement.
- History, revisions, dirty state, undo, redo, reset, save.
- The validation pipeline and the export of operation schemas.
- Viewer refresh after a change: reopen the edited bytes, invalidate caches,
  keep the view state, emit events.
- View-geometry helpers for host UIs.
- New error codes, limits and the `editing` capability flag.
- A test-only engine and adapter used to test the core without a real format.
- Documentation.

## Out of scope

- Any format-specific operation (modules 02 and later).
- Editing UI of any kind.
- Persistence of the history across reloads, autosave, storage.
- Collaboration and CRDTs (operations stay serializable to keep that possible).
- Two sessions editing the same file in two viewers.
- AI-specific helpers such as outlines and target resolution (module 07).

## Public API

All new types are exported from the package entry points next to the existing
contracts. Names below are binding; the plan may add internal helpers only.

### Entry points on the viewer

```ts
export type EditableFormat = "pdf" | "pptx" | "docx";

export interface ViewerApi {
  // …existing members…

  /** Starts editing the loaded document, or returns the session already started. */
  edit(options?: EditOptions): Promise<EditSession>;

  /** The active session of the loaded document, if `edit()` was called. */
  getEditSession(): EditSession | undefined;

  /** Client-space rectangle of a page-space rectangle; undefined if the page is not mounted. */
  pageToClient(pageIndex: number, rect: PageRect): ViewportRect | undefined;

  /** Page under a client-space point and the point in page space; undefined outside pages. */
  clientToPage(clientX: number, clientY: number): PageHit | undefined;
}

export interface EditOptions {
  readonly signal?: AbortSignal;
}

export interface DocumentCapabilities {
  // …existing flags…
  /** True when `edit()` is available for this document. */
  readonly editing: boolean;
}
```

### Session

`EditSession` is a union discriminated by `format`. Each format module adds
its session type (`PdfEditSession` in module 02); every member extends
`EditSessionBase`.

```ts
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
  save(options?: { readonly signal?: AbortSignal }): Promise<Uint8Array>;

  getElements(query?: ElementQuery): Promise<readonly TElement[]>;
  getElement(id: string): Promise<TElement | undefined>;
  elementsAt(pageIndex: number, point: PagePoint): Promise<readonly TElement[]>;
  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<readonly TextTarget[]>;
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
```

### Operations, schemas and binary data

```ts
/** A plain JSON object; format modules define the concrete union. */
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
```

Operation values are limited to JSON types (string, finite number, boolean,
`null`, arrays, plain objects) plus `Uint8Array` where a schema declares
`BinaryData`. Functions, class instances, `undefined` inside arrays, `NaN` and
`Infinity` are rejected.

### Elements, text targets and geometry

```ts
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

export interface PagePoint {
  readonly x: number;
  readonly y: number;
}

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
```

**Page space** uses the units of `DocumentInfo.pageSizes` at zoom 1: points
for PDF, CSS pixels for DOCX and PPTX. The origin is the top-left corner of the
page as displayed, `y` grows downwards, and page rotation is already applied.
Font sizes are always in points. Colours are `#RRGGBB` strings.

### Events

```ts
export interface ViewerEventMap {
  // …existing events…
  readonly editstatechange: EditStateChange;
  readonly documentchange: DocumentChange;
}

export interface EditStateChange extends EditState {
  readonly active: boolean;
  readonly format?: EditableFormat;
}

export interface DocumentChange {
  readonly revision: number;
  readonly reason: "apply" | "undo" | "redo" | "reset";
  readonly changedPages: readonly number[];
  readonly pageCount: number;
}
```

### Errors

New `ViewerErrorCode` values:

| Code                | When                                                                             | `details`                                                 |
| ------------------- | -------------------------------------------------------------------------------- | --------------------------------------------------------- |
| `edit-unsupported`  | `edit()` on a document or format that cannot be edited                           | `{ format, reason }`                                      |
| `invalid-operation` | A batch fails shape or engine validation                                         | `{ issues: OperationIssue[] }`                            |
| `edit-conflict`     | `expectedRevision` differs from `state.revision`                                 | `{ expectedRevision, revision }`                          |
| `edit-failed`       | The engine, saving, or reopening the edited bytes failed; the state was restored | `{ stage: "apply" \| "materialize" \| "reopen", cause? }` |

```ts
export interface OperationIssue {
  readonly operationIndex: number;
  /** JSON Pointer into the operation, for example "/style/color". */
  readonly path: string;
  /** Stable machine code, for example "required", "type", "range", "unknown-operation", "unknown-target". */
  readonly code: string;
  readonly message: string;
}
```

Existing codes keep their meaning: `aborted` for cancelled calls,
`lifecycle-error` for calls on an ended session, `resource-limit` for exceeded
limits. Problems found while validating a batch — including text that no font
can draw (issue code `font-unavailable`) — are always reported as
`invalid-operation` issues, never as separate errors, so a client handles one
error shape per batch.

### Limits

Two new `ResourceLimits` fields, configurable like the existing ones:

| Field               | Default | Meaning                                                              |
| ------------------- | ------- | -------------------------------------------------------------------- |
| `maxEditOperations` | 500     | Operations in one `apply()` call                                     |
| `maxEditHistory`    | 200     | Undoable batches kept; older ones are folded into the starting point |

Binary payloads count against `maxInputBytes`; engine work in one call counts
against `maxOperationMs`.

## Behaviour

### Session lifecycle

- `edit()` requires a document in the `ready` state — like every other
  document method it rejects with `lifecycle-error` when none is — whose
  `capabilities.editing` is true; otherwise it rejects with `edit-unsupported`.
  In the MVP that means `pdf`, `pptx` and `docx` loaded from their own formats,
  not converted from DOC, XLS or PPT.
- The first call loads the format engine lazily and gives it a copy of the
  original bytes. Concurrent calls share one pending start. Nothing related to
  editing is fetched or initialised before the first `edit()`.
- There is at most one session per loaded document; later `edit()` calls return
  it. `load()`, `close()` and `destroy()` end the session: pending calls reject
  with `aborted`, unsaved changes are discarded, `editstatechange` reports
  `active: false`, and every later call on the old session rejects with
  `lifecycle-error`.

### Applying a batch

Calls on a session (`apply`, `undo`, `redo`, `reset`, `save`) run one at a
time in call order. `apply()` goes through these steps:

1. **Preconditions.** The session is alive; `expectedRevision`, when given,
   equals `state.revision` (else `edit-conflict`); the batch is not empty and
   holds at most `maxEditOperations` operations.
2. **Shape check.** Each operation is JSON-only and names a known operation;
   all issues are collected before failing with `invalid-operation`.
3. **Engine validation** against the current document: targets exist, values
   are in range, fonts can draw the text, and so on. Any issue fails the whole
   batch with `invalid-operation`; nothing changes.
4. **Dry run** (`dryRun: true`): the engine applies the batch to a throwaway
   copy and the call returns the receipt that an immediate real apply would
   return, with `dryRun: true`. The state does not change.
5. **Apply.** The engine applies the batch, produces the edited bytes and the
   viewer reopens them through the regular adapter. If any of this fails, the
   engine and the viewer are restored to the previous state and the call
   rejects with `edit-failed`.
6. **Commit.** The batch becomes one history entry with its `label`; the
   revision grows by one; `editstatechange` and then `documentchange` are
   emitted; the call resolves with the receipt.

When `apply()` resolves, every read API (`getDocumentInfo`, `renderPage`,
`renderThumbnail`, `getPageText`, `search`, `selectText`, `copySelection`)
already reflects the new state. Hosts that issue many small changes should batch
them into one `apply()` call, since each call reopens the document once.

Aborting before step 5 rejects with `aborted` and changes nothing; aborting
during step 5 restores the previous state and rejects with `aborted`.

### History, revisions and dirty state

- The history is linear and each entry is one `apply()` batch. A new batch after
  `undo()` drops the redo entries.
- `undo()` and `redo()` move one entry back or forward; `reset()` returns to the
  original bytes and clears the history. Each successful call raises the
  revision by one and emits both events with the matching `reason`.
- `undo()` without anything to undo, and `redo()` without anything to redo,
  resolve with a receipt that changes nothing (`operationCount: 0`, same
  revision).
- When the history exceeds `maxEditHistory`, the oldest entry is folded into the
  starting point: it can no longer be undone, but `reset()` still returns the
  original bytes.
- `dirty` is true while the content differs from the content of the last
  `save()`, or from the original before the first save.
- Determinism: the same original and the same history always produce
  byte-identical output. Undoing back to revision 0 produces the original bytes.

### Saving

- `save()` returns the bytes of the current state and marks that state as
  saved. It changes neither the revision nor the document, and emits
  `editstatechange` only if `dirty` changed.
- Without changes it returns bytes identical to the original file.
- `getOriginalBytes()` and `downloadOriginal()` keep returning the original
  file for the whole session.

### Viewer refresh

- A change reopens the edited bytes through the document's normal adapter and
  then swaps the handle; the old handle is closed afterwards.
- Kept: zoom, fit mode, current page (clamped to the new page count) and scroll
  position (clamped).
- Reset: search results and selection (with `searchchange` and
  `selectionchange` set to `null`), cached text maps, the fuzzy-search index and
  render-cache keys (the revision becomes part of the key). `ViewerState` and
  `DocumentInfo` reflect the new page count and page sizes.
- With a mounted container, the visible pages re-render after the call
  resolves; headless viewers render on demand as before.

### View-geometry helpers

- `pageToClient()` and `clientToPage()` account for zoom, scroll position,
  device pixel ratio and per-page sizes, and agree with the rendered canvas to
  within one CSS pixel.
- They return `undefined` for headless viewers, for pages that are not
  currently mounted, and for spreadsheets.
- A result stays valid until the next `viewchange` or `documentchange` event.

## Engine interface for format modules

Internal and not exported from the package root. The plan may refine the
signatures; the responsibilities are binding.

```ts
interface EditEngineProvider {
  readonly formats: readonly EditableFormat[];
  load(original: Uint8Array, context: EditEngineContext): Promise<EditEngine>;
}

interface EditEngine {
  readonly schemas: OperationSchemaSet;
  validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly OperationIssue[]>;
  apply(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<EngineChange>;
  /** Edited bytes of the current state; the original bytes when nothing changed. */
  materialize(signal: AbortSignal): Promise<Uint8Array>;
  /** Rebuilds the state for a history prefix; used by undo, redo, reset and failure recovery. */
  restore(
    batches: readonly (readonly EditOperation[])[],
    signal: AbortSignal,
  ): Promise<void>;
  getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]>;
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

interface EngineChange {
  readonly createdIds: readonly string[];
  readonly changedPages: readonly number[];
  readonly pageCount: number;
  readonly warnings: readonly ViewerWarning[];
}
```

- A `DocumentAdapter` advertises editing through an optional `edit` provider,
  following the optional-capability pattern of `getTextMap`.
  `capabilities.editing` is `true` when the adapter has a provider for the
  document's format.
- Engines must be deterministic: no timestamps, random ids or
  environment-dependent output in materialized bytes.
- Engines may run in a worker; the core never assumes they share the main
  thread's memory.

## Work by layer

### Feature

- Contract additions: types, error codes, limits, the `editing` capability,
  the new `ViewerApi` members and events.
- `src/edit/`: session implementation, call queue, history, revision and dirty
  tracking, shape validation, schema registry, the engine provider lookup.
- Viewer integration: lazy engine start, refresh through the adapter with
  rollback on failure, cache invalidation, view-state preservation, events,
  ending the session on `load`/`close`/`destroy`.
- Viewport: revision in the render key, `pageToClient` and `clientToPage`.
- Test-only engine and adapter (under `test/`, not shipped) that implement the
  engine interface over a trivial document, used by the unit and browser tests.

### Tests

- **Unit:** lifecycle (start, reuse, end on load/close/destroy, calls on an
  ended session); `expectedRevision` conflicts; shape-check and engine issues
  with operation indexes and JSON pointers; atomicity (failing engine, failing
  reopen, abort at each step); dry run leaves state, history and events
  untouched; call queue ordering; undo/redo/reset semantics including no-op
  receipts and the redo tail; `maxEditHistory` folding; `dirty` across save,
  undo and reset; the event order; determinism and identical bytes after
  undoing to revision 0; laziness (no engine load before `edit()`).
- **Browser** (vanilla example with the test adapter): the visible page
  re-renders after `apply`, `undo` and `redo`; zoom, fit and scroll survive a
  change; search results and selection are cleared; a change in page count
  updates `ViewerState` and the page layout; `pageToClient` and `clientToPage`
  match the canvas at zoom 0.5, 1 and 2, device pixel ratio 1 and 2, and after
  scrolling.
- **Types:** a compile-only test proves that `session.format` narrows the
  session union and that operations are checked per format.

### Docs

- New `docs/api/editing.md`: lifecycle, operations and schemas, receipts,
  errors, events, page space and the geometry helpers, guidance for AI clients
  (batching, `expectedRevision`, `dryRun`).
- `docs/api/reference.md`: the new `ViewerApi` members, events, error codes,
  limits and capability flag.
- `docs/architecture.md`: a short section on the editing layer.

## Definition of done

- Every type, method, event, error code and limit above exists, is exported
  and is documented.
- The unit, browser and type tests listed above pass; the browser suite passes
  on Chromium, Firefox and WebKit.
- Viewers that never call `edit()` behave exactly as before: the existing
  suites pass unchanged, and no editing code or asset is requested before the
  first `edit()`.
- The `code` group of the size report grows by at most 20 KB Brotli, and
  format engines are reached only through dynamic `import()` or workers. (The
  size report has no lazy-asset group, so laziness itself is proven by the
  browser tests, not by the report.)
- `npm run check` passes.

## Decisions

Resolved on 2026-10-01 together with the approval of this spec:

1. `apply()` resolves only after the renderer has reopened the edited file.
   This gives the strongest guarantee — reads always match the state — at the
   cost of one reopen per call; hosts are asked to batch. A lazy reopen on the
   next read was rejected as weaker.
2. `undo()` and `redo()` with nothing to do resolve as no-ops instead of
   throwing, which avoids races in host UIs.
3. `maxEditHistory` defaults to 200 batches, to be revisited after measuring
   memory with real engines.
4. `reset()` cannot be undone; hosts that need it can `save()` first.
