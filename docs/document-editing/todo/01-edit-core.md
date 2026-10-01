# Module 01. `edit-core` — editing contract and viewer integration

**Status:** ✅ Revision 1 done 2026-10-01 · 🔄 Revision 2 (ACTION-821) drafted
2026-10-01, awaiting approval; no code changes until it is approved.

Revision 2 follows the architecture review of 2026-10-01 (report kept outside
the repository). It changes the public contract while nothing has been
published to npm, so the breaking parts cost nothing now and would cost a major
version later. Every change is marked **R2** below; unmarked text is unchanged
from revision 1.

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
    receipts, read envelopes stamped with the revision they were evaluated at.
  - Docmentis — a single JSON command entry point, so every edit is one undo step
    and is easy to automate and test.
- **R2** Architecture review of 2026-10-01 (Linear ACTION-821): interactive
  editing needs sub-element addressing, revision-stamped reads, ids that are
  never reused, a renderer-owned page model and a history that does not replay
  everything from the original.

## Interaction model — R2

- web-doc ships no editing UI (roadmap decision 1). The supported interaction
  model for hosts is **overlay editing with commit on blur or idle**: the host
  draws its own input surface over the viewer, lets the user type there, and
  commits the result as one `apply()` batch when the field loses focus or after
  an idle delay. One batch is one undo step and one reopen.
- **Word-like continuous typing with live reflow is out of scope** for this
  package. `apply()` reopens the file, so it is a commit path, not a keystroke
  path; see the roadmap's hard limits.
- What this module gives an overlay: stable element ids, `TextRange`
  addressing (below), revision-stamped reads, the `layoutchange` event that
  says when geometry is current, and the view-geometry helpers. Format modules
  add what the overlay draws with; for PDF the text-layout, hit-testing and
  render-without-element primitives are a separate ticket after ACTION-821.

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
- **R2** Checkpoints and the asset store that keep undo, redo, dry run and
  recovery from replaying the whole history with its binary payloads.
- Documentation.

## Out of scope

- Any format-specific operation (modules 02 and later).
- Editing UI of any kind, including caret, selection and text-input overlays.
- Persistence of the history across reloads, autosave, storage.
- Collaboration and CRDTs (operations stay serializable to keep that possible).
- Two sessions editing the same file in two viewers.
- AI-specific helpers such as outlines and target resolution (module 07).
- **R2** Transient previews, undo coalescing and per-keystroke application.

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
  /** R2 — Reserved: recorded with changes where a format keeps authorship (tracked changes). */
  readonly author?: string;
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
`EditSessionBase`. **R2** The union is defined in `src/edit/sessions.ts` and
is never imported by the engine layer; a format-agnostic caller uses
`applyJson()`, which every member accepts with the same signature.

```ts
export interface EditSessionBase<
  TOperation extends EditOperation,
  TElement extends EditElement,
> {
  readonly format: EditableFormat;
  /** R2 — Unique per session; stamped on state, receipts, read results and events. */
  readonly sessionId: string;
  /** Frozen snapshot, replaced (never mutated) on every state change. */
  readonly state: EditState;
  /** JSON Schemas of every operation this session accepts. */
  readonly schemas: OperationSchemaSet;

  apply(
    operations: readonly TOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** R2 — `apply()` for callers that hold operations as plain JSON; identical behaviour. */
  applyJson(
    operations: readonly EditOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  undo(options?: HistoryOptions): Promise<EditReceipt>;
  redo(options?: HistoryOptions): Promise<EditReceipt>;
  /** Drops every change and returns to the original bytes. Cannot be undone. */
  reset(options?: HistoryOptions): Promise<EditReceipt>;
  /** R2 — Bytes of the current state. Pure: changes neither the state nor `dirty`. */
  save(options?: SaveOptions): Promise<SavedDocument>;
  /** R2 — Tells the session the host has persisted the state named by `stateToken`. */
  markSaved(stateToken: string): void;
  /** R2 — Registers binary data once; returns an `asset:` reference usable in operations. */
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

export interface EditState {
  /** R2 */
  readonly sessionId: string;
  /** Starts at 0 and grows by one with every applied change (apply, undo, redo, reset). */
  readonly revision: number;
  /** True when the content differs from the state last passed to `markSaved()`, or from the original. */
  readonly dirty: boolean;
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  /** R2 — As the renderer reports it after the last reopen. */
  readonly pageCount: number;
}

export interface ApplyOptions {
  /** Rejects the call with `edit-conflict` unless `state.revision` equals this value. */
  readonly expectedRevision?: number;
  /** R2 — Rejects with `edit-conflict` unless it equals `sessionId`; for callers that outlive a session. */
  readonly expectedSessionId?: string;
  /** Validates and simulates the batch without changing anything. */
  readonly dryRun?: boolean;
  /** Free text stored with the history entry, for host UIs and audit logs. */
  readonly label?: string;
  /** R2 — Reserved: ISO 8601 time a format records with the change; never generated by web-doc. */
  readonly timestamp?: string;
  readonly signal?: AbortSignal;
}

export interface HistoryOptions {
  readonly expectedRevision?: number;
  readonly expectedSessionId?: string;
  readonly signal?: AbortSignal;
}

/** R2 */
export interface SaveOptions {
  readonly signal?: AbortSignal;
  // Format sessions extend this (for example the PDF save mode).
}

/** R2 */
export interface SavedDocument {
  readonly bytes: Uint8Array;
  /** Names the content state; pass it to `markSaved()` once the bytes are persisted. */
  readonly stateToken: string;
  readonly sessionId: string;
  readonly revision: number;
}

/** R2 */
export interface AssetOptions {
  readonly mimeType?: string;
  readonly signal?: AbortSignal;
}

/** R2 */
export interface ReadOptions {
  readonly signal?: AbortSignal;
}

/** R2 — Every read says which state it describes. */
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

export interface EditReceipt {
  readonly sessionId: string;
  /** `state.revision` after the call; unchanged for a dry run or a no-op. */
  readonly revision: number;
  readonly dryRun: boolean;
  readonly operationCount: number;
  /** Ids of the elements the batch created, in operation order. */
  readonly createdIds: readonly string[];
  /** R2 — Ids that no longer exist after the call, including every element of a deleted page. */
  readonly removedIds: readonly string[];
  /** R2 — Old id → new id, when a format has to rename an element; absent for formats that never do. */
  readonly remappedIds?: Readonly<Record<string, string>>;
  /** Page indexes in the resulting document whose content may have changed; a superset. */
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

/**
 * Binary payload: bytes, a base64 string for pure-JSON transports, or an
 * `asset:` reference returned by `addAsset()` (R2).
 */
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

**R2 — Assets.** A string payload of the form `asset:<sha-256 hex>` refers to
data registered with `addAsset()` (base64 never contains a colon, so the two
forms cannot be confused). Inline payloads are accepted as before, but the
session interns them: before a batch enters the history its binary payloads are
hashed, stored once, and replaced by asset references, so undo, redo, dry run
and recovery never copy them again. An unknown reference is an
`invalid-operation` issue coded `unknown-asset`. Assets live as long as the
session; they count against `maxInputBytes` once, when registered.

**R2 — Same-batch references.** Inside one batch, the target `"$<n>"` names
the first element created by the operation at index `n` of the same batch,
which must come earlier and must be an operation that creates elements. Such
targets are resolved while the batch is applied; a reference that resolves to
an element the operation cannot act on rejects the whole batch with
`invalid-operation` and leaves the document unchanged, like any other issue.
Receipts report the final ids in `createdIds`.

### Elements, text targets and geometry

```ts
export interface EditElement {
  /** Opaque, stable for the lifetime of the session, never reused for another element (R2). */
  readonly id: string;
  /** Format-specific kind, for example "text", "image", "shape", "table". */
  readonly kind: string;
  /** Page of the element, or of its first fragment. */
  readonly pageIndex: number;
  /** Axis-aligned bounds in page space, or those of the first fragment. */
  readonly bounds: PageRect;
  /** Clockwise rotation in degrees, when the element is rotated. */
  readonly rotation?: number;
  /** R2 — Untransformed frame for formats that keep one (OOXML shapes); `bounds` stays the AABB. */
  readonly frame?: ElementFrame;
  /** R2 — Every piece of an element that spans pages; absent when it has one. */
  readonly fragments?: readonly ElementFragment[];
  /** R2 — Where the element lives in a flow document; absent for the page body. */
  readonly story?: ElementStory;
  readonly text?: string;
  readonly parentId?: string;
  /** Names of the operations that accept this element as their target. */
  readonly operations: readonly string[];
}

/** R2 */
export interface ElementFragment {
  readonly pageIndex: number;
  readonly bounds: PageRect;
}

/** R2 */
export interface ElementFrame extends PageRect {
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

/** R2 */
export type ElementStory =
  | { readonly kind: "body" }
  | { readonly kind: "header" | "footer"; readonly scope: string }
  | { readonly kind: "footnote" | "endnote" | "comment"; readonly id: string }
  | { readonly kind: "notes" | "layout" | "master"; readonly id: string };

/** R2 — A place in an element's text, in UTF-16 code units of `EditElement.text`. */
export interface TextPosition {
  readonly elementId: string;
  readonly offset: number;
}

/** R2 — Half-open; may span elements in reading order. */
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

export interface EditFindOptions {
  readonly caseSensitive?: boolean;
  /** Inclusive 0-based page range. */
  readonly pageRange?: readonly [number, number];
  readonly maxResults?: number;
  readonly signal?: AbortSignal;
}

export interface TextTarget {
  /** Page of the first rectangle. */
  readonly pageIndex: number;
  readonly text: string;
  readonly rects: readonly PageRect[];
  /** Elements that contain the matched text, in reading order. */
  readonly elementIds: readonly string[];
  /** R2 — The match as text ranges, one per element it touches, in reading order. */
  readonly ranges: readonly TextRange[];
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

/** R2 — A colour as a format accepts it. PDF takes the string form only. */
export type EditColor = string | ThemeColor;

/** R2 — `#RRGGBB` or `#RRGGBBAA`, or "auto" where a format has automatic colours. */
export type ColorString = string;

/** R2 — A theme slot with optional modifiers, for OOXML; keeps the theme link. */
export interface ThemeColor {
  readonly theme: string;
  readonly mods?: Readonly<Record<string, number>>;
}
```

**Page space** uses the units of `DocumentInfo.pageSizes` at zoom 1: points
for PDF, CSS pixels for DOCX and PPTX. The origin is the top-left corner of the
page as displayed, `y` grows downwards, and page rotation is already applied.
Font sizes are always in points. Colours are `EditColor` values; the string
form is `#RRGGBB` or `#RRGGBBAA`.

**R2 — Offset convention.** Text offsets count UTF-16 code units of
`EditElement.text`, the visible text of the element. A tab, a line or page
break, an inline image and a field each count as exactly one placeholder
character, which formats render into `text` as U+0009, U+000A, U+FFFC and
U+FFFC respectively. Ranges are half-open. Every format, PDF included, follows
this convention, so a position obtained from `findText()` can be handed to any
operation that takes a range.

### Events

```ts
export interface ViewerEventMap {
  // …existing events…
  readonly editstatechange: EditStateChange;
  readonly documentchange: DocumentChange;
  /** R2 */
  readonly layoutchange: LayoutChange;
}

export interface EditStateChange extends EditState {
  readonly active: boolean;
  readonly format?: EditableFormat;
}

export interface DocumentChange {
  readonly sessionId: string;
  readonly revision: number;
  readonly reason: "apply" | "undo" | "redo" | "reset";
  /** May over-approximate; for flow formats every page from the first affected one. */
  readonly changedPages: readonly number[];
  readonly pageCount: number;
}

/** R2 — The viewport has laid out and painted the pages of `revision`; geometry helpers are current. */
export interface LayoutChange {
  readonly sessionId: string;
  readonly revision: number;
  /** Pages painted since the previous `layoutchange`. */
  readonly pages: readonly number[];
}
```

**R2** Listeners are isolated: an exception thrown by a listener is reported
through `reportError` (or `console.error` where it does not exist) and never
rejects the call that emitted the event.

### Errors

New `ViewerErrorCode` values:

| Code                | When                                                                             | `details`                                                             |
| ------------------- | -------------------------------------------------------------------------------- | --------------------------------------------------------------------- |
| `edit-unsupported`  | `edit()` on a document or format that cannot be edited                           | `{ format, reason }`                                                  |
| `invalid-operation` | A batch fails shape or engine validation                                         | `{ issues: OperationIssue[] }`                                        |
| `edit-conflict`     | `expectedRevision` or `expectedSessionId` differs from the session's             | `{ expectedRevision?, revision, expectedSessionId?, sessionId }` (R2) |
| `edit-failed`       | The engine, saving, or reopening the edited bytes failed; the state was restored | `{ stage: "load" \| "apply" \| "materialize" \| "reopen", cause? }`   |

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

Existing codes keep their meaning: `aborted` for cancelled calls — **R2**
including every call still queued when the session ends — `lifecycle-error`
for calls made after a session ended, `resource-limit` for exceeded limits.
Problems found while validating a batch — including text that no font can draw
(issue code `font-unavailable`) — are always reported as `invalid-operation`
issues, never as separate errors, so a client handles one error shape per
batch.

### Limits

`ResourceLimits` fields, configurable like the existing ones:

| Field                    | Default | Meaning                                                                             |
| ------------------------ | ------- | ----------------------------------------------------------------------------------- |
| `maxEditOperations`      | 500     | Operations in one `apply()` call                                                    |
| `maxEditHistory`         | 200     | Undoable batches kept; older ones are folded into the starting point                |
| `maxEditCheckpointBytes` | 64 MiB  | R2 — Memory for retained checkpoints; fewer checkpoints are kept when a file is big |

Binary payloads count against `maxInputBytes` once, when registered or
interned; engine work in one call counts against `maxOperationMs`.

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
  it. `load()`, `close()` and `destroy()` end the session: pending and queued
  calls reject with `aborted`, unsaved changes are discarded, `editstatechange`
  reports `active: false`, and every later call on the old session rejects with
  `lifecycle-error`.
- **R2** Each session has a `sessionId` that is unique across sessions and
  reloads (a random 128-bit value in URL-safe base64). It is not derived from
  the document.

### Applying a batch

Calls on a session (`apply`, `undo`, `redo`, `reset`, `save`) run one at a
time in call order. `apply()` goes through these steps:

1. **Snapshot** (**R2**). Synchronously, before the call is queued, the batch
   is deep-copied and frozen, so later changes to the caller's objects cannot
   reach the engine.
2. **Preconditions.** The session is alive; `expectedRevision` and
   `expectedSessionId`, when given, match the session (else `edit-conflict`);
   the batch is not empty and holds at most `maxEditOperations` operations.
3. **Shape check.** Each operation is JSON-only and names a known operation;
   all issues are collected before failing with `invalid-operation`. Inline
   binary payloads are interned into the asset store; asset references must
   exist.
4. **Engine validation** against the current document: targets exist, values
   are in range, fonts can draw the text, and so on. Any issue fails the whole
   batch with `invalid-operation`; nothing changes. Same-batch references are
   checked for form here and resolved in step 6.
5. **Dry run** (`dryRun: true`): the engine applies the batch, the edited bytes
   are produced (so an oversize result is caught), the engine is restored from
   the nearest checkpoint, and the call returns the receipt that an immediate
   real apply would return, with `dryRun: true`. The state does not change;
   nothing is reopened in the viewer.
6. **Apply.** The engine applies the batch and produces the edited bytes.
7. **Prepare and commit the reopen** (**R2**). The viewer opens the edited
   bytes next to the current document (`prepare`); any failure here, including
   an abort, restores the engine to the current state and rejects with
   `edit-failed` or `aborted`. Then the handle swap, the state update and the
   cache invalidation happen synchronously and cannot fail (`commit`); from this
   point the call ignores its signal and always completes. The old handle is
   closed afterwards; an error there is reported, not thrown.
8. **Commit.** The batch becomes one history entry with its `label`, the
   `pageCount` the renderer reported, and a `stateId`; the revision grows by
   one; the materialized bytes are kept as the last committed bytes;
   `editstatechange` and then `documentchange` are emitted; the call resolves
   with the receipt.

When `apply()` resolves, every read API (`getDocumentInfo`, `renderPage`,
`renderThumbnail`, `getPageText`, `search`, `selectText`, `copySelection`)
already reflects the new state. The painted geometry follows at the next
frame; `layoutchange` says when `pageToClient()` and `clientToPage()` describe
the new revision. Hosts that issue many small changes should batch them into
one `apply()` call, since each call reopens the document once.

Aborting before step 7's commit point rejects with `aborted` and changes
nothing; after it, the call completes as if it had not been aborted.

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
- **R2 — Checkpoints.** The session retains the materialized bytes of some
  committed states — the fold boundary and every `maxEditHistory / 4`-th
  commit, as far as `maxEditCheckpointBytes` allows — and the engine's `restore`
  starts from the nearest checkpoint at or before the target state instead of
  the original. Undo, redo, dry run and recovery therefore replay at most a
  quarter of the undoable history, and payloads are referenced from the asset
  store, never cloned. Checkpoints are an optimization: they never change the
  bytes a state produces.
- **R2 — Ids.** Every history entry carries a `stateId`, unique within the
  session and never reused, including after undo. Engines derive the ids of
  elements a batch creates from that `stateId`, so an id can never name two
  different elements during a session, and a replay of the same history
  reproduces the same ids. `removedIds` in receipts list what a batch, an undo
  or a redo made disappear.
- `dirty` is true while the content differs from the state last given to
  `markSaved()`, or from the original before the first `markSaved()`.
- Determinism: the same original and the same history always produce
  byte-identical output. Undoing back to revision 0 produces the original bytes.

### Saving

- **R2** `save()` is pure: it returns `{ bytes, stateToken, sessionId, revision }`
  and changes neither the state, the revision nor `dirty`. Without changes the
  bytes are identical to the original file. Format sessions may take a save
  mode in `SaveOptions`.
- **R2** `markSaved(stateToken)` records that the host has persisted that
  state. `dirty` becomes false when the current state is the one named by the
  token, and it is compared by content state, so undoing back to a saved state
  also makes `dirty` false. A token from another session is ignored with a
  reported error.
- **R2** A session whose recovery failed (see below) can still `save()`: the
  bytes of the last committed state are returned, with its token.
- `getOriginalBytes()` and `downloadOriginal()` keep returning the original
  file for the whole session.

### Viewer refresh

- A change reopens the edited bytes through the document's normal adapter and
  then swaps the handle; the old handle is closed afterwards (two phases, see
  step 7 above).
- Kept: zoom, fit mode, current page (clamped to the new page count) and scroll
  position (clamped).
- Reset: search results and selection (with `searchchange` and
  `selectionchange` set to `null`), cached text maps, the fuzzy-search index and
  render-cache keys. **R2** The render key of a page changes only when the page
  is in `changedPages`, so untouched pages keep their painted bitmaps.
  `ViewerState` and `DocumentInfo` reflect the new page count and page sizes.
- With a mounted container, the visible pages re-render after the call
  resolves and `layoutchange` follows the paint; headless viewers render on
  demand as before and emit `layoutchange` after `documentchange`.
- **R2** `pageCount` comes from the reopened document, never from the engine;
  the engine's count, when it reports one, is only cross-checked and a
  mismatch is a `fidelity-degraded` warning.

### Failure recovery

- A failure before the commit point restores the engine to the current history
  state (from the nearest checkpoint) and leaves the viewer untouched. The
  viewer and the engine therefore never disagree.
- A rollback that itself fails leaves the session unusable for changes:
  `apply`, `undo`, `redo` and `reset` reject with `edit-failed`
  (`details.recovered: false`), reads reject likewise, `save()` still works from
  the last committed bytes, and the next `edit()` starts a fresh session.
- A crashed engine worker ends the session the same way.

### View-geometry helpers

- `pageToClient()` and `clientToPage()` account for zoom, scroll position,
  device pixel ratio and per-page sizes, and agree with the rendered canvas to
  within one CSS pixel.
- They return `undefined` for headless viewers, for pages that are not
  currently mounted, and for spreadsheets.
- A result stays valid until the next `viewchange` or `layoutchange` event.

### Reads

- Reads (`getElements`, `getElement`, `elementsAt`, `findText`) run in the same
  queue as changes, so their result describes exactly the revision stamped on
  it. **R2** They accept an `AbortSignal` so a host can drop a hover query that
  a long apply overtook.
- **R2** Every read returns a `ReadEnvelope` with `sessionId` and `revision`.
  A client that wants to act on what it read passes both back as
  `expectedSessionId` and `expectedRevision`.

## Engine interface for format modules

Internal and not exported from the package root. The plan may refine the
signatures; the responsibilities are binding.

```ts
interface EditEngineProvider {
  readonly formats: readonly EditableFormat[];
  load(original: Uint8Array, context: EditEngineContext): Promise<EditEngine>;
  /** Wraps the core in the format's typed session. */
  createSession(core: EditSessionCore): EditSession;
}

/** R2 — A batch with the identity the core assigned to the state after it. */
interface EngineBatch {
  readonly stateId: number;
  readonly operations: readonly EditOperation[];
}

interface EditEngine {
  readonly schemas: OperationSchemaSet;
  validate(
    operations: readonly EditOperation[],
    signal: AbortSignal,
  ): Promise<readonly OperationIssue[]>;
  apply(batch: EngineBatch, signal: AbortSignal): Promise<EngineChange>;
  /** Edited bytes of the current state; the original bytes when nothing changed. */
  materialize(
    purpose: "show" | "save",
    options: MaterializeOptions,
    signal: AbortSignal,
  ): Promise<Uint8Array>;
  /** R2 — Rebuilds the state: `base` (a checkpoint, else the original) plus `batches`. */
  restore(
    target: {
      readonly base?: Uint8Array;
      readonly batches: readonly EngineBatch[];
    },
    signal: AbortSignal,
  ): Promise<void>;
  /** R2 — Keeps asset bytes for the session; operations refer to them by id. */
  putAsset(id: string, data: Uint8Array, signal: AbortSignal): Promise<void>;
  getElements(
    query: ElementQuery,
    signal: AbortSignal,
  ): Promise<readonly EditElement[]>;
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

interface EngineChange {
  readonly createdIds: readonly string[];
  readonly removedIds: readonly string[];
  readonly remappedIds?: Readonly<Record<string, string>>;
  /** A superset of the pages whose content changed. */
  readonly changedPages: readonly number[];
  /** Optional: the core takes the count from the renderer. */
  readonly pageCount?: number;
  readonly warnings: readonly ViewerWarning[];
}
```

- A `DocumentAdapter` advertises editing through an optional `edit` provider,
  following the optional-capability pattern of `getTextMap`.
  `capabilities.editing` is `true` when the adapter has a provider for the
  document's format.
- Engines must be deterministic: no timestamps, random ids or
  environment-dependent output in materialized bytes. **R2** Ids of created
  elements derive from the batch's `stateId`; a `restore` from a checkpoint
  must yield the same ids and bytes as a restore from the original.
- Engines may run in a worker; the core never assumes they share the main
  thread's memory. **R2** The session union is not visible to the engine layer.
- **R2** The host side of the viewer offers the session a two-phase reopen:
  `prepareDocument(bytes, signal)` may fail, `commitDocument(prepared)` is
  synchronous and returns the renderer's page count, `discardDocument(prepared)`
  releases an unused preparation.

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
- **R2** Session: `sessionId`, `applyJson`, read envelopes and signals, pure
  `save` and `markSaved`, last committed bytes, synchronous snapshot of the
  batch, listener isolation, `aborted` on end, `removedIds`, same-batch
  references, the asset store, checkpoints, `stateId`-carrying engine batches,
  the two-phase reopen host interface, renderer-owned page count,
  `layoutchange`, per-page render keys.

### Tests

- **Unit:** lifecycle (start, reuse, end on load/close/destroy, calls on an
  ended session); `expectedRevision` conflicts; shape-check and engine issues
  with operation indexes and JSON pointers; atomicity (failing engine, failing
  reopen, abort at each step); dry run leaves state, history and events
  untouched; call queue ordering; undo/redo/reset semantics including no-op
  receipts and the redo tail; `maxEditHistory` folding; `dirty` across save,
  undo and reset; the event order; determinism and identical bytes after
  undoing to revision 0; laziness (no engine load before `edit()`).
- **R2 Unit:** an abort or a throwing listener after the commit point leaves
  the viewer and the engine on the same state; a dry run whose restore fails
  does not break the session; a throwing listener never rejects a committed
  call; `markSaved` with a stale token keeps `dirty`; a broken session saves
  the last committed bytes; a batch mutated after `apply()` was called is
  applied as it was; queued calls reject with `aborted` on `end()`; ids after
  undo differ from the undone ones; checkpoints give the same bytes as a replay
  from the original; interned payloads are sent to the engine once; `$n`
  references resolve and fail as specified; read envelopes carry the revision
  of the queue position they ran at.
- **Browser** (vanilla example with the test adapter): the visible page
  re-renders after `apply`, `undo` and `redo`; zoom, fit and scroll survive a
  change; search results and selection are cleared; a change in page count
  updates `ViewerState` and the page layout; `pageToClient` and `clientToPage`
  match the canvas at zoom 0.5, 1 and 2, device pixel ratio 1 and 2, and after
  scrolling. **R2** `layoutchange` arrives after `documentchange` and the
  helpers are exact once it has; only pages in `changedPages` repaint.
- **Types:** a compile-only test proves that `session.format` narrows the
  session union, that operations are checked per format, and (**R2**) that
  `applyJson` is callable on the un-narrowed union.
- **R2 Performance:** an `apply()` latency measurement on 10-, 100- and
  500-page fixtures, recorded in the format spec; PDF now, the Office formats
  when their engines exist.

### Docs

- `docs/api/editing.md`: lifecycle, operations and schemas, receipts, errors,
  events, page space and the geometry helpers, guidance for AI clients
  (batching, `expectedRevision`, `dryRun`). **R2** The interaction model,
  envelopes, `save`/`markSaved`, assets, same-batch references, text ranges,
  `layoutchange`, the id rules.
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
- **R2** The PDF module passes its suite on the revised contract; the latency
  measurement is recorded.

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

**R2**, proposed 2026-10-01 after the architecture review:

5. The supported interaction model is overlay editing with commit on blur or
   idle; continuous typing with live reflow is out of scope for the package.
6. Reads return envelopes `{ sessionId, revision, items }`; paging
   (`total`, `limit`, `offset`) is left out and can be added without breaking.
7. `save()` is pure and `markSaved(token)` records persistence; the token names
   a content state, so a saved state reached again by undo is clean.
8. Ids derive from the history entry's `stateId` and are never reused.
9. History restores from retained checkpoints and references payloads through
   the asset store; the folding limit bounds undo depth, the checkpoints bound
   cost.
10. The page count is owned by the renderer; `changedPages` is a superset.
11. The viewer swap is two-phase; after the commit point an abort no longer
    rolls anything back.
12. Text addressing uses UTF-16 offsets over `EditElement.text` with one
    placeholder character per non-text item, for every format.
13. Overlay text-input primitives for PDF (text layout, hit testing, render
    without an element, selection mapping) are a separate ticket after
    ACTION-821; the types they need ship with this revision.

## Actual result

Revision 1, 2026-10-01:

- Public contracts, four error codes, the two limits, the `editing`
  capability, the optional `edit` and `reopen` adapter members and the two
  events shipped as specified; `src/edit/` holds the schema validator, the
  shape checks, the history and the session controller.
- `EditSessionController` serializes calls, validates in two stages, applies
  batches atomically with rollback, supports dry runs, undo, redo, reset, save,
  element and text queries, and bounds each call by `maxOperationMs`. A rollback
  that itself fails makes the session unusable instead of leaving it
  inconsistent.
- The viewer starts engines lazily through the adapter, reopens edited bytes
  with `reopen` or `open`, swaps handles, drops text maps, the fuzzy index
  key, search results and the selection, clamps the page index, keeps zoom,
  fit and scroll, and ends the session on `load`, `close` and `destroy`. The
  content revision is part of the viewport render key; the built-in UI follows
  `documentchange`.
- `pageToClient` and `clientToPage` read the painted slot geometry and
  match the canvas within one CSS pixel at zoom 0.5, 1 and 2, device pixel
  ratio 1 and 2, and after scrolling.
- Tests: 48 new unit tests (contracts, schema, history, session, viewer
  integration) over a fake engine and adapter; `tests/e2e/edit-core.spec.ts`
  passes on Chromium, Chromium at DPR 2, Firefox and WebKit and joined the
  matrix. Compiled eager code grew by 11.3 KB Brotli against `main`.
- The six pre-existing failures of the full Chromium suite on macOS (Linux-only
  SSIM goldens and a headless fullscreen assertion) reproduce without these
  changes and are unrelated.
- Added beyond the spec: `edit-failed` with `details.stage: "load"` when an
  engine fails to start, and an optional `getElement` on the engine
  interface so engines can answer id lookups without a scan.

Revision 2: pending. The review verified against revision 1 that folded
history entries are replayed forever, that an abort or a throwing listener
after the reopen left the viewer ahead of the engine, that ids repeated after
an undo, that a throwing listener rejected a committed call, that `end()`
rejected queued calls with `lifecycle-error`, and that a broken session could
not save; the revised behaviour above addresses each.
