# Editing API

`web-doc` can edit a loaded document through a programmatic session. The
package ships no editing user interface: the host application draws its own
toolbars, selection handles and inspectors on top of the viewer, and its AI
agent drives the same session with JSON operations. Every change is applied to
the original file and drawn by the engine that renders the viewer, so what the
user sees is what `save()` returns.

Editing is available for the formats a built-in engine supports; the
`capabilities.editing` flag of `getDocumentInfo()` says whether the loaded
document qualifies. Format-specific operations are documented in their own
sections of this page as they ship.

## Starting a session

```ts
const viewer = client.createViewer({ container });
await viewer.load(file, { fileName: file.name });

if (viewer.getDocumentInfo().capabilities?.editing) {
  const session = await viewer.edit();
  console.log(session.format, session.state.revision);
}
```

`edit()` starts the format engine the first time it is called and returns the
same session on every later call; `getEditSession()` returns it without
starting one. Nothing related to editing — code, workers or WebAssembly — is
fetched before the first `edit()`.

A session belongs to the loaded document. `load()`, `close()` and `destroy()`
end it: pending calls reject with `aborted`, unsaved changes are discarded, and
later calls on the old session reject with `lifecycle-error`.

`edit()` rejects with:

| Code               | When                                                                                    |
| ------------------ | --------------------------------------------------------------------------------------- |
| `lifecycle-error`  | No document is ready, or the viewer is destroyed.                                       |
| `edit-unsupported` | The document's format has no engine; `details.format` and `details.reason` say which.   |
| `edit-failed`      | The engine could not start (`details.stage` is `"load"`); a later `edit()` tries again. |

## Session state

```ts
interface EditState {
  readonly revision: number; // 0 at start; +1 for every apply, undo, redo and reset
  readonly dirty: boolean; // content differs from the last save (or the original)
  readonly canUndo: boolean;
  readonly canRedo: boolean;
  readonly pageCount: number;
}
```

`session.state` is a frozen snapshot replaced on every change. The viewer
emits `editstatechange` with the same fields plus `active` and `format`, both
when the session starts (`active: true`) and when it ends (`active: false`).

## Operations

Every change is a batch of plain JSON operations:

```ts
const receipt = await session.apply(
  [{ op: "insertTextBox", pageIndex: 0, rect, text: "Hello" }],
  { expectedRevision: session.state.revision, label: "Add title" },
);
```

Format sessions also expose one typed method per operation; each method is
`apply()` with a single operation and the same options. Use the methods from
TypeScript and the JSON form from anything that speaks JSON — an AI tool call,
a message from another worker, a recorded script.

Operation values are limited to JSON types (strings, finite numbers, booleans,
`null`, arrays and plain objects). Binary payloads such as images are declared
as `BinaryData`: a `Uint8Array` in-process, or a base64 string over a pure-JSON
transport.

### Schemas

`session.schemas` holds one JSON Schema (draft 2020-12) per operation, keyed by
its `op` name, with a `version` that is raised whenever an operation's shape
changes incompatibly. Hand them to an LLM as tool definitions, or validate
batches on the host before sending them.

### How a batch is applied

1. **Preconditions.** The session is alive; `expectedRevision`, when given,
   equals `state.revision`; the batch is not empty and holds at most
   `maxEditOperations` operations.
2. **Shape check** against the schemas, then **engine validation** against the
   current document (targets exist, values are in range, fonts can draw the
   text). Every issue of the batch is collected before the call rejects with
   `invalid-operation`.
3. **Dry run.** With `dryRun: true` the batch is simulated and the receipt an
   immediate real apply would produce is returned; nothing changes.
4. **Apply.** The engine applies the batch and the viewer reopens the edited
   bytes. If anything fails, the engine and the viewer are restored and the
   call rejects with `edit-failed`.
5. **Commit.** The batch becomes one undo step, the revision grows by one, and
   `editstatechange` then `documentchange` are emitted.

A batch is applied completely or not at all. When `apply()` resolves, every
read API — `getDocumentInfo()`, `renderPage()`, `getPageText()`, `search()`,
`selectText()` — already reflects the new content, and the viewport has
re-rendered the changed pages while keeping zoom, fit and scroll position.
Search results and the selection are cleared, with `searchchange` and
`selectionchange` set to `null`.

Each call reopens the document once, so prefer one batch of several operations
over several calls.

### Receipts

```ts
interface EditReceipt {
  readonly revision: number; // after the call; unchanged for a dry run or a no-op
  readonly dryRun: boolean;
  readonly operationCount: number;
  readonly createdIds: readonly string[]; // elements the batch created, in operation order
  readonly changedPages: readonly number[]; // page indexes in the resulting document
  readonly pageCount: number;
  readonly warnings: readonly ViewerWarning[];
}
```

## History and saving

| Method    | Effect                                                                                                                                                         |
| --------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `undo()`  | Reverts the latest batch. With nothing to undo it resolves with a no-op receipt (`operationCount: 0`, same revision).                                          |
| `redo()`  | Re-applies the batch undone last. A new `apply()` after an undo drops the redo tail.                                                                           |
| `reset()` | Returns to the original bytes and clears the history. It cannot be undone; `save()` first if that matters.                                                     |
| `save()`  | Returns the bytes of the current state and marks it as saved; without changes the bytes equal the original file. Neither the revision nor the document change. |

Each of `undo`, `redo` and `reset` accepts `expectedRevision` and a `signal`.
The history keeps `maxEditHistory` batches (default 200); older ones are folded
into the starting point — still applied, no longer undoable — while `reset()`
always returns to the original.

The same original and the same history always produce byte-identical output,
and undoing back to revision 0 produces the original bytes.
`getOriginalBytes()` and `downloadOriginal()` keep returning the original file
for the whole session; saving is the host's job:

```ts
const bytes = await session.save();
await upload(new Blob([bytes], { type: "application/pdf" }));
```

## Inspecting the document

Host UIs need to know what is on a page before they let the user touch it:

```ts
const elements = await session.getElements({ pageIndex: 0, kinds: ["text"] });
const hit = await session.elementsAt(0, { x: 120, y: 80 });
const targets = await session.findText("Total", { pageRange: [0, 2] });
```

Every element carries a stable `id`, a format-specific `kind`, `pageIndex`,
`bounds` in page space, optional `rotation`, `text` and `parentId`, and the
names of the `operations` that accept it as their target. `findText()` searches
the document the engine sees, so it can differ slightly from the viewer's
`search()`, which reads the renderer's text layer.

### Page space and view geometry

Page space uses the units of `DocumentInfo.pageSizes` at zoom 1 — points for
PDF, CSS pixels for Office formats — with the origin at the top-left corner of
the page as displayed, `y` growing downwards and page rotation already applied.
Font sizes are always points; colours are `#RRGGBB` strings.

Two viewer methods translate between page space and the host's own DOM:

```ts
const box = viewer.pageToClient(0, element.bounds); // { left, top, width, height } in client pixels
const hit = viewer.clientToPage(event.clientX, event.clientY); // { pageIndex, point } or undefined
```

They account for zoom, scroll position and per-page sizes and agree with the
rendered canvas to within one CSS pixel. They return `undefined` for headless
viewers, for pages that are not currently mounted in the viewport, for points
outside any page, and for spreadsheets. A result stays valid until the next
`viewchange` or `documentchange` event.

## Events

| Event             | Payload                                                                                                            |
| ----------------- | ------------------------------------------------------------------------------------------------------------------ |
| `editstatechange` | `EditState` plus `active` and `format`; emitted on start, on every change, when `dirty` flips on save, and on end. |
| `documentchange`  | `{ revision, reason: "apply" \| "undo" \| "redo" \| "reset", changedPages, pageCount }`, after `editstatechange`.  |

## Errors

| Code                | When                                                                       | `details`                                                           |
| ------------------- | -------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| `edit-unsupported`  | `edit()` on a document without an engine                                   | `{ format, reason }`                                                |
| `invalid-operation` | A batch fails shape or engine validation; nothing was applied              | `{ issues: OperationIssue[] }`                                      |
| `edit-conflict`     | `expectedRevision` differs from `state.revision`                           | `{ expectedRevision, revision }`                                    |
| `edit-failed`       | The engine, saving or reopening failed and the previous state was restored | `{ stage: "load" \| "apply" \| "materialize" \| "reopen", cause? }` |

```ts
interface OperationIssue {
  readonly operationIndex: number; // position in the batch
  readonly path: string; // JSON Pointer into the operation, e.g. "/style/color"
  readonly code: string; // "required", "type", "pattern", "unknown-operation", "unknown-target", "font-unavailable", …
  readonly message: string;
}
```

Problems found while validating — including text that no available font can
draw (`font-unavailable`) — are always reported as `invalid-operation`
issues, so a client handles one error shape per batch. `aborted`,
`lifecycle-error` and `resource-limit` keep their usual meaning.

## Limits

| Limit               | Default | Meaning                                                              |
| ------------------- | ------- | -------------------------------------------------------------------- |
| `maxEditOperations` | 500     | Operations in one `apply()` call                                     |
| `maxEditHistory`    | 200     | Undoable batches kept; older ones are folded into the starting point |

Binary payloads count against `maxInputBytes`; the engine work of one call
counts against `maxOperationMs`.

## Guidance for AI clients

- Read before writing: list elements or `findText()` to obtain ids and
  rectangles, then address operations by those ids.
- Send `expectedRevision` with every batch. An `edit-conflict` means the
  document changed since it was read — read again instead of retrying.
- Use `dryRun: true` to check a batch before applying it; the receipt reports
  the pages it would change and the elements it would create.
- Batch related operations: one `apply()` is one undo step for the user and
  one re-render for the viewer.
- Treat `invalid-operation` issues as structured feedback: `operationIndex`
  and `path` point at the exact field to fix.
