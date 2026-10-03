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

The supported interaction model is **overlay editing with commit on blur or
idle**: the host draws its own input surface over the viewer, lets the user
type there, and commits the result as one batch when the field loses focus or
after an idle delay. One batch is one undo step and one reopen of the file, so
`apply()` is a commit path, not a keystroke path; Word-like continuous typing
with live reflow is out of scope for this package.

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

Every session has a `sessionId` — 22 URL-safe characters, unique across
sessions and reloads — that is stamped on `state`, receipts and
`documentchange`. A caller that outlives a session (an agent working through a
host, for example) passes it back as `expectedSessionId` next to
`expectedRevision`; a mismatch rejects with `edit-conflict` before anything is
read or applied.

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
a message from another worker, a recorded script. `applyJson()` is `apply()`
typed for plain `EditOperation` values, so a format-agnostic caller can use it
on the session union without narrowing by `format` first.

Operation values are limited to JSON types (strings, finite numbers, booleans,
`null`, arrays and plain objects). Binary payloads such as images are declared
as `BinaryData`: a `Uint8Array` in-process, a base64 string over a pure-JSON
transport, or an asset reference.

### Assets

`session.addAsset(bytes)` stores binary data once, under its content id, and
returns a reference of the form `asset:<sha-256 hex>` that any `BinaryData`
field accepts (base64 never contains a colon, so the two forms cannot be
confused). Inline payloads are accepted as before, but the session interns
them: before a batch enters the history its bytes are hashed, stored once and
replaced by references, so undo, redo, dry runs and recovery never copy image
data again, and an AI client can register an image once and refer to it in
several batches. An unknown reference is an `invalid-operation` issue coded
`unknown-asset`. Assets live as long as the session and count against
`maxInputBytes` once, when registered.

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
`selectText()` — already reflects the new content, and the viewport repaints
the pages in `changedPages` while keeping zoom, fit and scroll position;
untouched pages keep their bitmaps. The painted geometry follows at the next
frame: `layoutchange` says when the view-geometry helpers describe the new
revision. Search results and the selection are cleared, with `searchchange`
and `selectionchange` set to `null`.

The reopen has two phases. Opening the edited bytes next to the current
document may fail or be aborted, and then nothing changes; the swap itself is
synchronous and cannot fail, so once it ran the call completes even if its
signal was aborted meanwhile — the screen and `save()` never disagree.
`pageCount` comes from the reopened document; an engine that reports a
different count adds a `fidelity-degraded` warning instead of being trusted.

Each call reopens the document once, so prefer one batch of several operations
over several calls.

`ApplyOptions.changeMode` chooses how the batch is written: `direct` (the
default) replaces content in place; `tracked` writes Word revisions for a
person to accept or reject, with `author` (required) and `timestamp` (the
revision's date; an ISO 8601 date-time whenever given, else
`invalid-value` at `/timestamp`). DOCX is the only format with a tracked form; the others
refuse it with `unsupported-change-mode`, as does a DOCX operation without a
tracked form. The [AI editing](./ai-editing.md) page describes it.

### Receipts

```ts
interface EditReceipt {
  readonly sessionId: string;
  readonly revision: number; // after the call; unchanged for a dry run or a no-op
  readonly dryRun: boolean;
  readonly operationCount: number;
  readonly createdIds: readonly string[]; // elements the batch created, in operation order
  readonly removedIds: readonly string[]; // ids that no longer exist, including those of a deleted page
  readonly remappedIds?: Readonly<Record<string, string>>; // old id → new id, formats that rename only
  readonly changedPages: readonly number[]; // may over-approximate; a superset of the changed pages
  readonly pageCount: number;
  readonly warnings: readonly ViewerWarning[];
}
```

An id is never reused: once an element is deleted or undone away, no later
element gets its id in the same session. Ids of created elements derive from
the history state the batch leads to, so a dry run names the ids the real
apply then uses, and a replay of the same history reproduces them.
`removedIds` lists what a call made disappear: the deleted elements of a batch
(every element of a deleted page included), the created elements of an undone
batch, and the deleted elements of a redone one.

### Same-batch references

A target of the form `"$<n>"` names the first element created by operation
`n` of the same batch, which must come earlier and must be an operation that
creates elements:

```ts
await session.apply([
  { op: "insertTable", pageIndex: 0, at, width, rows },
  { op: "setTableCell", target: "$0", row: 0, column: 1, text: "Q2" },
]);
```

References are resolved while the batch is applied. One that lands on an
element the operation cannot act on rejects the whole batch with
`invalid-operation`, like any other issue, and leaves the document unchanged.
Receipts report the final ids in `createdIds`.

## History and saving

| Method                  | Effect                                                                                                                                                                                                                                                                                  |
| ----------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `undo()`                | Reverts the latest batch. With nothing to undo it resolves with a no-op receipt (`operationCount: 0`, same revision).                                                                                                                                                                   |
| `redo()`                | Re-applies the batch undone last. A new `apply()` after an undo drops the redo tail.                                                                                                                                                                                                    |
| `reset()`               | Returns to the original bytes and clears the history. It cannot be undone; `save()` first if that matters.                                                                                                                                                                              |
| `save()`                | Returns `{ bytes, stateToken, sessionId, revision, warnings }` for the current state; without changes the bytes equal the original file. `warnings` names what the bytes do not guarantee (PDF: `privacy-not-guaranteed`). Pure: neither the revision, the document nor `dirty` change. |
| `markSaved(stateToken)` | Records that the host persisted the state a `save()` token names. `dirty` becomes false when the current content is that state — compared by content, so undoing back to a saved state is clean too. A token from another session is reported and ignored.                              |

Each of `undo`, `redo` and `reset` accepts `expectedRevision` and a `signal`.
The history keeps `maxEditHistory` batches (default 200); older ones are folded
into the starting point — still applied, no longer undoable — while `reset()`
always returns to the original.

The same original and the same sequence of calls always produce byte-identical
output, and undoing back to revision 0 produces the original bytes. Behind the scenes
the session retains the bytes of some committed states — every
`maxEditHistory / 4`-th change, within `maxEditCheckpointBytes` — and rebuilds
a state from the nearest such checkpoint instead of replaying everything from
the original. Checkpoints never change the content of a state; the saved
bytes of a state reached through a checkpoint can differ in layout from a
straight replay, which is why the guarantee is stated per sequence of calls.
`getOriginalBytes()` and `downloadOriginal()` keep returning the original file
for the whole session; saving is the host's job:

```ts
const { bytes, stateToken } = await session.save();
await upload(new Blob([bytes], { type: "application/pdf" }));
session.markSaved(stateToken); // only once the upload succeeded
```

A session whose recovery failed (`edit-failed` with `details.recovered: false`)
still answers `save()` with the bytes of its last committed state, so nothing
the user saw is lost. Listeners are isolated: an exception thrown by an event
listener is reported through `reportError` and never rejects the call that
emitted the event. When a session ends, calls still queued reject with
`aborted`; calls made afterwards reject with `lifecycle-error`.

## Inspecting the document

Host UIs need to know what is on a page before they let the user touch it:

```ts
const { items: elements, revision } = await session.getElements({
  pageIndex: 0,
  kinds: ["text"],
});
const { items: hit } = await session.elementsAt(0, { x: 120, y: 80 });
const { items: targets } = await session.findText("Total", {
  pageRange: [0, 2],
});
const { item } = await session.getElement(elements[0].id);
```

Every read returns an envelope — `{ sessionId, revision, items }` for lists,
`{ sessionId, revision, item }` for `getElement()` — naming the state it
describes: reads run in the same queue as changes, so a read queued behind an
`apply()` describes the document after it. A client that wants to act on what
it read passes both values back as `expectedSessionId` and `expectedRevision`.
Reads accept a `signal` (`findText()` takes it in its options) so a hover
query can be dropped when a long change overtakes it.

Every element carries a stable `id`, a format-specific `kind`, `pageIndex`,
`bounds` in page space, optional `rotation`, `text` and `parentId`, and the
names of the `operations` that accept it as their target. `findText()` searches
the document the engine sees, so it can differ slightly from the viewer's
`search()`, which reads the renderer's text layer.

### Page space and view geometry

Page space uses the units of `DocumentInfo.pageSizes` at zoom 1 — points for
PDF, CSS pixels for Office formats — with the origin at the top-left corner of
the page as displayed, `y` growing downwards and page rotation already applied.
Font sizes are always points. Colours are `EditColor` values: a string
(`#RRGGBB`, `#RRGGBBAA`, or `"auto"` where a format has automatic colours) or,
for Office formats, a theme slot `{ theme, mods? }` that keeps the theme link.
PDF accepts the string form only.

An element that spans pages (a Office paragraph or a repeated header) lists
every piece in `fragments`; `pageIndex` and `bounds` describe the first one,
and queries match any fragment. `story` says where a flow element lives
(`body`, `header`, `footer`, `footnote`, `endnote`, `comment`, `notes`,
`layout`, `master`); `frame` carries the untransformed box with rotation and
flips for formats that keep one, while `bounds` stays the axis-aligned box.

### Text positions and ranges

```ts
interface TextPosition {
  readonly elementId: string;
  readonly offset: number; // UTF-16 code units of EditElement.text
}
interface TextRange {
  readonly start: TextPosition; // half-open; may span elements in reading order
  readonly end: TextPosition;
}
```

Offsets count UTF-16 code units of `EditElement.text`, the visible text of the
element. A tab, a line or page break, an inline image and a field each count as
exactly one placeholder character (U+0009, U+000A, U+FFFC, U+FFFC). Every
format follows this convention, so a `TextRange` from `findText()` (its
`ranges` field) can be handed to any operation that takes one.

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

| Event             | Payload                                                                                                                                                                                                                      |
| ----------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `editstatechange` | `EditState` plus `active` and `format`; emitted on start, on every change, when `dirty` flips through `markSaved()`, and on end.                                                                                             |
| `documentchange`  | `{ sessionId, revision, reason: "apply" \| "undo" \| "redo" \| "reset" \| "restore", changedPages, pageCount }`, after `editstatechange`.                                                                                    |
| `layoutchange`    | `{ sessionId, revision, pages }` once the viewport has painted pages of that revision; `pageToClient()` and `clientToPage()` describe the new content from then on. A headless viewer emits it right after `documentchange`. |

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

| Limit                    | Default | Meaning                                                                |
| ------------------------ | ------- | ---------------------------------------------------------------------- |
| `maxEditOperations`      | 500     | Operations in one `apply()` call                                       |
| `maxEditHistory`         | 200     | Undoable batches kept; older ones are folded into the starting point   |
| `maxEditCheckpointBytes` | 64 MiB  | Memory for retained history checkpoints; fewer are kept for a big file |
| `maxOutlineNodes`        | 5 000   | Nodes one `getOutline()` returns; more is cut and reported             |
| `maxDescribeChars`       | 200 000 | Upper bound of `describe()`'s character budget                         |
| `maxEditCheckpoints`     | 20      | Named checkpoints alive at once (see [AI editing](./ai-editing.md))    |

Binary payloads count against `maxInputBytes`; the engine work of one call
counts against `maxOperationMs`.

## PDF

PDF documents are edited with [PDFium](https://pdfium.googlesource.com/pdfium/)
compiled to WebAssembly, running in a dedicated module worker, while PDF.js
keeps rendering. The worker (`workers/pdf-edit-worker.js`), the WebAssembly
module (`assets/pdfium/pdfium.wasm`) and the fallback font
(`fonts/noto-sans-latin-cyrillic.ttf`) are resolved against `assetBaseUrl` and
fetched on the first `edit()` of a PDF, never before. Hosts that serve them
elsewhere pass `edit: { workerUrl, wasmUrl, fallbackFontUrl }` to the PDF
adapter.

`session.format` is `"pdf"` and the session is a `PdfEditSession`: `apply()`
takes `PdfOperation` values, the inspection methods return `PdfElement`
values, and each operation has a typed method with the same name.

### Methods

Every method takes the operation's fields and the usual `ApplyOptions`
(`expectedRevision`, `dryRun`, `label`, `signal`) and resolves with an
`EditReceipt`.

| Method          | Fields                                                                                                                  | Notes                                                                                                                                                                                                                                                                                       |
| --------------- | ----------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `insertTextBox` | `pageIndex`, `rect`, `text` (1–20 000 chars), `style?: PdfTextBoxStyle`                                                 | Wraps the text inside `rect`; the new element's id is `createdIds[0]`.                                                                                                                                                                                                                      |
| `replaceText`   | `target` (a `text` or `textBox`), `text`, `range?: TextRange` inside the target                                         | A text box is laid out again; a text object keeps its font, size, colour and baseline. With `range` only that part changes: a text object is split around it when only a fallback font can draw the new text, the first part keeping the id (`invalid-range` for a range outside the text). |
| `setTextStyle`  | `target` (a `text` or `textBox`), `style: PdfTextBoxStyle`                                                              | Fields left out keep their value. Existing text objects accept `color` and `fontSize` only.                                                                                                                                                                                                 |
| `insertImage`   | `pageIndex`, `rect`, `data: BinaryData`, `mimeType: "image/png" \| "image/jpeg"`                                        | JPEG bytes are embedded as they are; PNG is decoded and stored losslessly with its alpha channel.                                                                                                                                                                                           |
| `insertShape`   | `pageIndex`, `shape: "rectangle" \| "ellipse"` with `rect`, or `shape: "line"` with `from` and `to`; `stroke?`, `fill?` | A rectangle or ellipse needs a stroke, a fill or both; a line needs a stroke.                                                                                                                                                                                                               |
| `setShapeStyle` | `target` (a `shape`), `stroke?: PdfStroke \| null`, `fill?: PdfFill \| null`                                            | `null` removes; absent keeps. A shape keeps at least one of the two.                                                                                                                                                                                                                        |
| `insertTable`   | `pageIndex`, `at`, `width`, `rows: string[][]`, `columnWidths?: number[]`, `style?: PdfTableStyle`                      | 1–100 rows, 1–20 columns, every row the same length; `columnWidths` are relative weights.                                                                                                                                                                                                   |
| `setTableCell`  | `target` (a `table`), `row`, `column`, `text` (up to 2 000 chars, empty clears)                                         | The table is laid out again from its stored inputs.                                                                                                                                                                                                                                         |
| `moveElement`   | `target`, exactly one of `to: PagePoint` (new top-left of the bounds) or `by: { dx, dy }`                               | Any element.                                                                                                                                                                                                                                                                                |
| `resizeElement` | `target`, `rect`                                                                                                        | Text boxes reflow inside `rect`; images, shapes and text objects stretch; tables cannot be resized.                                                                                                                                                                                         |
| `deleteElement` | `target`                                                                                                                | Removes the element and, for text boxes and tables, every object in it.                                                                                                                                                                                                                     |
| `insertPage`    | `index` (0 to the page count), `size?: { width, height }` (3–14 400 pt)                                                 | A blank page; the size defaults to the page before, else after, the position.                                                                                                                                                                                                               |
| `deletePage`    | `pageIndex`                                                                                                             | The last page cannot be deleted (issue code `last-page`). Annotations on the page go with it.                                                                                                                                                                                               |
| `movePage`      | `from`, `to` (the page's index after the move)                                                                          |                                                                                                                                                                                                                                                                                             |
| `rotatePage`    | `pageIndex`, `rotation: 0 \| 90 \| 180 \| 270` or `by: 90 \| 180 \| 270`                                                | `rotation` sets the clockwise angle; `by` turns from the page's current angle, the one the file was saved with included. Exactly one of the two. Page space turns with the page.                                                                                                            |

```ts
interface PdfTextBoxStyle {
  fontFamily?: string; // "Helvetica" (default), "Times", "Courier" or a registered family
  fontSize?: number; // 1–500 pt, default 12
  bold?: boolean;
  italic?: boolean;
  color?: string; // "#RRGGBB", default "#000000"
  align?: "left" | "center" | "right"; // default "left"
  lineHeight?: number; // multiple of the font size, 0.5–5, default 1.2
}

interface PdfStroke {
  color: string;
  width: number;
} // 0–100 pt; 0 is a hairline
interface PdfFill {
  color: string;
}

interface PdfTableStyle {
  fontFamily?: string; // as above, default "Helvetica"
  fontSize?: number; // default 10
  color?: string; // text colour
  borderColor?: string; // default "#000000"
  borderWidth?: number; // 0–20 pt, default 0.75
  cellPadding?: number; // 0–100 pt, default 4
  headerFill?: string; // fill of the first row; none by default
}
```

### Elements

```ts
type PdfElementKind =
  "text" | "image" | "shape" | "textBox" | "table" | "other";

interface PdfElement extends EditElement {
  kind: PdfElementKind;
  textStyle?: PdfTextStyle; // text, textBox
  shapeStyle?: PdfShapeStyle; // shape, table (the grid's stroke and header fill)
  table?: { rows: string[][] }; // table
}

interface PdfTextStyle {
  fontFamily: string; // the family the file declares, e.g. "Helvetica" or "Arial"
  fontSize: number; // points
  bold: boolean;
  italic: boolean;
  color: string;
}
```

| Kind      | What it is                                                              | Accepts                                                                        |
| --------- | ----------------------------------------------------------------------- | ------------------------------------------------------------------------------ |
| `text`    | One text object as stored in the file — often a word, a line or a run   | `replaceText`, `setTextStyle`, `moveElement`, `resizeElement`, `deleteElement` |
| `image`   | An image object                                                         | `moveElement`, `resizeElement`, `deleteElement`                                |
| `shape`   | A path object                                                           | `setShapeStyle`, `moveElement`, `resizeElement`, `deleteElement`               |
| `textBox` | A box created by `insertTextBox`; its lines are listed as one element   | `replaceText`, `setTextStyle`, `moveElement`, `resizeElement`, `deleteElement` |
| `table`   | A table created by `insertTable`; `text` joins cells by tab and newline | `setTableCell`, `moveElement`, `deleteElement`                                 |
| `other`   | Shadings, form XObjects and anything else                               | `moveElement`, `resizeElement`, `deleteElement`                                |

Ids look like `p0:o3` for objects of the original file and `p0:n2.0.0` for
elements an operation created. A file saved by an earlier session already
carries such ids, so a new element whose id one of them has takes the first
free `~n` after it, as in `p0:n2.0.0~1`. They are stable for the whole session —
across undo, redo and page moves — and the same history always yields the
same ids. Text boxes and tables are parametric: their objects carry a
`WebDoc` marked-content tag holding the inputs they were drawn from, so they
are listed as one element again after `save()` and a later `edit()` of the
saved file, in this or another session. A tag that fails validation leaves its
objects as plain `text` and `shape` elements, and so does a tag whose objects
no longer match its inputs — a box or table that another tool moved, resized
or retyped is never rebuilt from stale inputs.

`elementsAt()` lists the elements under a point top-most first. Bounds of
stroked shapes include the stroke, as PDFium reports them.

### Page space

Page space is the page as the viewer displays it at zoom 1: points, origin at
the top-left corner of the crop box, `y` growing downwards, `/Rotate` already
applied. Inserted content is upright on the displayed page whatever its
rotation, and `rotatePage` turns page space with the page. The engine converts
to and from PDF user space for every page, so the same rectangle means the same
thing to `renderPage()`, `pageToClient()` and the session.

### Text and fonts

`insertTextBox` splits paragraphs on `\n`, wraps words greedily to
`rect.width` (a word wider than the box is broken), applies `align` and
`lineHeight`, and writes one text object per line. Text that does not fit
`rect.height` is still drawn and reported as a `fidelity-degraded` warning with
`details.elementId`.

Fonts are chosen in this order:

1. The standard PDF fonts — Helvetica, Times and Courier with their bold and
   italic faces — when every character is WinAnsi-encodable. Nothing is
   embedded.
2. A TrueType or OpenType font registered through `ViewerClientOptions.fonts`
   whose family matches `fontFamily` (case-insensitively) and whose `cmap`
   covers the text. The font file is embedded whole; subsetting is not done.
3. The bundled Noto Sans Latin/Cyrillic fallback, fetched the first time it is
   needed. Using it, or a registered family without the requested bold or
   italic face, adds a `font-substitution` warning.

An unknown `fontFamily` is an issue coded `unknown-font`; text no available
font covers is `font-unavailable`; right-to-left and complex-script text is
`unsupported-script`. Text is horizontal and left-to-right.

`replaceText` on an existing text object keeps its font when that font can
draw the new text — a standard font for WinAnsi text, or an embedded TrueType
font whose `cmap` maps every character to a glyph with outline data (a subset
font can keep the entry for a glyph it emptied) — and otherwise redraws the
text at the same baseline, size and colour in a covering font with a
`font-substitution` warning. The reported
`fontFamily` of existing text is the family the file declares, not the face
PDFium substitutes for a font that is not embedded.

### Images, shapes and tables

- `insertImage` accepts JPEG (embedded unchanged, so the file keeps the
  original stream) and PNG (decoded in the worker, stored as a lossless bitmap
  with the alpha channel as a soft mask). The data counts against
  `maxInputBytes` and the decoded size against `maxDecodedPixels`
  (`resource-limit`); unreadable data is an `invalid-data` issue.
- Rectangles and lines are single paths; ellipses are four Bézier curves.
- `insertTable` draws the grid as one stroked path, an optional header fill and
  one text object per wrapped cell line. Column widths are `width` split by
  `columnWidths` (equal when omitted); each row is as tall as its tallest cell
  at a line height of 1.2. A table that runs past the bottom of the page is
  drawn anyway with a `fidelity-degraded` warning. `setTableCell` and
  `moveElement` redraw it from its stored rows; `resizeElement` is refused
  (`unsupported-target`).

### What stays unchanged

Annotations — links, highlights, comments, form fields — are separate from
page content and are not edited: a link stays where it was when the text under
it moves, and deleting a page removes its annotations. Existing text is edited
one object at a time; paragraphs are not reflowed. Encrypted PDFs cannot be
opened by the viewer and so cannot be edited.

### Saving and signatures

`save({ mode })` takes `"full"` or `"incremental"`:

- **Full** (the default for a document without signature fields) rewrites the
  file and compacts it, so content that was deleted or replaced is gone from
  the output and the bytes do not depend on which pages were read.
- **Incremental** (the default for a signed document) appends the changes to
  the original bytes: the output starts with the original file and earlier
  revisions — including signed ones — stay intact and recoverable, so a reader
  can still get at deleted content. PDFium writes every object it has parsed
  into that update, and rewriting a changed page makes it parse all of them,
  so after any change an incremental save is about twice the original size.

A full save that the compaction pass cannot read — PDFium's output with a
cross-reference stream, or an object it cannot delimit — still resolves: the
bytes are PDFium's uncompacted full save and `save()` adds a warning with code
`privacy-not-guaranteed` (`details.reason: "pdf-compaction"`) to its result,
because deleted or replaced content may then remain recoverable in the file.
The host decides what to do with such a file; a save that holds every
guarantee has an empty `warnings` list. No PDFium output has needed that
branch so far.

After a change the viewer reopens the compacted full save of an unsigned
document, so the copy it shows stays at about the original size; a signed
document is shown in the incremental form, so the signed revision stays intact
in every state the session may restore from. If compaction fails on this
display-only path the uncompacted full save is shown and the failure is
logged, without a warning: privacy is not at stake in what is only displayed.
Only pages an operation touched have their content rewritten. The same
sequence of calls produces byte-identical output, in this session or another
one, and saving without changes returns the original bytes in either mode.

A digitally signed PDF can be edited. The first change reports a
`fidelity-degraded` warning with `details.signatures` because the signatures
cover the original revision only; they remain valid for that revision (with an
incremental save), and a reader that checks them will show the document as
modified since signing. Editing never signs. The same first-change warning
names, in `details.features`, a DocMDP certification (`"docmdp"`, which any
change invalidates), a tagged structure (`"tagged"`: inserted content joins no
structure tree) and a PDF/A claim (`"pdfa"`: inserted standard fonts are not
embedded).

### `findText()` and `search()`

`findText()` searches the text PDFium extracts from the edited document and
returns page rectangles, the ids of the text objects holding the match and
`ranges` — one `TextRange` per element the match touches, with offsets into
that element's `text` (the whole box for a `textBox`, cells joined by tab and
newline for a `table`) — so its results can be passed straight to
`replaceText` or `deleteElement`. The
viewer's `search()` reads the PDF.js text layer, which can join or split runs
differently; use `findText()` to target edits and `search()` to highlight for
the user.

### PDF issue codes

Beyond the shape codes (`required`, `type`, `pattern`, `minimum`, `maximum`,
`min-items`, `max-items`, `additional-property`, `one-of`), engine validation
reports `unknown-target`, `unsupported-target`, `unsupported-style`,
`unknown-font`, `font-unavailable`, `unsupported-script`, `range` (geometry
outside the page, a bad row or column, ragged rows, too few or too many rows),
`invalid-data`, `required` (a shape without stroke and fill, a line without
`from`/`to`), `last-page` and `unknown-operation`.

### Overlay primitives

A host that lets a user edit text in place shows its own input over the
element, commits the change with one `apply()` on blur or idle, and puts the
selection back. These reads serve that flow; each returns the usual envelope
(`sessionId`, `revision`) and queues behind earlier calls like any read.

| Method                                          | Returns                  | What it gives                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| ----------------------------------------------- | ------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `getTextLayout(elementId)`                      | `ReadItem<TextLayout>`   | The lines of a `text`, `textBox` or `table` element in reading order: each line's `range` into `EditElement.text`, `text`, `bounds`, `baseline` (a page-space point), `glyphs` (`offset`, tight `box`, `advance`), `fontFamily`, `fontSize`, `color`. A line is one PDF text object: a text box's drawn lines, a table's cells.                                                                                                                                                                                                                                                                                                                                                                                                         |
| `getTextFont(elementId)`                        | `ReadItem<TextFont>`     | The font a `text` or `textBox` element is drawn in, ready for `FontFace`: `face.data` is the embedded TrueType program (`format: "truetype"`) with the OS/2, name and post tables a subsetting tool left out added and its tables on four-byte boundaries, as browsers require, or an embedded CFF subset wrapped as OpenType (`"opentype"`) with a Unicode cmap built from its glyph names (Adobe Glyph List names, `uniXXXX`, `uXXXX`) and the advance widths of its charstrings, subroutines followed. `key` is the same for every element in one font. Without a face, `missing` says why: `not-embedded` (show `family` by name, as PDF readers do), `cid-keyed`, `type1`, `no-unicode` or `unreadable` (Type 3 fonts among them). |
| `getPageLayout(pageIndex)`                      | `ReadItem<PageLayout>`   | Every text element's layout of a page plus the page's displayed `width` and `height`, in one read.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `positionAt(pageIndex, point)`                  | `ReadItem<TextPosition>` | The caret position nearest to a page-space point; past a glyph's middle in reading direction the caret goes after it. `undefined` on a page without text.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `rangeRects(range)`                             | `ReadResult<PageRect>`   | The rectangles a range covers, one per line fragment, for drawing a selection; empty for a range across pages.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| `renderPageWithout(pageIndex, ids, { scale? })` | `ReadItem<PageBitmap>`   | The page as PDFium draws it with those elements left out: RGBA pixels over white at `scale` device pixels per point (default 1, bounded by `maxDecodedPixels`). Nothing is reopened and the bytes do not change; unknown ids are ignored.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| `elementsForSelection(selection)`               | `ReadResult<TextRange>`  | The viewer's `TextSelection` as element ranges: a layout line whose box the selected run covers by half, else the one line that contains the run, else a line whose NFKC-folded text contains the run's; the run's text decides the offsets. Merged per element, in reading order.                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| `mapRange(range, fromRevision)`                 | `ReadItem<TextRange>`    | Where a range taken at `fromRevision` is now: ranged `replaceText` calls shift it, a deleted element makes it `undefined`, an undo brings it back, renamed ids are followed, and the element's current text bounds it. `undefined` when the session no longer remembers that revision.                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| `elementsAtSync(pageIndex, point)`              | `ReadResult<PdfElement>` | `elementsAt` from a main-thread cache of the last `getElements({ pageIndex })` result, without waiting behind a queued `apply()`; the revision is the cached one. `cachedPages` lists the pages it holds; changed pages are refreshed after every commit.                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |

`replaceText` takes an optional `range` (see the operations table), so the
committed change touches only what the user typed over.

The flow, end to end:

1. On hover, `elementsAtSync` (after one `getElements({ pageIndex })`) tells
   which element is under the pointer; `getTextLayout` gives its lines and
   glyph boxes to place the input. When the input opens, `getTextFont` gives
   the face to type in, so it shows the file's own glyphs and widths; load
   each `key` once.
2. When the user selects text in the viewer, `elementsForSelection` turns
   `viewer.getSelection()` into a range; `rangeRects` draws it.
3. While the input is open, `renderPageWithout` gives a bitmap of the page
   without the element, which the host lays over the page so the old text does
   not show through; `viewer.pageToClient` places it.
4. On blur or idle, one `replaceText` with the range commits the text. The
   receipt's `revision` and `documentchange` say when the page is current.
5. `mapRange` with the revision the range was taken at says where it is now;
   `viewer.selectText` restores the selection.

A face from `getTextFont` holds only what the file embeds. A subset font
draws only the characters the file uses, so list a fallback after it in the
input's `font-family`; the browser draws other characters in that one. The
face is declared regular, since its glyphs carry their own weight and slant:
load it with the default `FontFace` descriptors and keep the input at normal
weight and style, or the browser draws a synthetic bold or italic over it.
The face changes nothing the engine writes: `replaceText` keeps or
substitutes fonts as the font rules above say.

Glyph geometry comes from PDFium's text page mapped through the page's
rotation and crop box, never from PDF.js; where a line starts and ends agrees
with the PDF.js text layer within one CSS pixel, which a browser test keeps
true on rotated and cropped pages. Run boxes from the viewer are proportional
cuts of a run, so `elementsForSelection` trusts the selected text for the
offsets and uses the boxes to pick between repeated words.

### Performance

Each `apply()` saves the working copy and reopens it in PDF.js. Measured in
headless Chromium on an Apple M4 Pro with one text object per page: one
`insertTextBox` takes about 46 ms on 10 pages, 130 ms then 40 ms on 100 pages
and 72 ms then 56 ms on 500 pages (the first apply on a document pays for
PDF.js parsing it again); the browser suite fails above three seconds. Batch operations that belong
together, and keep `getElements()` queries to the pages you need — the engine
loads a page only when an operation or a query touches it.

## PPTX

PPTX decks are edited XML-first on the shared OOXML package layer: the
original package is the source of truth, every change is a patch to a slide
part, untouched ZIP entries are copied byte for byte, and the parts an edit
touched are stored uncompressed, so a saved deck is byte-identical whatever
engine wrote it. The package lives in a module worker
(`workers/ooxml-edit-worker.js`, resolved against `assetBaseUrl` and fetched
on the first `edit()` of a deck, never before; no WebAssembly). The same
`@silurus/ooxml` renderer that shows the deck reopens the edited bytes after
each change. `.pptm` and `.ppsx` decks take the same path; a macro part is
copied byte for byte and never read. Hosts that serve the worker elsewhere
pass `edit: { workerUrl }` to the Office adapter.

`session.format` is `"pptx"` and the session is a `PptxEditSession`: `apply()`
takes `PptxOperation` values, the inspection methods return `PptxElement`
values, and each operation has a typed method with the same name. Two more
reads describe the deck: `getSlides()` lists the slides in order with a key
that survives reordering (`"sld3"`, the slide part's number) and their
layout; `getLayouts()` lists every layout of every master with its id
(`"layout2"`), name and type, for `insertSlide`.

### Methods

Geometry is slide space: CSS pixels at 96 dpi (`EMU / 9525`), the unit of
`DocumentInfo.pageSizes` for presentations. Font sizes and line widths are
points.

| Method           | Fields                                                                                                                  | Notes                                                                                                                                                                                                                                                                                                                                                                                                                 |
| ---------------- | ----------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `replaceText`    | `target` (a `shape` with text), `text` (up to 100 000 chars), `range?: TextRange` on the target                         | Without `range` the whole body: new text takes the first run's properties, `\n` splits paragraphs (each copies the first paragraph's properties), `\v` is a line break. With `range` only that part: untouched runs keep their bytes, a collapsed range inserts with the style of the run before it, a range across a paragraph break merges under the first paragraph. A range that cuts a field is `invalid-range`. |
| `setTextStyle`   | `target` (a `shape` with text), `range?`, `style: PptxTextStyleChange`                                                  | Writes only the properties given, splitting runs at the range ends; `align` applies to every paragraph the range touches (all of them without a range). The paragraph end takes the change too, so text typed later inherits it.                                                                                                                                                                                      |
| `setShapeStyle`  | `target` (a `shape` or `connector`), `fill?: EditColor \| "none" \| null`, `line?: { color, width? } \| "none" \| null` | `"none"` writes no fill or no line; `null` drops the explicit value so the theme or style reference applies again; absent keeps it. A connector takes `line` only.                                                                                                                                                                                                                                                    |
| `moveElement`    | `target`, exactly one of `to: PagePoint` (new top-left of the bounds) or `by: { dx, dy }`                               | Any element; a group child moves in its group's space, a placeholder that inherited its frame gets an explicit one.                                                                                                                                                                                                                                                                                                   |
| `resizeElement`  | `target`, `rect`                                                                                                        | `rect` is the new `bounds`; a rotated element keeps its rotation and gets the frame whose rotated box is `rect`. Resizing a group scales its children.                                                                                                                                                                                                                                                                |
| `deleteElement`  | `target`                                                                                                                | Removes the element and the slide's relationships only it used; media parts stay in the package (PowerPoint drops orphans on its own save). Every element of a deleted group is in `removedIds`.                                                                                                                                                                                                                      |
| `insertTextBox`  | `pageIndex`, `rect`, `text`, `style?: PptxTextStyleChange`                                                              | A text box (`spAutoFit`, no fill) with the next free id of the slide; the new element's id is `createdIds[0]`.                                                                                                                                                                                                                                                                                                        |
| `insertImage`    | `pageIndex`, `rect`, `data: BinaryData`, `mimeType: "image/png" \| "image/jpeg"`                                        | The bytes are stored once as a media part (identical bytes already in the package are reused) and placed in `rect` as given; the bytes must carry the signature of their type (`invalid-value`).                                                                                                                                                                                                                      |
| `insertTable`    | `pageIndex`, `rect`, `rows: string[][]`, `columnWidths?: number[]`, `style?: { firstRow?, bandRow? }`                   | 1–100 rows, 1–20 columns, every row the same length; `columnWidths` are relative weights; rows share `rect.height` equally; the deck's default table style applies when it names one.                                                                                                                                                                                                                                 |
| `setTableCell`   | `target` (a `table`), `row`, `column`, `text`                                                                           | Replaces the cell's text like `replaceText` on a whole body; the cell's properties stay. A row or column outside the table is a `range` issue.                                                                                                                                                                                                                                                                        |
| `insertSlide`    | `index` (0 to the slide count), `layout?` (a `getLayouts()` id)                                                         | A slide with the layout's placeholders instantiated (titles, bodies, content, pictures, tables, charts, media), as PowerPoint's New Slide does; their ids are in `createdIds`. The default layout is that of the slide before the position, else the first.                                                                                                                                                           |
| `duplicateSlide` | `pageIndex`, `index?` (default: right after the source)                                                                 | Copies the slide and clones the parts only it may own (charts, diagrams, embeddings); layouts, images and media are shared; notes and comments are not copied. The copy's element ids are in `createdIds`.                                                                                                                                                                                                            |
| `deleteSlide`    | `pageIndex`                                                                                                             | Removes the slide, its notes slide and its presentation entry; the last slide cannot be deleted (issue code `last-slide`).                                                                                                                                                                                                                                                                                            |
| `moveSlide`      | `from`, `to` (the slide's index after the move)                                                                         | Reorders the slide list only.                                                                                                                                                                                                                                                                                                                                                                                         |

```ts
interface PptxTextStyleChange {
  fontFamily?: string; // written as a:latin; theme names such as "+mn-lt" pass through
  fontSize?: number; // 1–400 pt
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: EditColor; // "#RRGGBB", or { theme, mods } for a scheme colour
  align?: "left" | "center" | "right" | "justify";
}

type EditColor = string | { theme: string; mods?: Record<string, number> };
```

A theme colour names a scheme slot (`tx1`, `bg1`, `accent1` … `accent6`,
`hlink`, `folHlink`, or `dk1`, `lt1`, `dk2`, `lt2`) with optional DrawingML
modifiers (`lumMod`, `lumOff`, `tint`, `shade`, `alpha`, `satMod`, …) in
thousandths of a percent, as the file stores them. What was written in theme
form is read back in theme form, so the link to the theme survives a round
trip through a host's UI.

### Elements

```ts
type PptxElementKind =
  | "shape" // p:sp — a text box, a placeholder, an auto shape, WordArt
  | "image" // p:pic
  | "table" // p:graphicFrame holding a:tbl
  | "connector" // p:cxnSp
  | "group" // p:grpSp; its children carry parentId
  | "other"; // charts, diagrams, OLE objects, media frames

interface PptxElement extends EditElement {
  kind: PptxElementKind;
  name: string; // p:cNvPr/@name, as PowerPoint's selection pane shows it
  placeholder?: { type: string; idx?: number };
  textStyle?: PptxTextStyle; // a shape with a text body
  shapeStyle?: PptxShapeStyle; // shape, connector
  table?: { rows: string[][] }; // table
  hidden?: boolean; // p:cNvPr/@hidden: listed and editable, not drawn
}

interface PptxTextStyle {
  fontFamily: string; // theme fonts resolved ("+mn-lt" → the minor Latin face)
  fontSize: number; // points
  bold: boolean;
  italic: boolean;
  underline: boolean;
  color: EditColor; // "auto" when the file uses a fill the API cannot express
  align: "left" | "center" | "right" | "justify";
}

interface PptxShapeStyle {
  fill?: EditColor | "none"; // absent when inherited from the style or theme
  line?: { color: EditColor | "none"; width: number }; // points
}
```

| Kind        | What it is                                             | Accepts                                                                                                            |
| ----------- | ------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------ |
| `shape`     | A text box, placeholder, auto shape or WordArt         | `replaceText`, `setTextStyle` (with a text body), `setShapeStyle`, `moveElement`, `resizeElement`, `deleteElement` |
| `image`     | A picture                                              | `moveElement`, `resizeElement`, `deleteElement`                                                                    |
| `table`     | A table; `text` joins cells by tab and rows by newline | `setTableCell`, `moveElement`, `resizeElement`, `deleteElement`                                                    |
| `connector` | A connector line                                       | `setShapeStyle` (line only), `moveElement`, `resizeElement`, `deleteElement`                                       |
| `group`     | A group; its children list it as `parentId`            | `moveElement`, `resizeElement`, `deleteElement`                                                                    |
| `other`     | A chart, diagram, OLE object or media frame            | `moveElement`, `resizeElement`, `deleteElement`                                                                    |

Ids are `<slide key>:<shape id>`, for example `sld3:7`: the number of the
slide part and `p:cNvPr/@id`. They are stable across reordering and across
sessions on the same file, so a host can persist them. A file with duplicate
ids on one slide gets `#2`, `#3` suffixes on the later duplicates. A new
element takes an id above every id its slide holds or held in the session
(an id is never reused after a deletion, and a replay issues the same
ids); a new slide takes the next free slide number, as PowerPoint allocates
them, so a slide number is reissued only after a slide the session created
was deleted and its elements reported in `removedIds`.

`textStyle` describes the first run with text, resolved through the
placeholder chain (run, the shape's list style, the layout placeholder, the
master placeholder, the master text styles, the presentation defaults) and
the theme's fonts. `bounds` is the axis-aligned box of the rotated frame;
`frame` keeps the untransformed box with `rotation`, `flipH` and `flipV`. A
placeholder without a frame of its own reports the one it inherits from its
layout or master; a group child's frame is mapped through its groups.
Elements inside the fallback branch of `mc:AlternateContent` are listed with
no operations. Only the slide's own shapes are listed: what the renderer
composes from the layout or master is not editable here.

`elementsAt()` lists the elements under a point top-most first, groups after
their children. `findText()` searches the text of every shape and table and
returns the shape's bounds as the match rectangle: the engine has no glyph
geometry, so the viewer's `search()` remains the source of word rectangles.

### Text

A shape's `text` is its paragraphs joined by `\n`; inside a paragraph a line
break is `\v` (PowerPoint's own character for it), a tab is `\t` and a field
contributes its cached text. `TextRange` offsets count UTF-16 code units of
that text. Characters XML cannot carry (control characters, lone surrogates)
are `invalid-text`.

A shape whose body autofits (`a:normAutofit`) carries a font scale computed
for its old text; `replaceText` and `setTextStyle` drop that scale, so the
text shows unscaled until PowerPoint lays the shape out again — it may
overflow the box in the viewer until then.

Text is written as given, except that `\r\n` and `\r` become `\n` (the
element's `text` reads back normalized) and a range may not split a
surrogate pair (`invalid-range`). Slide-space coordinates are bounded to
±2.8 billion pixels (DrawingML's own limit) by the schemas.

### What stays unchanged

Everything an operation does not touch keeps its bytes: other shapes, other
slides, layouts, masters, themes, notes, comments, animations, transitions,
custom XML and extension lists. Layouts, masters and notes are not editable.
Table cells accept text changes only. Fonts are written by name; nothing is
embedded, and PowerPoint substitutes a missing face.

Two cases do reach beyond the element: deleting a shape that a slide's
animations target removes the slide's `p:timing` with it (an animation
pointing at a missing shape makes PowerPoint repair the file) and reports
a `fidelity-degraded` warning with `details.reason: "animations-removed"`;
and slide operations keep the deck's sections (`p14:sectionLst`) and custom
shows in step, so a new slide joins the section of the slide before it and
a deleted slide leaves both lists.

### PPTX issue codes

Besides the core codes, validation reports `unknown-target`, `invalid-target`
(a read-only element or one without the needed text body or style),
`invalid-range`, `invalid-text`, `invalid-value` (a colour that is neither
`#RRGGBB` nor a theme slot, image bytes that do not match their type, rows
of unequal length), `unknown-asset`, `unknown-layout`, `range` (a slide or
cell index), `required`/`conflict` (`moveElement` needs exactly one of `to`
and `by`) and `last-slide`.

### Performance

Each `apply()` saves the package and reopens it in the renderer, which lays
the slides out progressively on a reopen so the slide on screen paints
without waiting for the whole deck. Measured in the headless browser matrix
on an Apple M4 Pro with synthetic decks of one placeholder and one text box
per slide, a `replaceText` resolves in about 30 ms on 10 slides, 20–40 ms on
100 slides and 35–40 ms on 500 slides in Chromium; Firefox takes 90–145 ms
and WebKit 55–160 ms on 500 slides. An `insertTextBox` on the last of 500
slides takes 50–180 ms. The browser suite fails above three seconds on every
browser of the matrix. Batch operations that belong together.

## DOCX

DOCX documents are edited XML-first on the same OOXML package layer and in
the same module worker as PPTX (`workers/ooxml-edit-worker.js`, fetched on
the first `edit()` of a document). The worker holds the original package
and a block index of the body story: every paragraph and table of
`w:body`, block-level content controls unwrapped, the paragraphs of table
cells included. `.docm` documents take the same path; the macro part is
copied byte for byte and never read. Headers, footers, footnotes, endnotes
and comments are not listed in this release.

`session.format` is `"docx"` and the session is a `DocxEditSession`:
`apply()` takes `DocxOperation` values, the inspection methods return
`DocxElement` values, and each operation has a typed method with the same
name. The operations the session accepts are the ones `session.schemas`
lists; `operations` on each element names the ones that take it as their
target.

### Methods

Positions in a flow document are other elements, never page points. Font
sizes and spacing are points.

| Method                                                   | Effect                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| -------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `replaceText({ target, text, range? })`                  | Replaces the whole text of a paragraph, or the part a range covers (a collapsed range inserts). Newlines in `text` split the paragraph: the new paragraphs copy its properties (never a section break), get fresh ids reported in `createdIds`, and the text after the range moves to the last of them. `\t`, `\v` and `\f` write a tab, a line break and a page break.                                                                                                                                                                                   |
| `setTextStyle({ target, range?, style })`                | Changes `fontFamily`, `fontSize`, `bold`, `italic`, `underline`, `color` (`#RRGGBB`, `"auto"` or `{ theme }` with a Word theme colour name such as `accent1` or `text1`) and `highlight` (a Word highlight name or `"none"`) on the runs a range covers, splitting runs at its ends; a range that reaches the paragraph end also sets the paragraph mark, so text typed after it inherits the change.                                                                                                                                                     |
| `setParagraphStyle({ target, style })`                   | Changes `align` (`left`, `center`, `right`, `justify`) and `spacing` (`before` and `after` in points, `line` as a multiple of single spacing); other paragraph properties keep their bytes.                                                                                                                                                                                                                                                                                                                                                               |
| `insertParagraph({ before \| after, text, style? })`     | Adds paragraphs next to a paragraph or table of the body or of a cell (one per line of `text`, ids in `createdIds`): a paragraph reference lends its properties without any section break and its first run's style, a table reference gives a plain paragraph; `style` merges into that. Inserting after the last paragraph keeps the body's section properties last.                                                                                                                                                                                    |
| `deleteElement({ target })`                              | Removes a paragraph, a table or an inline picture. The last paragraph of the body or of a cell (`last-paragraph`) and a paragraph that ends a section (`section-break`) are refused; a body that would end with a table gets an empty paragraph (in `createdIds`). A picture's relationship goes with it when nothing else uses it; media parts stay. `removedIds` lists every element removed, a table's cell paragraphs included.                                                                                                                       |
| `moveElement({ target, before \| after })`               | Moves a paragraph or a table next to another element of the same body or cell, bytes intact and id kept; a paragraph that ends a section cannot move.                                                                                                                                                                                                                                                                                                                                                                                                     |
| `insertImage({ before \| after, data, mimeType, size })` | Adds a paragraph holding an inline PNG or JPEG at `size` points (the display pre-pass fits it to the column like any other picture). The bytes are stored once under `word/media` and related from the document; `createdIds` names the paragraph, then the picture.                                                                                                                                                                                                                                                                                      |
| `insertTable({ before \| after, rows, columnWidths? })`  | Adds a table next to a paragraph or table of the body (not inside a cell): a grid over the section's content width from the relative `columnWidths` (equal when omitted), the `TableGrid` style when the document defines it or single borders otherwise, one paragraph per cell with the cell's text (newlines become line breaks), and an empty paragraph after the table when the next block would be a table or the end of the body. `createdIds` names the table, then every cell paragraph, then that trailing paragraph. 1–100 rows, 1–20 columns. |
| `setTableCell({ target, row, column, text })`            | Replaces a cell's text in its first paragraph (properties and first run style kept, newlines as line breaks) and removes the cell's other paragraphs; a row or column outside the table is a `range` issue. Cells are counted as the file lists them, merged cells included.                                                                                                                                                                                                                                                                              |
| `getRevisions(elementId, options?)`                      | A read: the tracked changes a paragraph holds, in document order (`ins`, `del`, `moveFrom`, `moveTo`, `rPrChange`, `pPrChange` with `id`, `author`, `date`, `scope` and the text they cover); empty for other elements.                                                                                                                                                                                                                                                                                                                                   |

A table is named after its first paragraph, so an insertion, a move or a
deletion that changes which paragraph comes first in its first cell
renames the table: the receipt's `remappedIds` maps the old id to the new
one.

Every text method targets a paragraph that is not read-only. The text model
offsets of `range` are those of the element's `text`. A range that starts
or ends inside a field, or that cuts one edge of a hyperlink or an inline
content control, is `invalid-range`; a range that lies inside a hyperlink
edits its text, and one that covers a field, a hyperlink, a content
control or an inline picture whole removes it (the picture's id is
reported in `removedIds`). A paragraph break cannot be inserted inside a
hyperlink or content control. New text is styled like the run that held
the first replaced character, like the run before a collapsed range, else
like the run after it, else like the paragraph mark.

### Elements

| Kind        | What it is                                                | Id                                                  |
| ----------- | --------------------------------------------------------- | --------------------------------------------------- |
| `paragraph` | A `w:p` of the body or of a table cell                    | `p:<id>`                                            |
| `table`     | A `w:tbl` of the body; `table.rows` carries the cell text | `tbl:<id of its first paragraph>`                   |
| `image`     | An inline picture (`wp:inline`) inside a paragraph        | `img:<paragraph id>.<n>`, `n` counting the pictures |
| `other`     | An anchored drawing, an embedded object, an equation      | `other:<paragraph id>.<n>`; listed, never edited    |

A paragraph's `<id>` is its `w14:paraId` when the file has one (Word
writes one on every paragraph), otherwise a deterministic id from its
position among the unmarked paragraphs of the document, eight upper-case
hex digits. It is the `paragraphId` the viewer reports on every DOCX text
run and the name of the hidden bookmark the viewer's display copy carries,
so a selection, a text run and an element name the same paragraph. Ids are
stable for the session and never reused. A cell paragraph carries its
table as `parentId` and its cell as `cell` (`row` and `column`, from 0); it
moves only among the paragraphs of that cell. A picture or other object
carries its paragraph.

`text` of a paragraph is its logical text: `w:t` text, a tab `\t`, a line
break `\v`, a page or column break `\f`, an inline picture or embedded
object as one object character (U+FFFC), a field (`w:fldSimple` or a
complex field from `fldChar begin` to `end`) as its cached result, the runs
of hyperlinks and inline content controls, hidden runs as text, deleted
text left out. A table's `text` joins cells by tabs and rows by newlines.
`textStyle` is the resolved style of the paragraph's first run with text
(run properties, character style, paragraph style, defaults, theme fonts
and colours); `paragraphStyle` carries the style id, alignment, spacing
(points, or a multiple of single spacing for `lineRule: "auto"`) and list
numbering.

A paragraph that cannot be edited in place says why in `readOnlyReason`:
`tracked-changes` (it holds `w:ins`, `w:del` or a move, on its runs or on
its paragraph mark), `section-break` (its `w:pPr` carries a `w:sectPr`) or
`unsupported-content` (all its text comes from fields). Such a paragraph
only accepts the insertion operations that place a sibling next to it.
Changed properties alone (`w:rPrChange`, `w:pPrChange`) leave a paragraph
editable. `getRevisions(elementId)` lists a paragraph's revisions (kind,
id, author, date, the text they cover) in document order; see
[tracked changes](./ai-editing.md#tracked-changes).

### Geometry

The engine never lays the document out: geometry comes from the renderer.
Every DOCX text run the viewer reports carries `paragraphId`, and the
session joins elements with runs on the main thread: a paragraph's
`fragments` are the unions of its runs' boxes per page, a table's are the
unions of its cell paragraphs', a picture or other object takes its
paragraph's. `bounds` and `pageIndex` are those of the first fragment.

A query with `pageIndex` reads that page's runs and returns the elements
with a fragment on it (`intersects` then filters by that fragment). A query
without `pageIndex` lists every body element but reads only the pages the
viewer has already laid out; an element on no laid-out page has
`pageIndex` −1, empty `bounds` and an empty `fragments` list. A paragraph
that draws no run (an empty one, or one holding only pictures) is placed by
estimate: right after the placed paragraph before it in the same body or
cell, else right before the one after it, a picture paragraph with its
largest picture's declared extent (`imageSize` on the picture element) and
an empty paragraph with a line of its neighbour's height. `elementsAt()` reads the page's
runs and returns the paragraph under the point followed by its table.
`findText()` searches the paragraphs' logical text in document order and
places each match from the runs: the first page that holds the paragraph
is found in the viewer's cache, then in `pageRange`, then page by page,
and the paragraph's pages are grown around it; the rectangles are the
matched characters' share of their runs when the run text aligns with the
paragraph text, else the paragraph's runs on its first page. `pageRange`
bounds the matches by the page they are placed on, and `maxResults`
counts the matches on those pages. Page space is CSS pixels at 96 dpi,
the unit of `DocumentInfo.pageSizes` for documents.

### What stays unchanged

A session without changes saves the original bytes. Nothing outside the
body part is touched by any operation; headers, footers, notes, comments,
styles, numbering and settings keep their bytes. Inside the body part an
edit rebuilds only the paragraph it targets: every other paragraph keeps
its bytes, and within the rebuilt paragraph untouched runs, bookmarks,
comment markers, hyperlinks, content controls and fields keep theirs. A
rebuilt or new paragraph is written with a `w14:paraId` (the document
root gains the `w14` namespace declaration when it lacks one), so its id
survives a save and a later session; untouched paragraphs without one
keep none in the saved file. The copy the viewer shows carries an id on
every paragraph and a part of its own (`/webdoc/unauthored.xml`) naming
the ids the engine stamped, so the runs keep naming the engine's
paragraphs after edits that move paragraphs around, undo and redo
restore the exact bytes of the earlier state, and a checkpoint reopened
as the base of a later state still saves without the stamped ids.

Paragraphs of headers, footers, footnotes, endnotes and comments are
not listed and cannot be targeted. Their `paragraphId` on the viewer's
runs is positional when the file carries no `w14:paraId` for them, and
follows the body's: it can change after the first edit of a document,
once the body's paragraphs are stamped.

Each `apply()` reopens the document, which lays every page out again;
`changedPages` of the receipt and of `documentchange` is every page from
the first page of the edited paragraph to the end of the document, as
the viewer knew that paragraph's page before the change.

### DOCX issue codes

Besides the shared codes (`required`, `type`, `unknown-operation`,
`unknown-target`, `unknown-asset`), DOCX validation reports:

| Code             | Where                                                                             | Meaning                                                                                                                                                                                                          |
| ---------------- | --------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `invalid-target` | `/target`, `/before`, `/after`                                                    | The element is not of the kind the operation takes, is read-only, or the reference is not in the same body or cell.                                                                                              |
| `invalid-range`  | `/range`, `/text`                                                                 | Ends off the target, out of order or out of bounds, a split surrogate pair, a cut field, a cut hyperlink or content-control edge, or a paragraph break inside one.                                               |
| `invalid-text`   | `/text`, `/rows/<r>/<c>`                                                          | Control characters XML cannot carry, or a lone surrogate.                                                                                                                                                        |
| `invalid-value`  | `/style/color`, `/style/fontFamily`, `/before`, `/rows`, `/columnWidths`, `/data` | A colour that is not `#RRGGBB`, `auto` or a theme colour name; a bad font name; neither or both of `before` and `after`; ragged rows; one weight per column missing; bytes that are not the declared image type. |
| `last-paragraph` | `/target`                                                                         | The last paragraph of the body or of a cell cannot be deleted.                                                                                                                                                   |
| `section-break`  | `/target`                                                                         | A paragraph that ends a section cannot be deleted or moved.                                                                                                                                                      |
| `range`          | `/row`, `/column`                                                                 | The row or column lies outside the table.                                                                                                                                                                        |

### Performance

Inspection costs one scan of the body part when the session starts and
one read of the page's runs per query; the runs are the viewer's cached
text maps, computed on first use.

Each `apply()` saves the package, runs the display pre-pass and reopens
the document, which lays every page out again. Measured in the headless
browser matrix on an Apple M4 Pro with synthetic documents of one
paragraph per page, a `replaceText` resolves in 70–85 ms on 10 pages,
70–520 ms on 100 pages and 150 ms to 3.1 s on 500 pages in Chromium; an
`insertParagraph` on the last page takes 70 ms, 500 ms and 2.4 s
respectively. The slowest browser of the matrix takes up to 470 ms on 10
pages, 2.3 s on 100 and 9.5 s on 500. The browser suite fails above three
seconds on 10 and 100 pages and records 500. CI runs this measurement
separately with one worker on every matrix browser so parallel functional
tests cannot compete for its CPU time. Commit on idle: debounce
typing and commit on blur, batch the operations that belong together, and
expect a long document to take seconds per commit.

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
- Start a turn with `describe()` or `getOutline()`: the document as a prompt
  sees it, ids first. The [AI editing](./ai-editing.md) page documents the
  grammar and the budget.
