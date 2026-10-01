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

| Method          | Fields                                                                                                                  | Notes                                                                                               |
| --------------- | ----------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------- |
| `insertTextBox` | `pageIndex`, `rect`, `text` (1–20 000 chars), `style?: PdfTextBoxStyle`                                                 | Wraps the text inside `rect`; the new element's id is `createdIds[0]`.                              |
| `replaceText`   | `target` (a `text` or `textBox`), `text`                                                                                | A text box is laid out again; a text object keeps its font, size, colour and baseline.              |
| `setTextStyle`  | `target` (a `text` or `textBox`), `style: PdfTextBoxStyle`                                                              | Fields left out keep their value. Existing text objects accept `color` and `fontSize` only.         |
| `insertImage`   | `pageIndex`, `rect`, `data: BinaryData`, `mimeType: "image/png" \| "image/jpeg"`                                        | JPEG bytes are embedded as they are; PNG is decoded and stored losslessly with its alpha channel.   |
| `insertShape`   | `pageIndex`, `shape: "rectangle" \| "ellipse"` with `rect`, or `shape: "line"` with `from` and `to`; `stroke?`, `fill?` | A rectangle or ellipse needs a stroke, a fill or both; a line needs a stroke.                       |
| `setShapeStyle` | `target` (a `shape`), `stroke?: PdfStroke \| null`, `fill?: PdfFill \| null`                                            | `null` removes; absent keeps. A shape keeps at least one of the two.                                |
| `insertTable`   | `pageIndex`, `at`, `width`, `rows: string[][]`, `columnWidths?: number[]`, `style?: PdfTableStyle`                      | 1–100 rows, 1–20 columns, every row the same length; `columnWidths` are relative weights.           |
| `setTableCell`  | `target` (a `table`), `row`, `column`, `text` (up to 2 000 chars, empty clears)                                         | The table is laid out again from its stored inputs.                                                 |
| `moveElement`   | `target`, exactly one of `to: PagePoint` (new top-left of the bounds) or `by: { dx, dy }`                               | Any element.                                                                                        |
| `resizeElement` | `target`, `rect`                                                                                                        | Text boxes reflow inside `rect`; images, shapes and text objects stretch; tables cannot be resized. |
| `deleteElement` | `target`                                                                                                                | Removes the element and, for text boxes and tables, every object in it.                             |
| `insertPage`    | `index` (0 to the page count), `size?: { width, height }` (3–14 400 pt)                                                 | A blank page; the size defaults to the page before, else after, the position.                       |
| `deletePage`    | `pageIndex`                                                                                                             | The last page cannot be deleted (issue code `last-page`). Annotations on the page go with it.       |
| `movePage`      | `from`, `to` (the page's index after the move)                                                                          |                                                                                                     |
| `rotatePage`    | `pageIndex`, `rotation: 0 \| 90 \| 180 \| 270`                                                                          | Absolute clockwise rotation; page space turns with it.                                              |

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
elements an operation created. They are stable for the whole session —
across undo, redo and page moves — and the same history always yields the
same ids. Text boxes and tables are parametric: their objects carry a
`WebDoc` marked-content tag holding the inputs they were drawn from, so they
are listed as one element again after `save()` and a later `edit()` of the
saved file, in this or another session. A tag that fails validation leaves its
objects as plain `text` and `shape` elements.

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
draw the new text — a standard font for WinAnsi text, or an embedded font whose
`cmap` covers it — and otherwise redraws the text at the same baseline, size
and colour in a covering font with a `font-substitution` warning. The reported
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

`save()` returns the original bytes followed by an incremental update, so the
output always starts with the original file and earlier revisions stay intact.
Only pages an operation touched have their content rewritten. The same history
produces byte-identical output, in this session or another one, and undoing to
revision 0 — or saving without changes — returns the original bytes.

A digitally signed PDF can be edited. The first change reports a
`fidelity-degraded` warning with `details.signatures` because the signatures
cover the original revision only; they remain valid for that revision, and a
reader that checks them will show the document as modified since signing.
Editing never signs.

### `findText()` and `search()`

`findText()` searches the text PDFium extracts from the edited document and
returns page rectangles plus the ids of the text objects holding the match, so
its results can be passed straight to `replaceText` or `deleteElement`. The
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

### Performance

Each `apply()` saves the working copy and reopens it in PDF.js. One operation
on a ten-page document resolves well inside a second on a developer machine;
the browser suite fails above three seconds. Batch operations that belong
together, and keep `getElements()` queries to the pages you need — the engine
loads a page only when an operation or a query touches it.

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
