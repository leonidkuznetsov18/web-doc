# Module 02. `pdf-edit` — PDF editing on PDFium

**Status:** ✅ Revision 1 done 2026-10-01 (T9–T22) · ✅ Revision 2 (ACTION-821)
done 2026-10-01. **R2** marks the changes.

## Goal

Give integrators and AI agents the minimal PDF method set — text,
formatting, alignment, colours, tables, insertion, moving, pages — through the
`edit-core` session. Changes are written by PDFium WASM in a worker; PDF.js
keeps rendering the viewer.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md): fixed decisions 3–5 and the MVP
  method table.
- [`./01-edit-core.md`](./01-edit-core.md): the session, operation, element,
  geometry, error and engine contracts this module implements.
- Research of 2026-10-01:
  - PDF.js 6.2.108 never rewrites page content. It can add annotations, but it
    cannot edit, move or delete existing text and images, and it saves a text
    box that is not WinAnsi-encodable without an appearance stream, so other
    readers show it blank. This is why PDF changes go through PDFium.
  - `@embedpdf/pdfium` 2.15.1 exports everything the MVP needs, checked in its
    `dist/vendor/functions.d.ts`: text objects and fonts
    (`FPDFPageObj_CreateTextObj`, `FPDFText_SetText`, `FPDFText_LoadFont`,
    `FPDFText_LoadStandardFont`), paths (`FPDFPageObj_CreateNewPath`,
    `FPDFPageObj_CreateNewRect`, `FPDFPath_*`), images
    (`FPDFPageObj_NewImageObj`, `FPDFImageObj_LoadJpegFileInline`,
    `FPDFImageObj_SetBitmap`), transforms and colours
    (`FPDFPageObj_Transform`, `FPDFPageObj_SetFillColor`,
    `FPDFPageObj_SetStrokeColor`), insertion and removal
    (`FPDFPage_InsertObject`, `FPDFPage_RemoveObject`,
    `FPDFPage_GenerateContent`), pages (`FPDFPage_New`, `FPDFPage_Delete`,
    `FPDF_MovePages`, `FPDFPage_SetRotation`), text geometry (`FPDFText_*`),
    marked content with string parameters (`FPDFPageObj_AddMark`,
    `FPDFPageObjMark_SetStringParam`) and saving (`FPDF_SaveAsCopy`,
    `PDFiumExt_OpenFileWriter`). Its runtime exports `addFunction`, so file
    callbacks can be provided.
  - GenOffice's PDF editor (Apache-2.0): edits are data applied to the
    untouched original; text is replaced in place only when reading it back
    proves the result; there is a font fallback ladder; PDFium runs as a
    serialized singleton because the WASM heap is shared.

## Dependencies

- Module 01 `edit-core`.
- `@embedpdf/pdfium` 2.15.1, approved on 2026-10-01, pinned exactly.

## In scope

### Inspection

`getElements`, `getElement`, `elementsAt` and `findText` from `edit-core`,
returning `PdfElement` values:

```ts
export type PdfElementKind =
  | "text" // one text object as stored in the file (often a word or a line)
  | "image"
  | "shape" // a path object
  | "textBox" // a text box created by insertTextBox
  | "table" // a table created by insertTable
  | "other"; // shadings, form XObjects and anything else

export interface PdfElement extends EditElement {
  readonly kind: PdfElementKind;
  readonly textStyle?: PdfTextStyle; // text, textBox
  readonly shapeStyle?: PdfShapeStyle; // shape, table
  readonly table?: { readonly rows: readonly (readonly string[])[] }; // table
}

export interface PdfTextStyle {
  readonly fontFamily: string;
  readonly fontSize: number; // points
  readonly bold: boolean;
  readonly italic: boolean;
  readonly color: string; // #RRGGBB
  readonly align?: "left" | "center" | "right"; // textBox only
  readonly lineHeight?: number; // textBox only, multiple of fontSize
}

export interface PdfShapeStyle {
  readonly stroke?: { readonly color: string; readonly width: number };
  readonly fill?: { readonly color: string };
}
```

`findText` searches the text PDFium extracts from the edited document and
returns the matching rectangles and the ids of the text objects that hold them.
It can differ slightly from the viewer's `search()`, which reads PDF.js text;
the docs say so.

### Operations

`PdfEditSession` exposes one typed method per operation. Each method takes the
operation's fields without `op`, plus `ApplyOptions`, and calls `apply()` with
a single operation. Geometry is in page space; font sizes are in points.

| Operation       | Fields                                                                                                               | Category                       |
| --------------- | -------------------------------------------------------------------------------------------------------------------- | ------------------------------ |
| `insertTextBox` | `pageIndex`, `rect`, `text`, `style?: PdfTextBoxStyle`                                                               | text, formatting, alignment    |
| `replaceText`   | `target` (text or textBox), `text`                                                                                   | text                           |
| `setTextStyle`  | `target` (text or textBox), `style: PdfTextBoxStyle`                                                                 | formatting, alignment, colours |
| `insertImage`   | `pageIndex`, `rect`, `data: BinaryData`, `mimeType: "image/png" \| "image/jpeg"`                                     | insertion                      |
| `insertShape`   | `pageIndex`, `shape: "rectangle" \| "ellipse"` with `rect`, or `shape: "line"` with `from`, `to`; `stroke?`, `fill?` | insertion, colours             |
| `setShapeStyle` | `target` (shape), `stroke?: { color, width } \| null`, `fill?: { color } \| null`                                    | colours                        |
| `insertTable`   | `pageIndex`, `at: PagePoint`, `width`, `rows: string[][]`, `columnWidths?`, `style?: PdfTableStyle`                  | tables                         |
| `setTableCell`  | `target` (table), `row`, `column`, `text`                                                                            | tables                         |
| `moveElement`   | `target`, exactly one of `to: PagePoint` (new top-left of the bounds) or `by: { dx, dy }`                            | moving                         |
| `resizeElement` | `target` (text, textBox, image, shape, other), `rect`                                                                | moving                         |
| `deleteElement` | `target`                                                                                                             | moving                         |
| `insertPage`    | `index`, `size?: { width, height }` (default: the size of the neighbouring page)                                     | insertion                      |
| `deletePage`    | `pageIndex`                                                                                                          | pages                          |
| `movePage`      | `from`, `to`                                                                                                         | pages                          |
| `rotatePage`    | `pageIndex`, `rotation: 0 \| 90 \| 180 \| 270`                                                                       | pages                          |

```ts
export interface PdfTextBoxStyle {
  readonly fontFamily?: string; // "Helvetica" (default), "Times", "Courier" or a registered family
  readonly fontSize?: number; // 1–500 pt, default 12
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string; // default "#000000"
  readonly align?: "left" | "center" | "right"; // default "left"
  readonly lineHeight?: number; // 0.5–5, default 1.2
}

export interface PdfTableStyle {
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly color?: string;
  readonly borderColor?: string; // default "#000000"
  readonly borderWidth?: number; // default 0.75
  readonly cellPadding?: number; // default 4
  readonly headerFill?: string; // fill of the first row, none by default
}
```

Which operations an element accepts is listed in its `operations` field:

| Kind      | Operations                                                                     |
| --------- | ------------------------------------------------------------------------------ |
| `text`    | `replaceText`, `setTextStyle`, `moveElement`, `resizeElement`, `deleteElement` |
| `textBox` | `replaceText`, `setTextStyle`, `moveElement`, `resizeElement`, `deleteElement` |
| `image`   | `moveElement`, `resizeElement`, `deleteElement`                                |
| `shape`   | `setShapeStyle`, `moveElement`, `resizeElement`, `deleteElement`               |
| `table`   | `setTableCell`, `moveElement`, `deleteElement`                                 |
| `other`   | `moveElement`, `resizeElement`, `deleteElement`                                |

## Behaviour

### Engine and loading

- PDFium runs in a dedicated module worker, `dist/workers/pdf-edit-worker.js`,
  with `dist/assets/pdfium/pdfium.wasm`. Both are requested on the first
  `edit()` of a PDF and never before. One worker serves one session and is
  terminated when the session ends; calls into it are serialized.
- The worker loads the original bytes once. The edit state lives in that
  PDFium document; undo, redo and recovery rebuild it from the original bytes
  and the history (the `restore` step of the engine interface).

### Coordinates

- Page space is the page as PDF.js displays it at scale 1: points, origin at the
  top-left of the crop box, `y` down, `/Rotate` applied. The engine converts to
  and from PDF user space for every page.
- Inserted content appears upright on the displayed page, whatever the page's
  `/Rotate`.

### Element ids and persistence

- Ids are stable for the whole session — across undo, redo, page insertion,
  deletion and moves — and derived deterministically from the original page and
  object order and from the operation that created an element. **R2** The
  batch part of a created id is the history entry's `stateId`
  (`p0:n<stateId>.<op>.<k>`), which the core never reuses, so an id undone
  away is never given to a later element; a page created by a batch is keyed
  `q<stateId>.<op>` the same way. Deleting an element or a page reports every
  id it removes in `removedIds`; PDF never remaps ids.
- Text boxes and tables are parametric: their objects carry a marked-content
  tag (`WebDoc`) whose parameters store the element's kind and its inputs
  (rectangle, text, style, rows). Editing one rebuilds its objects from those
  parameters. After save and reopen in a new session they are listed again as
  one `textBox` or `table`. Tags from untrusted files are validated with the same
  schemas and limits as operations; an invalid tag makes its objects plain
  `text` and `shape` elements. **R2** A tag is also checked against what is
  actually drawn: the member objects' extracted text and the union of their
  bounds must match the stored inputs within tolerance, so a text box or
  table that another tool moved, resized or retyped degrades to plain
  objects instead of being rebuilt from stale inputs.

### Text and fonts

- `insertTextBox` lays out the text inside `rect.width`: paragraphs split on
  `\n`, greedy word wrapping (a word longer than the line is broken), the
  requested alignment and line height. Each line becomes one text object.
  Overflowing `rect.height` is allowed and reported as a warning.
- Fonts are chosen in this order: the standard PDF fonts (Helvetica, Times,
  Courier and their bold and italic variants) when every character is
  WinAnsi-encodable; otherwise a TrueType or OpenType font registered by the
  host through `ViewerClientOptions.fonts` whose family matches and covers the
  text; otherwise the bundled Unicode fallback font (see
  [Decisions](#decisions)). If no font covers the text, the operation fails
  validation with an issue coded `font-unavailable`.
- A missing bold or italic face falls back to the regular face with a
  `font-substitution` warning.
- `replaceText` on an existing `text` element keeps its font, size, colour and
  position, and changes it in place only when that is provably safe: the font is
  not a subset, or every new character already occurs in text drawn with the
  same font in this document. After changing it the engine reads the text back;
  a mismatch, or an unsafe case, replaces the object with a new text object in a
  fallback font at the same position, size and colour, with a
  `font-substitution` warning. **R2** As implemented, "safe" means glyph
  coverage of the embedded font: every code point of the new text (whitespace
  excepted) maps through the font's `cmap` to a glyph whose `glyf` entry has a
  non-zero length — a subset font can keep a `cmap` entry for an emptied
  glyph — and CFF-based fonts always fall back. A subset fixture with an
  emptied glyph proves it.
- `setTextStyle` on an existing `text` element supports `color` and `fontSize`
  (scaled around the top-left of its bounds); any other field is a validation
  issue coded `unsupported-style`. On a `textBox` every field is supported and
  the box is laid out again.
- MVP text is horizontal and left-to-right. Text that needs right-to-left or
  complex-script shaping fails validation with an issue coded
  `unsupported-script`.

### Images, shapes and tables

- JPEG data is embedded as is; PNG is decoded and embedded losslessly, with its
  alpha channel as a soft mask. Decoded pixels count against
  `maxDecodedPixels`, the data against `maxInputBytes`.
- Rectangles and lines are paths; ellipses are four Bézier curves. A missing
  `stroke` and `fill` together is a validation issue.
- `insertTable` draws the grid as paths and every cell as text wrapped to its
  column width; the row height follows the tallest cell. Tables are limited to
  100 rows and 20 columns in the MVP. `setTableCell` lays the table out again.

### Moving, resizing, deleting

- `moveElement` and `resizeElement` transform the element's objects; for
  `textBox`, `resizeElement` lays the text out again in the new rectangle
  instead of stretching it.
- `deleteElement` removes the element and, for groups, every object in it.
- Annotations are separate from page content and stay where they are, so a link
  over moved text no longer lines up; the docs mention it.

### Pages

- `insertPage` adds a blank page; `deletePage` refuses to delete the last page;
  `movePage` reorders pages; `rotatePage` sets an absolute rotation.
- Deleting a page removes its annotations; links that pointed to it stop
  working.

### Saving

- Only pages touched by an operation have their content regenerated; other
  pages keep their content streams.
- **R2** `save({ mode })` takes `"full"` or `"incremental"`. The default is
  `"full"` for a document without signature fields and `"incremental"` for a
  signed one. A full save rewrites the file (`FPDF_SaveAsCopy` without
  `FPDF_INCREMENTAL`), so deleted or replaced content is gone from the output
  — the privacy expectation of an enterprise host — and the bytes do not
  depend on which pages were read. An incremental save appends the changes to
  the original bytes, keeps earlier signed revisions intact, and leaves the
  old content recoverable; the docs say so. Without changes either mode
  returns the original bytes without calling PDFium.
- The viewer's reopen after a change (`materialize("show")`) uses the
  incremental form; the content is the same as the full form's. Whether it
  should use the full form instead is [open question 1](#open-questions).
- PDFium's incremental section repeats every object it has parsed, not only
  the changed ones (see [Spike results](#spike-results)), and
  `FPDFPage_GenerateContent` — the call that rewrites a changed page — parses
  every object of the document while checking whether the page's resources
  are shared (`IsPageResourceShared` in `cpdf_pagecontentgenerator.cpp`). So
  after any change an incremental save is the original file plus a copy of
  every object, about twice the original size, whatever pages were read
  (measured 2026-10-01, see [Compaction hardening](#compaction-hardening)).
  The engine still loads only the pages an operation or a query needs; a
  test proves a full save is identical with and without prior queries.
- Output is deterministic; a test guards it (see
  [Spike results](#spike-results)). **R2** A restore from a checkpoint (the
  bytes of an earlier state, incremental or full) yields the same ids and the
  same output as a replay from the original.
- A document with digital signatures can be edited; the first change emits a
  `fidelity-degraded` warning saying the signatures do not cover the new
  revision. **R2** The same first-change warning names, when present, a
  DocMDP certification (which any change invalidates), a tagged structure
  (inserted content is untagged) and a PDF/A claim (inserted standard fonts
  are not embedded); `details.features` lists them.

### Overlay primitives (ACTION-825, after revision 2)

A host overlay needs more than ids and bounds before a user can type into a
PDF. These session methods are specified here and implemented under
ACTION-825; they are additive, and the types they use (`TextPosition`,
`TextRange`, `ReadItem`, `ReadResult`) ship with `edit-core` revision 2.

- `getTextLayout(elementId, options?)` → `ReadItem<TextLayout>`: for a
  `text`, `textBox` or `table` element, its lines in reading order, each with
  its baseline, glyph boxes (`FPDFText_GetCharBox` mapped to the owning
  object), advance widths, font family, size and colour, and the `TextRange`
  of `EditElement.text` the line covers.
- `positionAt(pageIndex, point, options?)` → `ReadItem<TextPosition>`: the
  text position nearest to a page-space point (`FPDFText_GetCharIndexAtPos`).
- `rangeRects(range, options?)` → `ReadResult<PageRect>`: the rectangles a
  `TextRange` covers, one per line fragment, for drawing a selection.
- `renderPageWithout(pageIndex, elementIds, options)`: the page rendered by
  PDFium in the worker with the listed objects inactive
  (`FPDFPageObj_SetIsActive`), so the host's input surface can stand in for
  the element on screen without a reopen.
- `elementsForSelection(selection)` → `ReadResult<TextRange>`: maps the
  viewer's PDF.js selection to elements and ranges through an overlap ladder —
  rectangle overlap of at least 50 %, else the single containing object, else
  a text match with NFKC folding.
- `replaceText` gains an optional `range`; the engine splits the text object
  around the range only when it has to, keeping font, size, colour and
  baseline.
- Re-selection after a commit: stable ids and offsets let the host restore a
  `TextRange` after `documentchange`; `mapRange(range, fromRevision)` answers
  where a range went when a batch of this session moved its text.
- A main-thread per-page geometry cache: the last `getElements` result per
  page, refreshed on `documentchange`, read synchronously through
  `elementsAtSync(pageIndex, point)` so hover never waits behind an `apply()`.

### Viewer refresh

- After each change the viewer reopens the edited bytes with PDF.js. Each
  reopen starts a fresh PDF.js worker: PDF.js binds one `PDFWorker` to one
  loading task and destroys it with that task, and a port carries one
  `PDFWorker` at a time, so a worker cannot serve the new document while the
  old one is still displayed (found in task 9; the plan's worker reuse is
  dropped). The reopen cost is measured by the performance check in task 21.
- Target: a one-operation `apply()` on a 10-page PDF resolves within 1 second
  in Chromium on a developer machine. The browser suite fails above 3 seconds,
  to catch pathological regressions on slower CI machines.

## Out of scope

- **R2** The overlay text-input primitives are specified below under
  [Overlay primitives](#overlay-primitives-action-825-after-revision-2) and
  implemented under Linear ACTION-825, after this module's revision 2.
- Annotations: highlights, comments, ink, stamps; form filling.
- Redaction and OCR.
- Treating several text objects as one editable paragraph, or reflowing
  existing paragraphs.
- Bold or italic changes on existing text; vertical, right-to-left and
  complex-script text.
- Font subsetting (PDFium embeds whole font files).
- Digital signing. (A full save does rewrite the file and drops unreachable
  objects since revision 2; compaction beyond that — object renumbering,
  garbage collection of unused resources inside kept objects — stays out.)
- Encrypted PDFs (the viewer already refuses them).

## Work by layer

### Feature

- Add `@embedpdf/pdfium@2.15.1`; copy `pdfium.wasm` into
  `dist/assets/pdfium/`; update `THIRD_PARTY_NOTICES.md` with PDFium
  (BSD-3-Clause), the EmbedPDF wrapper (MIT; its repository has moved to
  Apache-2.0) and the libraries compiled into the WASM.
- Worker entry and RPC operations; the engine (object model, ids, coordinate
  conversion, operations, layout, fonts, images, marked-content tags, saving).
- `PdfEditSession` typed methods and the JSON Schemas of every operation.
- The PDF adapter's `edit` provider and reuse of the PDF.js worker on reopen.
- First task of the plan: a spike that proves, in Node and in Chromium,
  incremental saving through the wrapper's file writer, deterministic output,
  text objects with a standard and a TrueType font, inline JPEG loading through
  `addFunction`, marked-content parameters that survive save and reopen, and
  the page operations. A failed check updates this spec before work continues.

### Tests

- **Unit** (Node, PDFium WASM): every operation, then save, reopen with PDFium
  and check objects, text, bounds (±0.5 pt) and page count; coordinate
  conversion on pages rotated 0°, 90°, 180°, 270° and with an offset crop box;
  id stability across undo, redo and page moves; text-box and table tags after
  save and reopen; font selection, fallback and `font-unavailable`;
  determinism; save without changes equals the original; the output starts
  with the original bytes.
- **Browser** (vanilla example): every typed method through the public API; the
  changed area of the visible page re-renders; save, load the saved bytes in a
  fresh viewer and check the text through `getPageText`; no request for the
  worker or `pdfium.wasm` before `edit()`; a crashed worker surfaces as a typed
  error and the session can be restarted.
- **Fixtures:** PDFs generated inside the tests with PDFium (several pages,
  rotated pages, an offset crop box, text in standard fonts, an image), plus
  public corpus files for robustness.
- **R2 Unit:** full save drops deleted content and is identical with and
  without prior queries; ids after undo differ; a checkpoint restore equals a
  replay; a moved or retyped marked group degrades to plain objects; a subset
  TrueType fixture proves the in-place `replaceText` coverage rule; two text
  boxes embed one font file; warnings for DocMDP, tagged and PDF/A fixtures;
  `findText` ranges round-trip through `EditElement.text` offsets.
- **R2 Performance:** `apply()` latency of one `insertTextBox` on 10-, 100-
  and 500-page fixtures in Chromium, recorded under Actual result.

### Docs

- PDF section of `docs/api/editing.md`: every method and its fields, element
  kinds, page space, fonts and their limits, what stays unchanged (annotations,
  links), signatures, and the difference between `findText` and `search()`.

## Definition of done

- Every operation in the table above has a typed method, a JSON Schema, unit
  tests and a browser round-trip test.
- Inspection returns all six element kinds with correct bounds on rotated and
  cropped pages.
- Inserted Latin text uses the standard fonts; inserted Cyrillic text uses a
  registered font or the fallback font; both are extracted correctly by PDF.js
  after saving.
- Saving without changes returns the original bytes; undoing to revision 0
  returns the original bytes; edited output starts with the original bytes and
  is deterministic.
- No PDFium asset is requested before `edit()`, as a browser test proves. The
  size report measures the PDFium WASM, the edit worker and the fallback font,
  and stays within the 20 MiB Brotli target; the license gate passes.
- The PDF editing browser suite passes on Chromium, Firefox and WebKit.
- `npm run check` passes.
- **Overlay primitives (ACTION-825):** `getTextLayout`, `getPageLayout`,
  `positionAt`, `rangeRects`, `renderPageWithout`, `elementsForSelection`,
  `mapRange` and `elementsAtSync` exist with unit tests; where a line starts
  and ends agrees with the PDF.js text layer within 1 CSS px on rotated and
  cropped pages; `positionAt` and `rangeRects` round-trip; a suppressed
  render leaves the element out and the next normal render shows it; a
  browser test selects text, resolves it, replaces it with a ranged
  `replaceText` and restores the selection; `docs/api/editing.md` describes
  the primitives and the interaction model.
- **R2** The module passes on `edit-core` revision 2: envelopes, `stateId`
  ids, save modes, removed ids, ranges in `findText`, the staleness check and
  the extra warnings, with the latency numbers recorded.

## Spike results

Task 1 of the plan, 2026-10-01, `@embedpdf/pdfium` 2.15.1 in Node 22
(`packages/viewer/test/pdfium-bridge.test.ts`):

- **WASM in Node.** The package's ESM build runs in Node when the WASM bytes
  are supplied; instantiation takes about 12 ms. Unit tests can therefore drive
  the real engine without a browser. The bytes must always be supplied: the
  package default would fetch the module from a public CDN.
- **Incremental save works.** `FPDF_SaveAsCopy(document, writer,
FPDF_INCREMENTAL)` with the wrapper's `PDFiumExt_OpenFileWriter` returns the
  original bytes followed by an update; the added text reads back after
  reopening. The update repeats every object PDFium parsed (catalog, page tree,
  touched page dictionaries and their resources), not only the changed ones.
- **Deterministic output.** The same edit produced byte-identical files across
  two WASM instances, across saves a second and a half apart, and within one
  instance after unrelated allocations. PDFium derives the second `/ID` element
  without time or heap addresses, and it copies the original `/Info` dictionary
  instead of stamping a new date. No post-processing of the file id is needed;
  the determinism test stays as a guard.
- **Marked content.** A `WebDoc` mark with a JSON string parameter, including
  non-ASCII text, survives save and reopen and reads back unchanged.
- **Pages.** Insert, move, rotate and delete all survive an incremental save;
  a quarter-turned page reports its displayed (swapped) width and height.
- **Inline JPEG.** `FPDFImageObj_LoadJpegFileInline` works with an
  `FPDF_FILEACCESS` whose block reader is a WASM table callback created with
  `addFunction`; the JPEG is embedded without re-encoding.
- **Not covered yet:** TrueType loading (`FPDFText_LoadFont`) is probed with
  the fallback font in task 16, and the module worker in Chromium in task 9.

## Decisions

Resolved on 2026-10-01 together with the approval of this spec:

1. **Unicode fallback font.** A TrueType build of the bundled Noto Sans
   Latin/Cyrillic face (regular weight) ships next to the other packaged fonts
   (`packages/viewer/fonts/`, published as `dist/fonts/`, verified by the font
   manifest like the WOFF2 files) and is fetched only when inserted text needs
   it. Every saved
   PDF that uses it grows by roughly the font's size, because PDFium embeds
   whole fonts; subsetting is later work. Host-registered TrueType and OpenType
   fonts still take precedence.
2. **Granularity of existing text.** The MVP lists PDF text objects as they are
   stored, even when they are single words or fragments; grouping them into
   lines and paragraphs is later work.
3. **Signed PDFs.** Editing is allowed; the first change emits a
   `fidelity-degraded` warning, and the incremental save keeps the signed
   revision intact.
4. **R2, 2026-10-01.** Full save is the default for unsigned files (privacy;
   read-independent bytes); incremental stays for signed files and on request.
5. **R2, 2026-10-01.** The overlay text-input primitives are a separate
   ticket after ACTION-821.
6. **R2, 2026-10-01.** The in-place `replaceText` rule stays glyph-coverage
   based (the embedded font's `cmap`; CFF subsets always fall back) rather
   than GenOffice's ASCII-only rule; a subset fixture proves it.

## Open questions

1. **Reopen form after a change.** `materialize("show")` uses PDFium's
   incremental save, which after any change is the original bytes plus a copy
   of every object (about 2× the file; 103 MB for a 52 MB scan, in the same
   time as a full save: 857 ms against 825 ms). The full form with the
   compaction pass is the original size and renumbers objects, which nothing
   in the session depends on (ids live in marks). Keep the incremental form
   for the reopen, or switch to the full form? Raised by the peer review of
   2026-10-01; Leonid's call.
2. **Compaction failure behaviour.** When the compaction pass cannot read
   PDFium's full-save output (a cross-reference stream, a malformed object
   header, an unterminated stream), `save()` now fails closed with
   `edit-failed` and `details.reason: "pdf-compaction"`, and the session keeps
   its last committed bytes. The alternative is to return PDFium's
   uncompacted output with a warning, which keeps the save working but leaves
   deleted content recoverable in the file — the privacy property the full
   save exists for. PDFium's full save has produced only classic objects with
   a `trailer` in every run so far (corpus, fixtures, fuzz loop), so the
   branch is untaken in practice. Fail closed, or fall back with a warning?
   Raised by the peer review of 2026-10-01; Leonid's call.

## Actual result

- PDFium 2.15.1 runs in `dist/workers/pdf-edit-worker.js` with
  `dist/assets/pdfium/pdfium.wasm`; Node tests drive the same handler through a
  loopback worker. All 15 operations in the table above ship with a typed
  method on `PdfEditSession`, a JSON Schema in `pdfOperationSchemas`
  (version 1), unit tests and a browser round trip; a client that knows only
  the schemas performs all of them through `apply()`.
- Inspection returns the six element kinds with bounds in page space, verified
  against `FPDF_PageToDevice` on rotated and cropped pages. Ids are stable
  across undo, redo and page moves. Text boxes store their inputs in the
  `WebDoc` mark of every line object; tables store theirs once, on the grid and
  header-fill paths, and their cell text objects carry only the id, so a
  100×20 table does not repeat its contents. Invalid marks, and table members
  whose head mark is missing, are listed as plain objects.
- Fonts: the standard fonts for WinAnsi text, host-registered TrueType and
  OpenType fonts by family and `cmap` coverage, then the bundled Noto Sans
  Latin/Cyrillic TTF, fetched only when first needed. `replaceText` keeps an
  existing font when its `cmap` covers the new text and otherwise redraws at
  the same baseline in a covering font with a `font-substitution` warning.
  `FPDFFont_GetFamilyName` reports PDFium's substitute for non-embedded fonts,
  so the declared base font name is used instead.
- Images: JPEG inline through `FPDFImageObj_LoadJpegFileInline`; PNG decoded
  by the host (`createImageBitmap` and `OffscreenCanvas` in the worker, a small
  codec in Node tests) and stored as a BGRA bitmap, which PDFium writes with a
  soft mask. Headers are parsed before decoding so `maxInputBytes` and
  `maxDecodedPixels` reject oversized data early.
- Saving is incremental and deterministic; the signed-PDF warning comes from
  `FPDF_GetSignatureCount` on the first batch, and the signed revision's bytes
  are kept intact.
- Tests: 67 PDF unit tests (bridge, inspection, text boxes, transforms, pages,
  fonts, existing text, shapes, images, tables, hardening) in a viewer suite
  of 208; `tests/e2e/edit-pdf.spec.ts` holds 11 browser tests and the full
  matrix passes 100/100 on Chromium, Chromium at DPR 2, Firefox and WebKit.
  One `insertTextBox` on a ten-page file, including the PDF.js reopen, takes
  about 0.2 s in headless Chromium against the 3 s ceiling. The size gate
  reports 16.1 MiB Brotli with every optional font, under the 20 MiB target;
  the license gate and `npm run check` pass.
- Found along the way: `FPDF_MovePages` takes the destination as the resulting
  index; PDFium grows a stroked path's bounds by the full stroke width per
  side (half a point for a hairline); PDFium joins text objects that share a
  baseline into one extracted line; PDF.js cannot serve a reopened document
  from the worker that still shows the old one; the edit-core browser fixture
  had been missing `createSession` since the PDF session landed and was fixed
  during this close-out.
- Added beyond the spec: `Fields<T>` distributes over the operation union so
  `insertShape` keeps its per-shape fields; `signatureCount` on the engine
  model; `last-page`, `invalid-data` and `range` issue codes; `elementsAt`
  lists the top-most element first.
- Deferred, as listed under out of scope: annotations, paragraph reflow, bold
  or italic on existing text, font subsetting, right-to-left text.

### Revision 2 (ACTION-821)

- Ids of created elements derive from the history state id
  (`p0:n<stateId>.<op>.<k>`, page keys `q<stateId>.<op>`); `removedIds` lists
  deleted elements and every element of a deleted page; `$<n>` references are
  resolved while applying with the handler's own checks run on the resolved
  operation.
- `save({ mode })`: full is the default for unsigned files and goes through a
  compaction pass, because PDFium's full save still writes objects a
  regenerated page no longer references — the pass keeps the objects
  reachable from the trailer and rewrites the xref, so deleted content is
  gone and the bytes do not depend on which pages were read; incremental
  stays the default for signed files. The viewer reopens the incremental form.
- <a id="compaction-hardening"></a>**Compaction hardening (peer review,
  2026-10-01).** The pass is a byte-level PDF lexer (comments, literal and
  hex strings, names, numbers, keywords, delimiters) over PDFium's output: it
  delimits each object through its dictionary and stream — a direct
  `/Length`, an indirect one defined earlier in the file, or a measured
  `endstream` when the length is missing or wrong — collects `n g R`
  references, keeps the objects reachable from the trailer dictionary, and
  writes a fresh cross-reference table with one subsection per run of object
  numbers. Anything else (a cross-reference stream, a malformed header, an
  unterminated object) raises `PdfCompactionError` (`edit-failed`,
  `details.reason: "pdf-compaction"`); see open question 2. Proven by
  `test/pdf-edit-compact.test.ts`: hand-written files whose names, strings
  and comments spell `endobj` and `stream`, a reference split over lines, an
  indirect length defined after its stream, a bare `null` object, an orphan
  stream whose data says `endstream`; every public corpus PDF (two of them —
  `hello.pdf`, `pdfjs-cff-cid.pdf` — use object and cross-reference streams,
  which PDFium's full save turns into classic objects) edited, saved in full
  and reopened in PDFium (page count and per-page text equal to the original,
  the new box present) and in PDF.js 6.2.108 under Node (page count and
  first-page text); idempotence (compacting the output changes nothing); and
  `compactPdf` as a target of `scripts/fuzz-js.mjs` (2,000 mutations of a
  classic PDF seed among others, slowest case 1.4 ms, no hang). The corpus
  run found and fixed one fault: offsets that pointed at the whitespace before
  each object (byte 0 for the first), which PDFium repaired by rebuilding the
  table and, on `pdfjs-arabic-cid-truetype.pdf`, lost the page content on the
  way. Sizes and times, Node 22 on an Apple M4 Pro, one `insertTextBox` then
  a save, PDFium's save and the compaction pass timed separately:

  | File                          | Original | Show (incremental) | Full save | + compaction | PDFium + pass |
  | ----------------------------- | -------: | -----------------: | --------: | -----------: | ------------: |
  | hello.pdf                     |  3,560 B |            7,759 B |   4,520 B |      4,157 B |  0.2 + 0.5 ms |
  | pdfjs-arabic-cid-truetype.pdf | 39,370 B |           78,412 B |  38,713 B |     38,667 B |  0.3 + 0.9 ms |
  | pdfjs-cff-cid.pdf             | 17,022 B |           34,134 B |  17,435 B |     17,078 B |  0.1 + 0.2 ms |
  | pdfjs-complex-truetype.pdf    | 57,178 B |           90,498 B |  31,377 B |     31,353 B |  0.8 + 0.2 ms |
  | pdfjs-mmtype1.pdf             |  8,044 B |           16,392 B |   8,246 B |      8,225 B |  0.1 + 0.2 ms |
  | pdfjs-noembed-jis7.pdf        | 15,306 B |           31,418 B |  15,589 B |     15,567 B |  0.1 + 0.4 ms |
  | pdfjs-standard-fonts.pdf      | 94,800 B |          191,863 B |  93,211 B |     93,155 B |  0.6 + 1.2 ms |
  | 200 pages of raw RGB images   |  51.6 MB |  103.2 MB (825 ms) |   51.6 MB |      51.6 MB |    821 + 7 ms |

  Every corpus file saves in full, compaction included, in under 2 ms; the
  compaction pass costs 7 ms on the 51.6 MB file against 821 ms for PDFium's
  own save (and 825 ms for the incremental one). PDFium's full save is not
  always smaller than the original: it writes classic objects for files that
  used object streams (`hello.pdf`, `pdfjs-cff-cid.pdf`).

- The first-change warning names signatures, a DocMDP certification, a
  tagged structure (raw `/MarkInfo` or a structure tree on the first page)
  and a PDF/A claim (`pdfaid:part` in the uncompressed XMP) in
  `details.features`.
- A marked text box or table whose objects no longer match its stored inputs
  (text compared whitespace-normalized; bounds within the box's first line
  band and width, or the table's frame) is listed as plain objects.
- `findText` returns `ranges` with offsets into each element's text; lines of
  a text box and cells of a table are located inside the element's text.
- In-place `replaceText` requires every character to map through the
  embedded font's `cmap` to a glyph with outline data; the fallback font is
  also fetched for families only it can stand in for. PDFium embeds the
  fallback TrueType once per document (`/FontFile`), proven by a test.
- Latency of one `insertTextBox` through `apply()` — engine, incremental
  save, PDF.js reopen and the page-count update — on fixture PDFs with one
  text object per page, headless Chromium (Playwright) on an Apple M4 Pro,
  2026-10-01: 10 pages 46 ms; 100 pages 130 ms for the first apply and 40 ms
  for the second; 500 pages 72 ms and 56 ms. The first apply on a document
  pays for PDF.js parsing it again from scratch; the steady state is well
  inside the 1 s target and the 3 s ceiling the browser suite asserts.
  (`tests/e2e/edit-pdf.spec.ts`, "measures apply() latency on a …-page PDF".)

### Overlay primitives (ACTION-825)

Implemented 2026-10-01 to 2026-10-02 (tasks 32–37 of `tasks/todo.md`), all
additive to the revision-2 session:

- `getTextLayout(elementId)` and `getPageLayout(pageIndex)` read PDFium's
  text page once and map every character to its object and element (the
  `findText` mapping): a line is one text object, so a text box's drawn lines
  and a table's cells come out in drawing order, each with its `range` into
  the element's text, tight glyph boxes (`FPDFText_GetCharBox`), advances
  (`FPDFText_GetLooseCharBox`), baseline (`FPDFText_GetCharOrigin`), font
  family, size and colour, all through `PageGeometry` so rotation and crop
  boxes hold. `positionAt` uses `FPDFText_GetCharIndexAtPos` with a 2 pt
  tolerance, then the nearest glyph box, and puts the caret after a glyph
  past its middle in reading direction. `rangeRects` unions glyph boxes per
  line across the elements a range spans in reading order.
- `renderPageWithout` marks the elements' objects inactive
  (`FPDFPageObj_SetIsActive`), draws the page with `FPDF_RenderPageBitmap`
  into a BGRA bitmap, converts it to RGBA, restores the flags, and transfers
  the pixels; bitmaps over `maxDecodedPixels` are refused with
  `resource-limit`. The session's bytes do not change.
- `elementsForSelection` resolves the viewer's PDF.js runs through the
  ladder (box overlap ≥ 50 % of the line, else the single containing line,
  else an NFKC-folded text match); the run's text decides the offsets, since
  the viewer cuts runs proportionally, and the covered glyphs pick between
  repeated words. GenOffice's technique, attributed in
  `THIRD_PARTY_NOTICES.md`.
- `replaceText` with `range`: a text box splices and reflows; a text object
  is rewritten in place when its font covers the spliced text, else split
  into up to three objects on the same baseline (before and after in the
  original font, the middle in the fallback), each starting where the
  previous one's advance ends; the first part keeps the id. Every part is
  read back before the old object goes; a part the original font cannot
  encode sends the whole object to the fallback as before.
- `mapRange` walks a log of the session's committed calls (operations and
  receipts, 512 entries) forwards for batches and redos and backwards for
  undos and resets, then clamps against the live element.
- `elementsAtSync` answers from a per-page cache of the last
  `getElements({ pageIndex })` result, refreshed for changed pages after
  every commit.
- Core plumbing: `EditSessionCore.readItem` / `readItems` queue a typed
  session's engine reads behind earlier calls and stamp the envelope; five
  worker requests (`edit-text-layout`, `edit-position-at`,
  `edit-range-rects`, `edit-render-without`, `edit-page-layout`).
- Found on the way: PDFium's text extraction appends a generated space to a
  text object that a gap follows, so `EditElement.text` of such an object
  ends with a space the file does not contain; the split path tolerates it
  when reading parts back.
- Tests: 28 unit tests across `pdf-edit-layout`, `pdf-edit-render`,
  `pdf-edit-selection`, `pdf-edit-replace-range` and `pdf-edit-session`
  (selections built from real PDF.js runs under Node); browser spec
  `tests/e2e/edit-pdf-overlay.spec.ts` (7 tests): line start and end within
  1 CSS px of the PDF.js text-layer span on pages rotated 0/90/180/270 and
  on a cropped page, the mouse-selection → `elementsForSelection` →
  `replaceText` → `mapRange` → `selectText` flow, and a suppressed render
  laid over the page with the normal render unchanged.
