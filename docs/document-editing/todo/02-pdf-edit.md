# Module 02. `pdf-edit` — PDF editing on PDFium

**Status:** Approved 2026-10-01; implementation not started

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
  object order and from the operation that created an element.
- Text boxes and tables are parametric: their objects carry a marked-content
  tag (`WebDoc`) whose parameters store the element's kind and its inputs
  (rectangle, text, style, rows). Editing one rebuilds its objects from those
  parameters. After save and reopen in a new session they are listed again as
  one `textBox` or `table`. Tags from untrusted files are validated with the same
  schemas and limits as operations; an invalid tag makes its objects plain
  `text` and `shape` elements.

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
  `font-substitution` warning.
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
- `save()` writes an incremental update: the original bytes followed by the
  appended changes, so earlier signed revisions stay intact. Without changes
  it returns the original bytes without calling PDFium.
- PDFium's incremental section also repeats every object it has parsed, not
  only the changed ones (see [Spike results](#spike-results)). The engine
  therefore loads only the pages an operation or a query needs.
- Output is deterministic; a test guards it (see
  [Spike results](#spike-results)).
- A document with digital signatures can be edited; the first change emits a
  `fidelity-degraded` warning saying the signatures do not cover the new
  revision.

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

- Annotations: highlights, comments, ink, stamps; form filling.
- Redaction and OCR.
- Treating several text objects as one editable paragraph, or reflowing
  existing paragraphs.
- Bold or italic changes on existing text; vertical, right-to-left and
  complex-script text.
- Font subsetting (PDFium embeds whole font files).
- Digital signing; compacting saves that rewrite the whole file.
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
