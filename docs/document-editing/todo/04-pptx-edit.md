# Module 04. `pptx-edit` — PPTX editing on the package layer

**Status:** ✅ Done 2026-10-02 (T44–T49, Linear ACTION-812), except the
ticket's release criterion, which is Leonid's call. Drafted and implemented
under the instruction of 2026-10-02 to execute the Linear plan without
stopping; the decisions below were taken in the draft and are open to review.

## Goal

Give hosts and AI clients the PPTX method set of the roadmap through the same
`EditSession` contract as PDF: inspect the shapes of a slide, change text and
formatting, recolour, move, resize and delete shapes, insert text boxes,
images and tables, and add, duplicate, delete and reorder slides. Every edit
is a patch to the slide XML through `ooxml-package` (module 03); the viewer
reopens the edited bytes with the same `@silurus/ooxml` 0.88 engine it already
uses, so the canvas always shows the file `save()` returns.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md): fixed decisions 1 (API only,
  overlay editing with commit on blur or idle), 2 (one operation vocabulary),
  3 (the original file is the source of truth), 4 (rendering stays with the
  existing engines), 6 (PPTX is edited XML-first), 7 (`pptm` and `ppsx`
  follow the base format; macro parts copied and never executed), 9 (order);
  the MVP method set by category; the hard limits (no live reflow; "what you
  see is what you save" holds against web-doc's renderer, PowerPoint may lay
  the same bytes out differently).
- [`./01-edit-core.md`](./01-edit-core.md): `EditEngine`, `EngineBatch`
  with `stateId`, `restore({ base, batches })`, `materialize("show" | "save")`,
  `EditElement` with `frame`, `story` and `parentId`, `TextRange` offsets in
  UTF-16 code units of `EditElement.text`, `ThemeColor`, the 500-page latency
  requirement of revision 2 ("the Office formats when their engines exist").
- [`./03-ooxml-package.md`](./03-ooxml-package.md): `OoxmlPackage`,
  transactions, patches with read-back, `addMedia`, `uniquePartName`,
  snapshots, decisions 2 (changed entries stored), 5 (overlay state), 8 (no
  worker of its own) and 9 (dangling relationship targets warn).
- Linear ACTION-812: text in shapes, placeholders and table cells; run
  formatting (bold, italic, underline, size, font), paragraph alignment, text
  colour; shape fill and line; move, resize, delete (`a:xfrm`); insert a text
  box, an image (media part + relationship) and a table; add from a layout,
  duplicate, delete and reorder slides (`p:sldIdLst` + relationships);
  inspection with stable ids (`p:cNvPr/@id` + slide), bounds in slide space,
  text and style; `.pptm` / `.ppsx` on the same path; every operation visible
  after the call and surviving save and reload; saved files open without a
  repair prompt in PowerPoint and Keynote (manual, fixture set); unit and
  browser tests per method; `apply()` latency on 10-, 100- and 500-slide
  decks recorded in `docs/api/editing.md`.
- ECMA-376 Part 1 (PresentationML and DrawingML): `p:sldIdLst/p:sldId`
  (`@id` ≥ 256, unique; `@r:id` into the presentation's relationships),
  `p:cSld/p:spTree` with `p:sp`, `p:pic`, `p:graphicFrame`, `p:grpSp`,
  `p:cxnSp`; `p:nvSpPr/p:cNvPr/@id` unique within a slide; `p:nvPr/p:ph`
  (`@type`, `@idx`) inheriting position and text style from the layout, then
  the master; `a:xfrm` (`a:off`, `a:ext` in EMU, `@rot` in 60,000ths of a
  degree, `@flipH`, `@flipV`; `a:chOff`/`a:chExt` child space of a group);
  `a:txBody` (`a:bodyPr`, `a:lstStyle`, `a:p` with `a:pPr`, `a:r` with
  `a:rPr` and `a:t`, `a:br`, `a:fld`, `a:endParaRPr`); colours as
  `a:srgbClr` or `a:schemeClr` with modifiers; `a:tbl` inside
  `a:graphic/a:graphicData`; `p:pic` with `a:blip/@r:embed`.
- Research of 2026-10-01 (ideas only, no code):
  - GenOffice (Apache-2.0; attributed in `THIRD_PARTY_NOTICES.md`): text
    edits trace runs back to their source so formatting survives, rebuild
    only the paragraphs they must and keep `a:bodyPr` and `a:lstStyle`;
    provenance flags stop a rebuild from baking inherited values into the
    file; picture insertion deduplicates media, adds a content-type Default
    and the next `rId`; slide insertion touches the Override, the
    presentation relationship and `p:sldId`; a stored `a:normAutofit`
    `fontScale` goes stale after a text edit; validation against the schema
    counts only violations an edit introduced.
  - The `@silurus/ooxml` 0.88 reading: no mutation or model API; shapes carry
    `cNvPr` id and name, pictures and tables only the id; groups are
    flattened with the transform baked in; hidden shapes are dropped; a
    stored `fontScale` is always applied; `getElementBoundsByIds` returns EMU
    frames with rotation and flips; `collectSlideRuns` merges adjacent runs
    and carries `shapeId`, which web-doc's adapter currently discards.
  - Maxgent's PPTX editor fork (MIT): addresses elements by part name plus
    `cNvPr` id and limits edits to slide-origin top-level shapes because
    groups and hidden nodes broke its mapping — a warning this module heeds
    by walking the XML itself.

## Dependencies

- Module 01 `edit-core` (revision 2) and module 03 `ooxml-package`, both
  done.
- `@silurus/ooxml-pptx` (`@silurus/ooxml` 0.88.0) as the renderer, unchanged.
  No new runtime dependency.

## In scope

### Inspection

`getElements`, `getElement`, `elementsAt` and `findText` from `edit-core`,
returning `PptxElement` values; two PPTX reads for slides and layouts.

```ts
export type PptxElementKind =
  | "shape" // p:sp — a text box, a placeholder, an auto shape, WordArt
  | "image" // p:pic
  | "table" // p:graphicFrame holding a:tbl
  | "connector" // p:cxnSp
  | "group" // p:grpSp; its children carry parentId
  | "other"; // charts, diagrams, OLE objects, media frames

export interface PptxElement extends EditElement {
  readonly kind: PptxElementKind;
  /** `p:cNvPr/@name`, as PowerPoint shows it in the selection pane. */
  readonly name: string;
  /** Present for a placeholder: `p:ph/@type` (default "body") and `@idx`. */
  readonly placeholder?: { readonly type: string; readonly idx?: number };
  /** First run with text, resolved through the placeholder chain and the theme. */
  readonly textStyle?: PptxTextStyle; // shape with a text body
  readonly shapeStyle?: PptxShapeStyle; // shape, connector
  readonly table?: { readonly rows: readonly (readonly string[])[] }; // table
}

export interface PptxTextStyle {
  readonly fontFamily: string; // theme fonts resolved ("+mn-lt" → the minor Latin face)
  readonly fontSize: number; // points
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly color: EditColor; // "#RRGGBB", or { theme, mods } when the file uses a scheme colour
  readonly align: "left" | "center" | "right" | "justify";
}

export interface PptxShapeStyle {
  /** "none" for a:noFill; absent when inherited from the shape style or the theme. */
  readonly fill?: EditColor | "none";
  readonly line?: {
    readonly color: EditColor | "none";
    readonly width: number;
  }; // points
}

export interface PptxSlideInfo {
  readonly pageIndex: number;
  /** Stable for the session: the slide part's number, as in element ids. */
  readonly key: string; // "sld3"
  readonly layout: string; // a PptxLayoutInfo id
  readonly hidden: boolean;
}

export interface PptxLayoutInfo {
  readonly id: string; // "layout2"
  readonly name: string; // p:cSld/@name, e.g. "Title and Content"
  readonly type?: string; // p:sldLayout/@type when present
  readonly master: string; // "master1"
}
```

`EditElement` fields are used as the core defines them: `bounds` is the
axis-aligned box of the shape in slide space, `frame` the untransformed box
with `rotation`, `flipH` and `flipV`, `text` the shape's text in the text
model below, `parentId` the enclosing group, `story` absent (slide body;
layouts, masters and notes are not editable in this module). `findText`
searches the text model of every shape and table cell of the edited deck and
returns one `TextTarget` per match with the shape's frame as its rectangle
(the engine has no glyph geometry; the docs say so, as they do for PDF's
`findText` versus `search()`).

Two reads on `PptxEditSession`, through `readItems`:

| Read           | Returns                      | Notes                                           |
| -------------- | ---------------------------- | ----------------------------------------------- |
| `getSlides()`  | `ReadResult<PptxSlideInfo>`  | In presentation order; keys survive reordering  |
| `getLayouts()` | `ReadResult<PptxLayoutInfo>` | Every layout of every master, for `insertSlide` |

### Operations

`PptxEditSession` exposes one typed method per operation, each a
single-operation `apply()`. Geometry is in slide space (CSS pixels at 96 dpi,
`EMU / 9525`); font sizes and line widths are in points.

| Operation        | Fields                                                                                                                   | Category                       |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------ | ------------------------------ |
| `replaceText`    | `target` (shape), `text`, `range?: TextRange` (both ends on the target; a collapsed range inserts)                       | text                           |
| `setTextStyle`   | `target` (shape), `range?`, `style: PptxTextStyleChange`                                                                 | formatting, alignment, colours |
| `setShapeStyle`  | `target` (shape, connector), `fill?: EditColor \| "none" \| null`, `line?: { color, width? } \| "none" \| null`          | colours                        |
| `moveElement`    | `target`, exactly one of `to: PagePoint` (new top-left of `bounds`) or `by: { dx, dy }`                                  | moving                         |
| `resizeElement`  | `target`, `rect: PageRect` (new `bounds`; the rotation is kept)                                                          | moving                         |
| `deleteElement`  | `target`                                                                                                                 | moving                         |
| `insertTextBox`  | `pageIndex`, `rect`, `text`, `style?: PptxTextStyleChange`                                                               | insertion                      |
| `insertImage`    | `pageIndex`, `rect`, `data: BinaryData`, `mimeType: "image/png" \| "image/jpeg"`                                         | insertion                      |
| `insertTable`    | `pageIndex`, `rect`, `rows: string[][]`, `columnWidths?: number[]` (weights), `style?: PptxTableStyle`                   | tables                         |
| `setTableCell`   | `target` (table), `row`, `column`, `text`                                                                                | tables                         |
| `insertSlide`    | `index` (0 to the slide count), `layout?: string` (a layout id; default: the layout of the slide before, else the first) | slides                         |
| `duplicateSlide` | `pageIndex`, `index?` (default: right after the source)                                                                  | slides                         |
| `deleteSlide`    | `pageIndex` (the last slide cannot be deleted)                                                                           | slides                         |
| `moveSlide`      | `from`, `to` (the slide's index after the move)                                                                          | slides                         |

```ts
export interface PptxTextStyleChange {
  readonly fontFamily?: string; // written as a:latin/@typeface; theme names ("+mn-lt") allowed
  readonly fontSize?: number; // 1–400 pt
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean; // a:rPr/@u = "sng" | "none"
  readonly color?: EditColor; // "#RRGGBB" → a:srgbClr; { theme, mods } → a:schemeClr with modifiers
  readonly align?: "left" | "center" | "right" | "justify"; // a:pPr/@algn on the paragraphs in range
}

export interface PptxTableStyle {
  /** Header-row and banded-row flags of the table; both default to true. */
  readonly firstRow?: boolean;
  readonly bandRow?: boolean;
}
```

Which operations an element accepts is listed in its `operations` field:

| Kind        | Operations                                                                                                         |
| ----------- | ------------------------------------------------------------------------------------------------------------------ |
| `shape`     | `replaceText`, `setTextStyle` (with a text body), `setShapeStyle`, `moveElement`, `resizeElement`, `deleteElement` |
| `image`     | `moveElement`, `resizeElement`, `deleteElement`                                                                    |
| `table`     | `setTableCell`, `moveElement`, `resizeElement`, `deleteElement`                                                    |
| `connector` | `setShapeStyle` (line only), `moveElement`, `resizeElement`, `deleteElement`                                       |
| `group`     | `moveElement`, `resizeElement`, `deleteElement`                                                                    |
| `other`     | `moveElement`, `resizeElement`, `deleteElement`                                                                    |

`EditColor` is the core's union: `#RRGGBB` or `{ theme, mods }` where `theme`
is a scheme slot (`tx1`, `bg1`, `accent1` … `accent6`, `hlink`, `folHlink`,
or the master-mapped `dk1`, `lt1`, `dk2`, `lt2`) and `mods` maps DrawingML
modifier names (`lumMod`, `lumOff`, `tint`, `shade`, `alpha`, `satMod`) to
their values in thousandths of a percent, as the file stores them.

## Out of scope

- Overlay text-input primitives (`getTextLayout`, `positionAt`,
  `rangeRects`, `renderPageWithout`, `elementsForSelection`, `mapRange`): a
  separate ticket after this module, as ACTION-825 was for PDF. They need
  the adapter to keep the renderer's `shapeId`, `font` and frame on text
  runs, which this module does not change.
- Editing layouts, masters, notes, comments, sections, transitions,
  animations and headers/footers; master-view editing.
- Formatting of table cells beyond their text (borders, fills, merges), row
  and column insertion; charts, SmartArt and OLE objects beyond move, resize
  and delete; connectors' connection sites.
- Re-deriving `a:normAutofit` font scales (decision 9 keeps the file
  consistent without it).
- Embedded font subsets: text in a new font family is written by name only;
  PowerPoint substitutes when the font is absent.
- Password-protected and ZIP64 packages (refused by the package layer).

## Behaviour

### Engine and loading

- The engine runs in a dedicated module worker,
  `dist/workers/ooxml-edit-worker.js`, requested on the first `edit()` of a
  PPTX and never before; one worker serves one session and is terminated
  when the session ends. The worker holds the `OoxmlPackage`; the main
  thread holds a client that implements `EditEngine` over the existing
  worker RPC (`edit-init`, `edit-open`, `edit-validate`, `edit-apply`,
  `edit-materialize`, `edit-restore`, `edit-put-asset`, `edit-elements`,
  `edit-element`, `edit-elements-at`, `edit-find-text`, `edit-dispose`, plus
  `edit-pptx-slides` and `edit-pptx-layouts`). The worker is shared with
  module 06: the open payload names the format and the worker instantiates
  the matching engine.
- The engine class (`PptxEditEngine`) is thread-agnostic and is what the
  unit tests drive in Node; the worker and the PDF-style loopback fixture
  wrap it. The PDF worker client's transport is generalized into a base
  class both formats use; the PDF client keeps its overlay reads on top.
- `OfficeDocumentAdapter` gains an `edit` provider for `pptx`, `pptm` and
  `ppsx` (the session's `format` is `"pptx"` for all three; the core maps
  the variants) and a `reopen` that loads the edited bytes into a fresh
  `PptxPresentation` with the same options as `open`. Nothing of the
  previous handle is reusable (the engine owns its worker), so `reopen` is
  `open` plus the closing of the old handle by the viewer. `progressiveLayout`
  is turned on for reopens if the spike shows it shortens the time to the
  first painted slide without changing what `getPageText` returns.
- The state is the original bytes plus the package overlay; `restore` opens
  `base` as a new package and replays the batches; `materialize` is the
  package's `save()`; no changes returns the original bytes (module 03).
  `materializeDocument` carries the package's warnings (dangling
  relationship targets, which this module avoids by construction).

### Coordinates

- Slide space is CSS pixels at 96 dpi: `px = EMU / 9525`, matching the
  viewer's `pageSizes` for presentations. The engine writes EMU values
  rounded to the nearest integer, so a move followed by a read returns the
  position to within 1/9525 px.
- A shape's `frame` is its own `a:xfrm` when present; a placeholder without
  one inherits from the layout placeholder with the same `idx`, else the
  same `type`, else the master's. A group child's frame is mapped from the
  group's child space (`a:chOff`/`a:chExt`) through every enclosing group.
  `bounds` is the axis-aligned box of the rotated and flipped frame.
- `moveElement` and `resizeElement` write an explicit `a:xfrm` on the
  target (creating one on an inheriting placeholder, with the inherited
  extent), converting slide space back through the group chain for group
  children. A group's own frame is not recomputed when a child moves
  (PowerPoint renders children by their child-space coordinates); resizing
  a group scales `a:ext` and leaves `a:chExt` alone, so its children scale
  with it as in PowerPoint.
- `elementsAt` returns the elements whose rotated frame contains the point,
  topmost first (reverse document order), groups after their children.

### Element ids and persistence

- An element id is `<slideKey>:<cNvPrId>`, for example `sld3:7`: the number
  of the slide part (`ppt/slides/slide3.xml`) and `p:cNvPr/@id`. Ids are
  therefore stable across reordering and across sessions on the same file,
  and are what a host should persist. A file with duplicate `cNvPr` ids on
  one slide (invalid but seen in the wild) gets `#2`, `#3` suffixes on the
  later duplicates, in document order.
- Elements created by the session get the next free `cNvPr` id of their
  slide (one more than the largest in the slide, layouts and master
  included, as PowerPoint does); a new slide gets the next free slide number
  of the package (one more than the largest among the original and current
  slide parts). Both are functions of the package state, so a replay after
  undo reproduces them. A slide number is reissued only after a slide the
  session created was deleted; the ids of that slide's elements were reported
  in `removedIds`, which hosts must treat as final. This is the one place
  the module relaxes the core's "never reused" rule; the docs state it.
- `remappedIds` is never used: no operation renames a shape.
- `removedIds` lists every element of a deleted slide and of a deleted
  group; `changedPages` lists the slide of every touched element, every
  slide from the insertion or deletion point onwards for slide operations,
  and both ends of a move.

### Text model

- `EditElement.text` of a shape is its paragraphs joined by `\n`; inside a
  paragraph `a:br` is `\v` (U+000B, PowerPoint's own line-break character),
  `a:t` text is itself, a tab run is `\t`, and `a:fld` contributes its cached
  text. A table's `text` is its cells in reading order joined by `\t` within
  a row and `\n` between rows; `table.rows` carries them separately.
- `TextRange` offsets index this text in UTF-16 code units. A range that
  starts or ends inside a field is `invalid-range`: fields are replaced as a
  whole or not at all.
- `replaceText` without a range replaces the whole text body: paragraphs
  are rebuilt from the new text (`\n` splits paragraphs, `\v` becomes `a:br`),
  every new paragraph takes the `a:pPr` of the first original paragraph and
  every run the `a:rPr` of the first original run that had text; `a:bodyPr`,
  `a:lstStyle` and the last paragraph's `a:endParaRPr` are kept as they are.
  An empty text leaves one empty paragraph with the kept `a:endParaRPr`.
- `replaceText` with a range is minimal: when the range lies inside one run,
  only that `a:t` changes; when it spans runs, the first touched run keeps
  its `a:rPr` and receives the new text, the other covered runs are
  shortened or removed, and `\n` in the new text splits the paragraph at
  that point (the new paragraph copies the `a:pPr`). Runs the range does
  not touch keep their bytes (GenOffice's source-run idea, without its
  rebuild).
- `setTextStyle` splits runs at the range ends and writes only the
  properties given: `b`, `i`, `u`, `sz` (hundredths of a point), `a:latin`
  (`a:ea` and `a:cs` left alone), `a:solidFill`. `align` is written as
  `a:pPr/@algn` on every paragraph the range touches (all of them without a
  range). Properties not in the change keep their bytes, so inherited values
  are never baked in.
- Fonts: `fontFamily` is written as given; `+mj-lt`, `+mn-lt` and the other
  theme font names pass through. Reported `fontFamily` resolves theme names
  through `ppt/theme/theme1.xml` of the slide's master.
- Autofit (decision 9): when the text or text style of a shape whose
  `a:bodyPr` holds `a:normAutofit` changes, the stored `fontScale` and
  `lnSpcReduction` attributes are dropped (the element stays), so neither
  web-doc's renderer nor PowerPoint applies a scale computed for the old
  text; PowerPoint recomputes it on its next layout of the shape.
- Characters forbidden in XML 1.0 (`\u0000`–`\u0008`, `\u000E`–`\u001F`,
  lone surrogates) are rejected with `invalid-text`; `<`, `&` and `>` are
  escaped by the package layer.

### Shapes

- `setShapeStyle` writes `a:solidFill` or `a:noFill` into `p:spPr` for the
  fill and `a:ln` with `a:solidFill`/`a:noFill` and `@w` (points × 12,700)
  for the line; `null` removes the explicit value so the shape falls back to
  its `p:style` reference or the theme; absent keeps the bytes. The element
  order of `p:spPr` (`a:xfrm`, geometry, fill, `a:ln`, effects) is
  respected when a child is created. A shape inside `mc:AlternateContent`
  is edited in its `mc:Choice` branch only when the branch is the one the
  renderer draws (the `p:sp` outside the fallback); otherwise the shape is
  reported with `operations: []`.
- `insertTextBox` creates a `p:sp` with `txBox="1"`, `a:prstGeom prst="rect"`,
  `a:noFill`, `a:bodyPr wrap="square" rtlCol="0"` with `a:spAutoFit`, an
  empty `a:lstStyle`, and paragraphs from `text` with the given style on
  every run; the name is `TextBox <id>` as PowerPoint names them.
- `deleteElement` removes the element's XML and the relationships of the
  slide that only it referenced (`r:embed`, `r:link`, `r:id` attributes in
  its subtree). Parts those relationships pointed at (images, charts) are
  left in the package when no other relationship references them; the saved
  file is valid and PowerPoint drops orphans on its own save. The docs say
  so.
- Hidden shapes (`p:cNvPr/@hidden="1"`) are listed with their bounds and
  accept every operation; the renderer does not draw them, which the docs
  note.

### Images and tables

- `insertImage` stores the bytes through the package layer's `addMedia`
  (SHA-256 deduplication, `ppt/media/image<n>.<ext>`, a content-type
  Default for the extension), adds an `image` relationship from the slide
  and inserts a `p:pic` with `a:blip r:embed`, `a:stretch/a:fillRect`,
  `a:prstGeom prst="rect"` and the given frame. The image is placed in
  `rect` as given; the engine does not read the image's pixel size.
- `insertTable` inserts a `p:graphicFrame` with the table graphic:
  `a:tblPr` with `firstRow`/`bandRow` and the deck's default table style id
  when `ppt/tableStyles.xml` defines one, `a:tblGrid` from `columnWidths`
  (equal weights by default) scaled to `rect.width`, rows of equal height
  summing to `rect.height`, one paragraph per cell with the cell text. The
  element is a `table`.
- `setTableCell` rebuilds the paragraphs of one `a:tc` from `text` (the
  text model's `\n` and `\v` rules), keeping the cell's `a:tcPr` and the
  first run's `a:rPr`; empty text leaves an empty paragraph.

### Slides

- `insertSlide` creates `ppt/slides/slide<N>.xml` from the chosen layout: a
  `p:sld` whose `p:spTree` holds one empty `p:sp` per layout placeholder
  that PowerPoint instantiates (type `title`, `body`, `subTitle`, `ctrTitle`,
  `obj`, `pic`, `tbl`, `chart`, `dgm`, `media`, `clipArt`; date, footer and
  slide-number placeholders are not copied, as PowerPoint does), each with
  the layout placeholder's `p:ph` and an empty text body; a `.rels` part
  with the `slideLayout` relationship; a content-type Override; a
  `presentation.xml.rels` relationship; and a `p:sldId` (`@id` one more than
  the largest, at least 256) at `index` in `p:sldIdLst`. `docProps/app.xml`
  slide counts are not updated (PowerPoint ignores them).
- `duplicateSlide` copies the slide part byte for byte to a new slide
  number and clones what the copy must own: its `.rels` part; parts the
  relationships reach that PowerPoint does not allow two slides to share
  (charts, diagrams, embeddings and their own `.rels`, recursively, with
  `uniquePartName`); images, media and the layout are shared by relationship.
  Notes slides and comments are not copied. The new `p:sldId` goes at
  `index`.
- `deleteSlide` removes the `p:sldId`, the presentation relationship, the
  slide part with its `.rels` and Override, and the slide's notes slide part
  (with its `.rels` and Override). Media and layouts stay. The last slide
  cannot be deleted (`invalid-operation`).
- `moveSlide` reorders `p:sldIdLst` only.
- Layout ids are `layout<N>` from the part number, master ids `master<N>`;
  `insertSlide` with an unknown layout fails validation with
  `unknown-layout`.

### Placeholders and inheritance

- Position, text style and colours of a placeholder are inherited from the
  layout and the master, and `textStyle` reports the resolved values of the
  first run that has text (run → paragraph level of the shape's
  `a:lstStyle` → layout placeholder → master placeholder → master
  `p:txStyles` → the theme). What is written is always explicit on the slide
  shape; inherited values are never copied into the file unless an operation
  sets them.
- Elements that the renderer composes from the layout or the master
  (decorations, footers) are not listed: only the slide's own `p:spTree` is
  editable, as the ticket's "stable ids" require.

### Saving

- `materialize("show")` and `materialize("save")` are the package's
  `save()`: untouched entries verbatim, changed entries stored (module 03,
  decision 2). `PptxSaveOptions` has no fields of its own in this module;
  `compression: "deflate"` of the layer is not exposed yet.
- A saved file differs from the original only in the slide parts, `.rels`
  parts, `[Content_Types].xml` and `presentation.xml` that the operations
  touched, plus added media; the browser test compares entry by entry.
- `vbaProject.bin` of a `.pptm` is copied byte for byte and never read; the
  viewer's existing VBA warning is still reported on open.

### Viewer refresh

- After each committed call the core materializes the shown form and the
  viewer reopens it through `OfficeDocumentAdapter.reopen`: a full
  `PptxPresentation.load` of the edited bytes on the main thread, as viewing
  does. The cost is the renderer's parse and preflight of every slide; the
  spike measures it on 10, 100 and 500 slides and the browser suite fails
  above three seconds per `apply()`, as for PDF. Batching guidance goes in
  the docs.
- `changedPages` drives the viewport's repaint; thumbnails and the text map
  of the changed slides are invalidated by the core's existing path.

### Limits

- `maxEditOperations`, `maxEditHistory`, `maxEditCheckpointBytes` and
  `maxOperationMs` apply as for PDF; the package layer's `maxZipEntryBytes`
  and `maxExpandedOfficeBytes` bound inflation. A text of more than 100,000
  code units or a table of more than 100 × 20 cells is `range`-invalid.

## Work by layer

### Feature

Proposed tasks for `tasks/plan.md` Phase 8, each a commit with tests:

- **T44 Spike, engine skeleton, inspection.** `PptxEditEngine` over
  `OoxmlPackage` with open, `restore`, `materialize`, schemas and an empty
  operation set; the slide index (`presentation.xml`, `p:sldIdLst`, slide
  and layout part resolution), element listing with frames, inheritance and
  group transforms, `elementsAt`, `findText`, `getSlides`, `getLayouts`;
  the shared OOXML edit worker, the generalized worker client, the adapter's
  `edit` provider and `reopen`, the `pptm`/`ppsx` mapping. Spike first:
  `PptxPresentation.load` time for synthetic 10-, 100- and 500-slide decks
  with and without `progressiveLayout`; the renderer's bounds
  (`getElementBoundsByIds`) against the engine's for every shape of the
  corpus (the browser test keeps this check).
- **T45 Text.** `replaceText` (whole and ranged), `setTextStyle`, the text
  model, autofit handling, colours and fonts.
- **T46 Shapes.** `setShapeStyle`, `moveElement`, `resizeElement`,
  `deleteElement` (groups, placeholders, exclusive relationships),
  `insertTextBox`.
- **T47 Images and tables.** `insertImage`, `insertTable`, `setTableCell`.
- **T48 Slides.** `insertSlide`, `duplicateSlide`, `deleteSlide`,
  `moveSlide`.
- **T49 Browser round trip, latency, fixtures, docs, gate.** Every method
  through `viewer.edit()` in Chromium with a render check and a save-reload
  check; `apply()` latency on 10, 100 and 500 slides; a script that writes
  the edited fixture set for the manual PowerPoint and Keynote check;
  `docs/api/editing.md` PPTX section; roadmap and architecture; `npm run
check` and the matrix; Linear proofs; ACTION-812 Done except its release
  criterion.

### Tests

- **Unit (`node:test`):** `PptxEditEngine` on `sample.pptx`,
  `chart-point-colors.pptx` and decks from a new `pptx-builder.ts` fixture
  (a valid minimal deck: presentation, master, layout, theme, slides with
  placeholders, groups, tables, pictures, a hidden shape, a shape inside
  `mc:AlternateContent`, a `normAutofit` body, duplicate `cNvPr` ids, a
  `.pptm` with `vbaProject.bin`): every operation's XML result read back
  through the scanner; every element kind's bounds; inheritance cases; the
  text model's round trip; every issue code; the no-change identity and the
  replay identity (apply, undo, redo gives the same bytes); the dry run
  leaves bytes unchanged; batches with same-batch references
  (`insertSlide` then `insertTextBox` on `"$0"`).
- **Browser (Playwright):** `tests/e2e/edit-pptx.spec.ts` in the matrix:
  the worker is fetched only on `edit()`; every method on `sample.pptx`
  renders a visible change inside the element's bounds and survives
  `save()` → `viewer.load()`; engine bounds match the renderer's
  `getElementBoundsByIds` for every shape; a saved package differs from the
  original only in the expected entries; the latency test on 10, 100 and 500
  slides.
- **Fuzz:** no new target; the package layer's targets cover the inputs.
- **Manual (Leonid):** `npm run fixtures:pptx` writes
  `artifacts/pptx-fixtures/<operation>.pptx` for every operation on
  `sample.pptx`; they are opened in PowerPoint and Keynote and the result
  recorded in this spec's Actual result.

### Docs

- `docs/api/editing.md`: a `## PPTX` section between `## PDF` and the AI
  guidance: methods, elements and ids, slide space, text model, colours,
  slides, what stays unchanged, issue codes, performance.
- This spec's Spike results and Actual result; the roadmap's capability map,
  status and decisions log; `docs/architecture.md` (the PPTX engine next to
  the package layer); `THIRD_PARTY_NOTICES.md` (GenOffice entry extended
  with the text-edit ideas); `tasks/plan.md` and `tasks/todo.md` Phase 8.

## Definition of done

- Every operation of the table works through `apply()` and its typed
  method, is visible in the viewer after the call, and survives `save()`
  and reload, with a unit test and a browser test each.
- `save()` without changes returns identical bytes; undo to revision 0
  returns identical bytes; a saved file differs from the original only in
  the touched entries.
- `getElements` returns every slide-level element of the corpus with bounds
  that match the renderer's within one CSS pixel.
- `apply()` latency on 10-, 100- and 500-slide decks is recorded in
  `docs/api/editing.md` and this spec; every measured call stays under three
  seconds.
- The edited fixture set opens in PowerPoint and Keynote without a repair
  prompt (manual, recorded here).
- `npm run typecheck`, the unit suite, `npm run fuzz:js`, `npm run check`
  and the browser matrix pass; the size report lists the one new worker
  asset and nothing else.

## Decisions

Taken in the draft of 2026-10-02; open for review.

1. **Own XML inspection, no second renderer instance.** Bounds, text and
   styles come from the slide XML with placeholder inheritance and group
   transforms computed by the engine. The renderer's `getElementBoundsByIds`
   is a test oracle, not a runtime dependency: a session must answer reads
   in the worker without the main thread's presentation, and the renderer
   drops hidden shapes and flattens groups.
2. **One OOXML edit worker for PPTX and DOCX.** The package layer is
   thread-agnostic (03, decision 8) and inflating a deck must not block
   painting; one worker script keeps the asset budget to a single new entry
   and lets module 06 reuse the transport.
3. **Ids are `<slideKey>:<cNvPrId>`**, stable across reordering and across
   sessions; new ids are the next free ones, as PowerPoint allocates them,
   so a replay reproduces them and the file looks as if PowerPoint wrote it.
   The reuse caveat after deleting a created slide is documented rather than
   avoided with synthetic names (`slide1003.xml`-style numbering would be
   visible to anyone who inspects the package).
4. **Text model with `\n` and `\v`.** Paragraph and line breaks are
   distinct in PowerPoint and must stay distinct in replacements; `\v` is the
   character PowerPoint itself uses for `a:br` in its object model and
   clipboard.
5. **Minimal text patches keep untouched runs' bytes.** Whole-body
   replacement keeps the first paragraph's and run's properties; ranged
   replacement edits runs in place. Rebuilding every paragraph on each edit
   (GenOffice's fallback) is avoided because provenance flags would then be
   needed to stop inherited values from being written.
6. **Colours keep their theme form.** Writing `{ theme, mods }` produces
   `a:schemeClr`; reading reports it, so a host round-trips the theme link
   instead of flattening it (the architecture review's correction).
7. **`insertSlide` instantiates layout placeholders** as PowerPoint's New
   Slide does, not an empty `p:spTree`: the host gets editable title and
   body shapes at once and the file matches what PowerPoint would produce.
8. **`duplicateSlide` clones owned parts and shares media.** Charts,
   diagrams and embeddings cannot be referenced by two slides without
   PowerPoint repairing the file; images can.
9. **Stale autofit scales are dropped, not recomputed.** Re-deriving
   PowerPoint's font-scale ladder needs text measurement the engine does
   not have; an unscaled `a:normAutofit` is valid and PowerPoint recomputes
   it on its next layout. The renderer then shows the text unscaled, which
   can overflow the box until PowerPoint resaves; the docs say so.
10. **Full reopen per `apply()`, progressive for edited bytes.** The
    renderer has no update API (03's reading); `reopen` is a fresh load with
    `progressiveLayout` on, so the slide on screen paints without waiting
    for the deck's preflight (Firefox needs seconds for 500 slides); the
    three-second ceiling guards it in every browser of the matrix.
11. **Orphaned parts stay.** Deleting an element removes its relationships
    only; a reference count across every slide's relationships would be
    needed to remove parts safely, and PowerPoint discards orphans on save.
12. **`pptm` and `ppsx` on the same provider**, with the session's format
    `pptx`: the engine does not care about the content type of the main
    part, and the package layer already guarantees `vbaProject.bin`.
13. **Overlay primitives deferred** to a follow-up ticket, as for PDF.

## Open questions

None blocking. Items to confirm at review: decision 3's reuse caveat,
decision 9's visible overflow until PowerPoint resaves, and whether
`compression: "deflate"` should be exposed through `PptxSaveOptions` now.

## Spike results

T44, 2026-10-02, headless Chromium on an Apple M4 Pro, decks from
`test/fixtures/pptx-builder.ts` (one inherited title and one text box per
slide), `PptxPresentation.load` in main mode, then the first slide painted
at 960 px, then `waitUntilLayoutComplete`:

| Slides | Full load | Full first paint | Progressive load | Progressive first paint | Progressive complete |
| -----: | --------: | ---------------: | ---------------: | ----------------------: | -------------------: |
|     10 |     59 ms |            69 ms |            49 ms |                   50 ms |                57 ms |
|    100 |    120 ms |           122 ms |           209 ms |                  210 ms |               262 ms |
|    500 |    214 ms |           215 ms |           142 ms |                  143 ms |               284 ms |

In Chromium a full reopen of a 500-slide deck costs about a fifth of a
second. The browser matrix then showed Firefox taking about four seconds
for the same full load (its preflight of every slide is far slower), while
its progressive load resolves in 0.18 s and lays the rest out in the
background. The adapter therefore reopens edited bytes with
`progressiveLayout: true` (originals open as before): `apply()` resolves
when the shown slide can paint, and the three-second ceiling holds in
every browser of the matrix (decision 10, amended at T49).

Engine bounds against the renderer's `getElementBoundsByIds`: every shape
of `sample.pptx` (two inherited placeholders per slide, through the "Title
Slide" and "Title and Content" layouts down to the master body frame) and
of the built decks matches within one CSS pixel with the same rotation,
and the renderer reports every one of them with origin `slide`.

## Actual result

Implemented 2026-10-02 in `packages/viewer/src/edit/pptx/` (`types.ts`,
`schemas.ts`, `model.ts`, `geometry.ts`, `text.ts`, `style.ts`,
`elements.ts`, `engine.ts`, `operations.ts`, `handlers.ts`, `text-write.ts`,
`text-ops.ts`, `shape-ops.ts`, `image-table-ops.ts`, `slide-ops.ts`,
`handler.ts`, `provider.ts`, `session.ts`), `src/ooxml-edit-worker.ts`,
`src/edit/worker-engine.ts` (the worker transport shared with the PDF
client), the Office adapter's `edit` provider and the viewer's variant
mapping; tasks 44–49 of `tasks/todo.md`.

- **Every operation of the table ships** with its typed method, unit tests
  with byte-level expectations and a browser round trip on the corpus deck;
  the renderer's `getElementBoundsByIds` on the saved bytes agrees with the
  engine's frames within one CSS pixel after every geometry change, and the
  extracted text, the pixel checks and the reloads confirm what the
  renderer draws.
- **Deviations from the draft**, all in the API section and the docs:
  `PptxElement.hidden` reports `p:cNvPr/@hidden`; new `cNvPr` ids are
  allocated from the slide's own ids (ECMA-376 scopes them per part, and
  the renderer already prefers the slide on a clash with the layout);
  `insertTable` writes the id that `ppt/tableStyles.xml` names in `@def`
  (PowerPoint writes it even when the list defines nothing);
  `insertSlide` and `duplicateSlide` return the new slide's element ids in
  `createdIds`, so a batch can fill what it created; `insertImage` reuses
  identical bytes already in the package, not only within one batch;
  validation tracks the slide count through a batch; `moveElement` reports
  `required`/`conflict` for `to` and `by`.
- **Tests**: `pptx-edit-inspect` (6), `pptx-edit-text` (6),
  `pptx-edit-shapes` (5), `pptx-edit-tables` (4), `pptx-edit-slides` (5);
  `tests/e2e/edit-pptx.spec.ts` (8, in the matrix): worker fetched only on
  `edit()`, geometry oracle on the corpus and built decks, text, shapes,
  pictures and tables, slides, latency, the renderer load spike; the text
  test also compares the saved package entry by entry with the original
  (only `ppt/slides/slide1.xml` differs). `npm run fixtures:pptx` writes
  sixteen edited decks (one per operation plus "everything") to
  `artifacts/pptx-fixtures/` for the manual PowerPoint and Keynote check.
- **Latency** (`apply()` in the headless browser matrix, Apple M4 Pro,
  decks from the builder with one placeholder and one text box per slide;
  each call saves the package and reopens it progressively in the
  renderer; first then second `replaceText`, then an `insertTextBox` on the
  last slide):

  | Slides | Chromium             | Firefox                | WebKit               |
  | -----: | -------------------- | ---------------------- | -------------------- |
  |     10 | 32 ms, 204 ms, 80 ms | 79 ms, 198 ms, 53 ms   | 52 ms, 51 ms, 50 ms  |
  |    100 | 123 ms, 22 ms, 15 ms | 207 ms, 179 ms, 166 ms | 42 ms, 39 ms, 42 ms  |
  |    500 | 35 ms, 40 ms, 60 ms  | 93 ms, 144 ms, 177 ms  | 161 ms, 57 ms, 55 ms |

  Every call stays far inside the three-second ceiling. Before the
  progressive reopen, Firefox needed 3.9 s for an `apply()` on 500 slides
  because its full preflight of the deck takes four seconds; the decision 10
  amendment removed that.

- **Manual check (pending, Leonid)**: the sixteen fixtures in
  `artifacts/pptx-fixtures/` are to be opened in PowerPoint and Keynote; the
  result goes here. The ticket's release criterion is likewise Leonid's.
- Gates: unit suite, `fuzz:js`, the browser matrix and `npm run check`
  green at the T49 commit; the size report lists the one new asset,
  `workers/ooxml-edit-worker.js` (175 KB, 41 KB gzip), fetched only on the
  first `edit()` of a deck.
