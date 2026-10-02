# Module 06. `docx-edit` — DOCX editing on the package layer

**Status:** In progress 2026-10-02 (Linear ACTION-814): T54 inspection
landed; T55–T58 follow. Written after module 05's spike so it builds on
what the 0.88 engine really exposes; module 05 landed the same day.

## Goal

Give hosts and AI clients the DOCX method set of the roadmap through the
same `EditSession` contract as PDF and PPTX: inspect paragraphs, tables and
inline pictures with stable ids and page geometry, replace and restyle
paragraph text, insert, delete and move paragraphs, insert tables and edit
their cells, insert inline pictures. Every edit is a patch to the story
part through `ooxml-package` (module 03); untouched paragraphs keep their
bytes; the viewer reopens the edited bytes through the display pre-pass of
module 05 with the same `@silurus/ooxml` engine.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md): fixed decisions 1 (API only;
  overlay editing with commit on blur or idle), 2, 3, 4, 6 (XML-first), 7
  (`docm` follows; macros copied and never executed); the MVP method set's
  DOCX column; the hard limits (no live reflow: `apply()` is a commit path).
- [`./01-edit-core.md`](./01-edit-core.md): `EditElement` with
  `fragments` (an element that spans pages) and `story`, `TextRange`
  offsets in UTF-16 code units with tabs, breaks, inline images and fields
  as one placeholder character each, `DocumentChange.changedPages` "for
  flow formats every page from the first affected one".
- [`./03-ooxml-package.md`](./03-ooxml-package.md): patches, transactions,
  `addMedia`, snapshots.
- [`./05-docx-engine-upgrade.md`](./05-docx-engine-upgrade.md): the
  display pre-pass, the `_wd<id>` bookmark bridge, `TextRun.paragraphId`,
  and the spike's reading of 0.88 (`source.path` into the engine's model,
  `w14:paraId` not read by the engine, no glyph geometry).
- Linear ACTION-814: text (replace, insert, delete; runs, bookmarks, comment
  ranges and fields the edit does not touch preserved); run formatting,
  paragraph alignment and spacing, text colour and highlight; insert,
  delete and move paragraphs; insert a table and edit cell text; insert an
  inline image; inspection with stable ids, page bounds from the renderer,
  text and style; `.docm` on the same path; files open without a repair
  prompt in Word and Pages (manual fixture set); untouched parts and
  paragraphs byte-identical; latency on 10-, 100- and 500-page documents
  that sets the commit-on-idle guidance.
- ECMA-376 Part 1 (WordprocessingML): `w:body` holds paragraphs, tables and
  a trailing `w:sectPr`; a section break lives in the `w:pPr` of the last
  paragraph of its section; a cell ends with a paragraph; `w:t` needs
  `xml:space="preserve"` for leading or trailing spaces; `w:rPr` and
  `w:pPr` children are ordered (`rStyle, rFonts, b, …, color, …, sz, …,
highlight, u, …` and `pStyle, keepNext, …, numPr, …, spacing, ind, …, jc,
…, rPr, sectPr`); `w14:paraId` is unique within a part, below
  `0x80000000`; bookmark ids are unique and paired; `wp:docPr/@id` is
  unique in the document.
- Research of 2026-10-01 (ideas only, no code; GenOffice attributed in
  `THIRD_PARTY_NOTICES.md`, SuperDoc never copied):
  - GenOffice (Apache-2.0): a flat block index over the original bytes
    (`docxIndex`, `originalXml`, `rawPPr`, `rawRPr`), an edited block
    consuming its anchor so section properties and revisions are never
    duplicated, property merges that keep original bytes when the value is
    unchanged and interleave rebuilt children in schema order, field
    balancing, section properties kept unless page setup changed, a save
    fast path that returns the original bytes, and its open issues (regex
    surgery breaking on single quotes and attribute order, multi-byte
    splice offsets) that the package layer's scanner avoids by design.
  - SuperDoc (AGPL): "a DOCX is a package of related XML parts, not one
    editor tree"; preservation is node-local (keep the original element
    attached to every edited node); run-structured writes with
    non-inherited properties only.
  - The `@silurus/ooxml` reading: 0.72.2 and 0.88 models drop `w:sdt`
    wrappers, field codes and hidden runs, so no model can be the source of
    a faithful save (module 03, decision 6).

## Dependencies

- Modules 01, 03 and 05 done; the OOXML edit worker of module 04 (the
  DOCX engine is its second format).
- No new runtime dependency.

## In scope

### Inspection

`getElements`, `getElement`, `elementsAt` and `findText` from `edit-core`,
returning `DocxElement` values for the **body story**; headers, footers,
footnotes and endnotes are listed with their `story` and `operations: []`
(read-only in this module).

```ts
export type DocxElementKind =
  | "paragraph" // w:p of the body or a table cell
  | "table" // w:tbl
  | "image" // an inline picture (w:drawing/wp:inline) inside a paragraph
  | "other"; // anchored drawings, OLE objects, text boxes, equations: listed, not edited

export interface DocxElement extends EditElement {
  readonly kind: DocxElementKind;
  /** Present for a paragraph: resolved style of its first run with text. */
  readonly textStyle?: DocxTextStyle;
  /** Present for a paragraph. */
  readonly paragraphStyle?: DocxParagraphStyle;
  /** Present for a table. */
  readonly table?: { readonly rows: readonly (readonly string[])[] };
  /** Present for a paragraph whose text cannot be edited (tracked changes, a section break, a field-only paragraph). */
  readonly readOnlyReason?:
    "tracked-changes" | "section-break" | "unsupported-content";
}

export interface DocxTextStyle {
  readonly fontFamily: string; // theme fonts resolved
  readonly fontSize: number; // points
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  readonly color: EditColor; // "#RRGGBB", { theme } for a theme colour, "auto"
  readonly highlight?: DocxHighlight; // one of Word's sixteen highlight names
}

export interface DocxParagraphStyle {
  readonly styleId?: string; // w:pStyle
  readonly align: "left" | "center" | "right" | "justify";
  /** Points; absent values inherit from the style. */
  readonly spacing: {
    readonly before?: number;
    readonly after?: number;
    readonly line?: number;
  };
  readonly numbering?: { readonly numId: number; readonly level: number };
}
```

- **Ids.** A paragraph's id is `p:<id>` where `<id>` is its `w14:paraId`
  when the file has one, else the id module 05 generates from its position
  among the unmarked paragraphs of its part (the same walk, so the display
  pre-pass's `_wd<id>` bookmark and the engine agree). A table is
  `tbl:<id of its first paragraph>`; an inline picture `img:<paragraph
id>.<n>` (the n-th drawing of the paragraph, 0-based). Ids are stable for
  the session: a paragraph the session rebuilds or inserts is written with
  its `w14:paraId`, and the shown copy carries an id for every paragraph
  (see Saving), so deleting an unmarked paragraph does not renumber the
  rest. Across sessions, ids of paragraphs that a previous session never
  touched can differ when paragraphs were inserted or deleted in between;
  Word-written files, which carry `w14:paraId` on every paragraph, are
  stable across sessions.
- **Geometry from the renderer.** The engine has no layout: `bounds`,
  `fragments` and `rotation` come from the viewer's text map, where every
  run carries `paragraphId` (module 05, T53). The `DocxSession` on the main
  thread joins the engine's elements with the runs of the pages the query
  names: a paragraph's fragments are the unions of its runs' boxes per
  page; a table's are the unions of its cells' paragraphs; an inline
  picture's box comes from the picture run. A query without `pageIndex`
  lists every body element, with `bounds` for the pages already painted
  and an empty `fragments` list for the rest; `elementsAt` and `findText`
  rectangles are run boxes. `pageIndex` of an element is that of its first
  fragment.
- **Text model.** `text` of a paragraph: `w:t` text, a tab `\t`, a line
  break `\v`, a page or column break `\f`, an inline picture `￼`, a
  field (`w:fldSimple`, or a complex field from `fldChar begin` to `end`)
  as its cached result text, hyperlinks and inline `w:sdt` as their runs;
  hidden runs (`w:vanish`) count as text. A table's `text` joins cells by
  tab and rows by newline.

### Operations

`DocxEditSession` exposes one typed method per operation. Positions in a
flow document are other elements, not page points.

| Operation           | Fields                                                                                                                             | Category            |
| ------------------- | ---------------------------------------------------------------------------------------------------------------------------------- | ------------------- |
| `replaceText`       | `target` (a paragraph), `text`, `range?: TextRange` on the target (collapsed inserts)                                              | text                |
| `setTextStyle`      | `target` (a paragraph), `range?`, `style: DocxTextStyleChange`                                                                     | formatting, colours |
| `setParagraphStyle` | `target` (a paragraph), `style: { align?, spacing?: { before?, after?, line? } }`                                                  | alignment, spacing  |
| `insertParagraph`   | exactly one of `before` / `after` (a paragraph or table id), `text`, `style?: DocxTextStyleChange`                                 | insertion           |
| `deleteElement`     | `target` (a paragraph, table or inline picture)                                                                                    | structure           |
| `moveElement`       | `target` (a paragraph or table), exactly one of `before` / `after` (a paragraph or table id of the same story)                     | structure           |
| `insertTable`       | exactly one of `before` / `after`, `rows: string[][]`, `columnWidths?: number[]`                                                   | tables              |
| `setTableCell`      | `target` (a table), `row`, `column`, `text`                                                                                        | tables              |
| `insertImage`       | exactly one of `before` / `after`, `data: BinaryData`, `mimeType: "image/png" \| "image/jpeg"`, `size: { width, height }` (points) | insertion           |

```ts
export interface DocxTextStyleChange {
  readonly fontFamily?: string; // w:rFonts ascii and hAnsi
  readonly fontSize?: number; // 1–400 pt → w:sz in half-points
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean; // w:u single / none
  readonly color?: EditColor; // "#RRGGBB" → w:color; { theme } → w:color with themeColor and the theme's value
  readonly highlight?: DocxHighlight | "none"; // w:highlight
}

export type DocxHighlight =
  | "yellow"
  | "green"
  | "cyan"
  | "magenta"
  | "blue"
  | "red"
  | "darkBlue"
  | "darkCyan"
  | "darkGreen"
  | "darkMagenta"
  | "darkRed"
  | "darkYellow"
  | "darkGray"
  | "lightGray"
  | "black"
  | "white";
```

Which operations an element accepts is listed in its `operations` field:

| Kind        | Operations                                                                                                                          |
| ----------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `paragraph` | `replaceText`, `setTextStyle`, `setParagraphStyle`, `insertParagraph`, `insertTable`, `insertImage`, `moveElement`, `deleteElement` |
| `table`     | `setTableCell`, `insertParagraph`, `insertTable`, `insertImage`, `moveElement`, `deleteElement`                                     |
| `image`     | `deleteElement`                                                                                                                     |
| `other`     | none                                                                                                                                |

A paragraph with `readOnlyReason` accepts only the insertion operations
that place a sibling next to it.

## Out of scope

- Headers, footers, footnotes, endnotes, comments, tracked changes
  (authoring or accepting; a paragraph inside `w:ins`/`w:del`/`w:moveFrom`/
  `w:moveTo` is read-only), fields and hyperlinks (creation), lists
  (creating numbering; an existing list paragraph keeps its `w:numPr`),
  styles and the styles part, sections and page setup, floating pictures,
  text boxes, equations, nested-table structure (rows and columns), cell
  formatting, merges.
- Overlay text-input primitives (caret, range rectangles, suppressed
  render): a follow-up ticket, as for PDF and PPTX.
- Word-like live reflow: `apply()` reopens the document (hard limit).

## Behaviour

### Engine and loading

- The DOCX engine runs in the OOXML edit worker of module 04 (the open
  payload names the format); the engine class is thread-agnostic for the
  Node tests. It opens the **original** bytes with the package layer and
  builds the block index of the body story: every top-level `w:p` and
  `w:tbl` (and the paragraphs inside cells and inline `w:sdt`), with its
  id, its text model and the run items it is made of.
- After each committed call the viewer reopens the materialized bytes
  through the Office adapter, which runs the display pre-pass of module 05
  and loads the result; `changedPages` is every page from the first page
  of the first changed element onwards (flow reflow), as the core allows.

### Ids and the display copy

- The engine computes ids with module 05's walk over the original part:
  `w14:paraId` where present, else positional among unmarked paragraphs.
- `materialize("show")` writes a `w14:paraId` on **every** paragraph of
  the body part that lacks one (declaring `w14` and `mc:Ignorable` when
  needed), so the pre-pass's bookmarks and the viewer's `paragraphId` carry
  the engine's ids and never positional ones; `materialize("save")` writes
  ids only on the paragraphs the session rebuilt or inserted, so untouched
  paragraphs keep their bytes (ticket).
- Element ids are never reused within a session; a deleted paragraph's id
  is reported in `removedIds`.

### Text

- `replaceText` rebuilds only the paragraph it targets, from run items:
  untouched runs, hyperlinks, inline `w:sdt`, bookmarks, comment range
  markers, `w:proofErr` and other children keep their bytes in place;
  touched runs are re-serialized with their `w:rPr` bytes kept and the
  new text written as `w:t` (`xml:space="preserve"` when it starts or ends
  with a space), `w:tab` for `\t`, `w:br` for `\v` and `w:br w:type="page"`
  for `\f`. A whole replacement keeps the paragraph's `w:pPr` and the
  first text run's `w:rPr`; a ranged one styles new text like the run that
  held the first replaced character (the run before the caret for an
  insertion). A range that cuts a field, a hyperlink's edge, an inline
  picture or an inline `w:sdt` boundary is `invalid-range`; a range that
  covers them whole removes them. Newlines in `text` are paragraph breaks:
  the paragraph is split, the new paragraphs copy its `w:pPr` without
  `w:sectPr` and get fresh ids (in `createdIds`).
- `setTextStyle` splits runs at the range ends and merges the change into
  each covered run's `w:rPr` in schema order, keeping unknown children
  (GenOffice's merge idea: a property whose value does not change keeps
  its bytes); the paragraph mark's `w:pPr/w:rPr` takes the change when the
  range reaches the paragraph end, so new text inherits it. Theme colours
  write `w:color w:val` resolved from the theme part plus `w:themeColor`.
- `setParagraphStyle` writes `w:jc` and `w:spacing` (`w:before`,
  `w:after`, `w:line` with `w:lineRule="auto"`, twentieths of a point)
  into `w:pPr` in schema order, creating `w:pPr` when absent and keeping
  everything else, `w:sectPr` included.

### Structure

- `insertParagraph` writes a new `w:p` next to the reference element with
  the reference paragraph's `w:pPr` (without `w:sectPr`; a table reference
  gives a plain paragraph), a fresh `w14:paraId`, and runs styled like the
  reference's first run unless `style` says otherwise. Inserting after the
  last paragraph of the body keeps the body's `w:sectPr` last.
- `deleteElement` removes a paragraph, a table or an inline picture (the
  drawing run). The last paragraph of the body or of a cell, and a
  paragraph whose `w:pPr` holds a `w:sectPr`, are refused
  (`last-paragraph`, `section-break`); a table immediately followed by the
  body's `w:sectPr` is followed by an empty paragraph after deletion only
  if the body would otherwise end with a table. Relationships only the
  removed drawing used are removed; media parts stay (as in PPTX).
- `moveElement` cuts the element's bytes and inserts them next to the
  reference within the same container (body, or the same cell); a
  paragraph carrying a `w:sectPr` cannot move.

### Tables and pictures

- `insertTable` writes `w:tbl` with `w:tblPr` (`w:tblStyle` naming
  `TableGrid` when the styles part defines it, else `w:tblBorders` single
  lines, `w:tblW w:type="auto"`, `w:tblLook`), `w:tblGrid` from the column
  weights over the section's content width, rows of cells with `w:tcW` and
  one paragraph each (with a fresh id); a paragraph follows the table when
  the next sibling would otherwise be a table or the `w:sectPr`.
- `setTableCell` rebuilds the cell's first paragraph like a whole-text
  replacement and removes the cell's other paragraphs; the cell's `w:tcPr`
  stays. A row or column outside the table is a `range` issue.
- `insertImage` stores the bytes once (`addMedia`, SHA-256), relates the
  part from the document and writes a paragraph holding
  `w:r/w:drawing/wp:inline` with `wp:extent` from `size` (points × 12,700),
  a unique `wp:docPr/@id`, `pic:pic` with `a:blip r:embed`. The display
  pre-pass fits it to the section like any other inline picture.

### Saving

- `save()` is the package's `save()` with the id rule above; no changes
  returns the original bytes; a `.docm` keeps `vbaProject.bin` untouched.
- `DocxSaveOptions` has no fields of its own in this module.

### Latency and commit-on-idle

- Each `apply()` saves the package, runs the pre-pass and reopens the
  document, which lays out the whole document (upstream measures ≈0.05 s
  for 30 pages and ≈0.2 s for 108 pages; a 500-page document can take
  seconds). The browser suite records `apply()` on 10-, 100- and 500-page
  documents built by a fixture generator; 10 and 100 pages must stay
  inside the three-second ceiling, 500 pages is recorded and sets the
  docs' commit-on-idle guidance (debounce typing, commit on blur).

### Limits

- `maxEditOperations`, `maxEditHistory`, `maxEditCheckpointBytes`,
  `maxOperationMs`, and the package layer's inflation limits apply; a text
  over 100,000 code units or a table over 100 × 20 cells is a `range`
  issue.

## Work by layer

### Feature

Proposed tasks for `tasks/plan.md` Phase 10 (after Phase 9):

- **T54 Engine skeleton and inspection.** Block index, ids, text model,
  styles (run and paragraph, theme fonts and colours), the worker format
  switch, the Office adapter's provider for `docx`/`docm`, the session's
  geometry join with the viewer's runs, `elementsAt`, `findText`, no-change
  identity, restore; browser test against the corpus document and built
  fixtures (`docx-builder.ts`).
- **T55 Text and formatting.** `replaceText`, `setTextStyle`,
  `setParagraphStyle`, the `w:rPr`/`w:pPr` schema-order merge, fields,
  hyperlinks, inline `w:sdt`, paragraph splits, the id rule of
  `materialize`.
- **T56 Structure and pictures.** `insertParagraph`, `deleteElement`,
  `moveElement`, `insertImage`.
- **T57 Tables.** `insertTable`, `setTableCell`.
- **T58 Round trip, latency, fixtures, docs, gate.** Every method in the
  browser with render and reload checks and entry-by-entry comparison;
  latency on 10/100/500 pages; `npm run fixtures:docx` for the Word and
  Pages check; `docs/api/editing.md` DOCX section; roadmap; `npm run check`;
  Linear proofs.

### Tests

- Unit: the engine on `sample.docx` and on built documents covering every
  construct above (fields, hyperlinks, inline sdt, bookmarks, comment
  ranges, tracked changes, section breaks, tables, pictures, `docm`);
  byte-level expectations as in the PPTX suites; replay identity; every
  issue code.
- Browser: `tests/e2e/edit-docx.spec.ts` in the matrix; the geometry join
  checked against `getPageText` and the text map; the reflow of
  `changedPages`.
- Manual (Leonid): the fixture set in Word and Pages.

### Docs

- `docs/api/editing.md` `## DOCX` section; this spec's Actual result; the
  roadmap; `THIRD_PARTY_NOTICES.md` (GenOffice entry extended with the
  DOCX ideas).

## Definition of done

- Every operation works through `apply()` and its typed method, is visible
  after the call and survives save and reload, with unit and browser tests.
- No-change and undo-to-zero return identical bytes; a saved file differs
  from the original only in the touched entries, and inside the story part
  only in the paragraphs the session touched (and their new `w14:paraId`).
- Every body element of the corpus document is listed with ids that match
  the viewer's `paragraphId` runs and with bounds that cover its runs.
- Latency recorded; the fixture set handed over; `npm run check` and the
  matrix green.

## Decisions

Proposed in the draft; open for review.

1. **Ids from `w14:paraId` with module 05's positional fallback**, written
   into the shown copy for every paragraph and into the saved file only
   for touched paragraphs: session-stable ids, byte-identical untouched
   paragraphs.
2. **Geometry from the renderer, joined on the main thread.** The engine
   never lays out; the viewer's runs carry `paragraphId`, and the session
   joins them. This is what "page bounds from the renderer" in the ticket
   means, and it keeps the worker free of the engine.
3. **Paragraph-local rebuilds, run items with kept bytes** (GenOffice's
   index-over-bytes and SuperDoc's node-local preservation as ideas): the
   smallest patch is one paragraph; nothing outside it changes.
4. **Positions are elements, not points.** Insertion and moves name a
   sibling (`before`/`after`); page points belong to fixed-layout formats.
5. **Read-only where fidelity cannot be guaranteed:** tracked-change
   paragraphs, section-break paragraphs, and `other` content are listed
   and left alone rather than rewritten approximately.
6. **500 pages is recorded, not gated.** The full-document layout of the
   renderer is the cost; the docs set the commit-on-idle guidance from the
   measurement.

## Open questions

- Whether `moveElement` should also accept a page point for hosts that drag
  paragraphs on the canvas (the session could resolve the point through
  `elementsAt`); left out of the MVP.
- Whether `materialize("show")` should also write the pre-pass bookmarks
  itself to save one scan on each reopen; measured in T58.

## Actual result

- **T54 (2026-10-02)**: `src/edit/docx/` holds the engine skeleton.
  `ids.ts` is the id walk the display pre-pass now shares (`collectIds`,
  `assignParagraphIds`: `w14:paraId` or a generated id, one sequence over
  the main part and the story parts), so engine ids and the pre-pass
  bookmarks agree by construction. `model.ts` indexes the body: `w:p` and
  `w:tbl` in order, block-level `w:sdt`/`w:customXml` unwrapped, cell
  paragraphs under their table (a table nested in a cell is left out),
  `tbl:<first paragraph id>`, `img:<pid>.<n>` for inline pictures,
  `other:<pid>.<n>` for anchored drawings, objects, equations and
  alternate content. `text.ts` reads the paragraph text model (tabs,
  `\v`, `\f`, U+FFFC for pictures and objects, simple and complex fields as
  their cached result as one item that remembers its runs, hyperlinks,
  inline sdt, tracked insertions read and deletions dropped, symbols,
  no-break hyphens, note references as empty items), every item keeping
  its `w:r`, child, `w:rPr` and wrapper for the patches of T55. `style.ts`
  resolves run and paragraph styles through run properties, the character
  and paragraph style chains, the default paragraph style, the document
  defaults and the theme fonts; spacing attributes inherit one by one.
  `engine.ts` serves the `EditEngine` contract with a handler map that
  grows per task (`IMPLEMENTED_OPERATIONS` empty, so no element accepts an
  operation yet), identity without changes, snapshot rollback, restore,
  `findText` over paragraph text without geometry; `elementsAt` is the
  session's. The worker handler switches on the open payload's format
  (`docx` → `DocxEditEngine`, the PPTX reads refused on it); the Office
  adapter advertises `docx` and `docm` and picks the provider and session
  by format; `EditEngineProvider.createSession` now receives an
  `EditSessionAccess` (the viewer's cached text runs), which `DocxSession`
  uses for the geometry join: fragments per page from the runs'
  `paragraphId`, tables through their cell paragraphs, inline objects
  through their paragraph, page-scoped queries reading one page, page-less
  queries reading only cached pages, `elementsAt` through run boxes
  (paragraph then table), `findText` placing matches from the runs with
  pages located cache-first. Known limits recorded in the docs: a
  picture's box is its paragraph's (the renderer lists no picture
  geometry); since T56 a paragraph that draws no run (empty, or pictures
  only) is placed by estimate next to its placed neighbour, a picture
  paragraph sized from its largest picture's extent (`imageSize` on the
  picture element). Tests: `docx-edit-inspect.test.ts` (10: ids against the
  pre-pass, text model, read-only reasons, styles, search, identity,
  restore, broken package, schema conformance, corpus) and
  `docx-edit-session.test.ts` (5: join, cached pages, hit test, search
  rectangles, worker protocol); `tests/e2e/edit-docx.spec.ts` in the
  matrix (worker loads on `edit()` only, every corpus paragraph's bounds
  cover its runs, hit test, search, identity; a paragraph split by a page
  break joined across two pages with its table). Unit 384/384.
- **T55 (2026-10-02)**: `replaceText`, `setTextStyle` and
  `setParagraphStyle` in `src/edit/docx/text-ops.ts` over the writer in
  `write.ts`. A paragraph is rebuilt from units (its direct children, a
  complex field's runs as one): untouched units keep their bytes, a
  touched run is re-serialized around the change with its `w:rPr`, a
  hyperlink or content control is entered when the range lies inside it,
  bookmarks and comment markers stay in place, covered fields, links,
  controls and pictures go (pictures reported in `removedIds`), and the
  paragraph is written back as one element with `w14:paraId` (the root
  gains `xmlns:w14` and `mc:Ignorable` once). Newlines split the
  paragraph: fresh ids derived from the batch's state id and the
  operation's position (so a replay issues the same), properties copied
  without `w:sectPr`, the tail moved to the last paragraph. Run and
  paragraph properties merge in schema order (`w:rFonts` keeps its other
  faces and drops the theme overrides, `w:sz` and `w:szCs` together,
  theme colours written with the theme's resolved value, `w:spacing`
  keeps its other attributes); the paragraph mark takes a style change
  that reaches the end. Ids stay stable across edits: the engine tracks
  the ids of paragraphs without `w14:paraId` in document order and
  removes those it stamps or deletes, the model reads ids from that list,
  `materialize("show")` stamps every remaining paragraph in a copy (so
  the pre-pass and the viewer's runs name the engine's paragraphs and a
  checkpoint opened as a base keeps every id) while `materialize("save")`
  carries ids only where the session wrote. Reflow: an operation reports
  `reflowFrom` (the paragraph id), the core asks the host for that
  paragraph's first page before the document is replaced and repaints
  from it to the end, for apply, undo and redo alike (`EngineChange`,
  `EditSessionHost.pageOf`, `HistoryEntry.reflowFrom`). Validation:
  control characters and lone surrogates, unknown and read-only targets,
  offsets, surrogate splits, field cuts, wrapper edge cuts, paragraph
  breaks inside a wrapper, colours (`#RRGGBB`, `auto`, Word theme colour
  names). Tests: `docx-edit-text.test.ts` (10, byte-level) and the
  browser round trip in `edit-docx.spec.ts` (replace, restyle, align,
  split, `getPageText` after each, `documentchange` pages, reload of the
  saved bytes, undo to identical bytes, stable id of the untouched
  paragraph). Unit 394/394.
- **T56 (2026-10-02)**: `insertParagraph`, `deleteElement`, `moveElement`
  and `insertImage` in `src/edit/docx/structure-ops.ts`. Insertion is
  one `insertBefore`/`insertAfter` patch per paragraph next to the
  reference's node (a paragraph reference lends its `w:pPr` without
  `w:sectPr` and its first run's `w:rPr`, a table reference a plain
  paragraph; `style` merges in); the body's `w:sectPr` stays last by
  construction. Deletion removes the block (or the `w:drawing` of a
  picture, with the relationship only it used) and reports every element
  under it in `removedIds` and every `w:p` in `removedParagraphIds` so the
  engine's id list stays right; the last paragraph of a body or cell and
  a section-break paragraph are refused (`last-paragraph`,
  `section-break`); a body that would end with a table gains an empty
  paragraph. A move cuts the block's bytes and re-inserts them with a
  `w14:paraId` written on every paragraph in them (nested ones included),
  so ids survive the move and the unauthored list only shrinks. Pictures
  are a new paragraph with `wp:inline`, `wp:docPr` ids above any in the
  part, bytes stored once under `word/media` (identical bytes reuse the
  part), the `wp` and `r` namespaces declared on the root when missing.
  A table is named after its first paragraph, so an insertion before it,
  a move within the first cell or a deletion of it renames the table;
  the operation reports `remappedIds` and the engine chains them through
  a batch (`EngineChange.remappedIds`). Reflow starts at the reference or
  the earlier of a move's two elements. Tests:
  `docx-edit-structure.test.ts` (5) and the browser round trip in
  `edit-docx.spec.ts` (insert with style, move before, picture after,
  `getPageText` order, delete, reload, reset to identical bytes). Unit
  399/399.
