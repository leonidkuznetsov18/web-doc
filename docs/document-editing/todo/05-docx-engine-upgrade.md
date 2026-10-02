# Module 05. `docx-engine-upgrade` — one `@silurus/ooxml` engine, on the 0.88 line

**Status:** Draft 2026-10-02 with spike results (Linear ACTION-813).
**Waiting for approval:** the ticket and the roadmap make the change of the
pinned `@silurus/ooxml` version an "ask first" item; this spec records what
the spike found so the decision can be taken on data. Nothing in
`package.json` changes before Leonid approves.

## Goal

Render DOCX (and XLSX, which shares the package) with the same
`@silurus/ooxml` 0.88 engine that PPTX already uses, so the viewer ships one
engine copy; keep the oversized-inline-image fitting that the old engine
needed a model patch for, as an XML pre-pass on the package layer; and give
`docx-edit` (module 06) a stable bridge from every rendered run to its `w:p`
in the source XML, with `w14:paraId` where the file has it and a generated
id where it does not.

## Requirement sources

- [`../00-roadmap.md`](../00-roadmap.md): fixed decisions 4 (rendering stays
  with the existing engines) and 6 (XML-first editing); the boundaries ("ask
  first: changing a pinned dependency"); the capability map's `05` row.
- [`./03-ooxml-package.md`](./03-ooxml-package.md): the package, the scanner
  and the patches the pre-pass uses.
- Linear ACTION-813: bump DOCX to the 0.88 line and drop the second engine
  copy and the npm alias; re-implement the image-fitting workaround as an
  XML pre-pass; expose `w14:paraId` (or a stable fallback id) on rendered
  runs and paragraphs; the DOCX visual and regression suites stay green (no
  SSIM regression beyond the existing thresholds) and the public corpus
  renders without new failures; the bundle does not grow; the license gate
  passes; a release is published and the monorepo bump verified on DOCX
  previews (Leonid's).
- Research of 2026-10-01 (`@silurus/ooxml` internals): 0.73.0 replaced the
  lazy DOCX pagination with an eager, immutable layout built inside
  `load()`, which broke web-doc's trick of adjusting the parsed model before
  the first layout; 0.73.0 also added `paragraphId` (`w14:paraId`),
  `source` (`{ story, storyInstance, path }`) and `sourceRunIndex` to text
  runs; the public `document` getter survives in main mode but no longer
  feeds layout.

## Spike results (T50, 2026-10-02)

`tests/e2e/docx-engine-spike.spec.ts` renders each fixture twice in
headless Chromium with deterministic fonts: through the viewer (0.72.2) and
through the 0.88 DOCX engine that the example build already vendors as the
PPTX engine (`/vendor/ooxml-pptx/docx.mjs`), then compares the two
renderings with the fidelity gate's SSIM and inspects the 0.88 run data.

| Fixture                       | Pages | SSIM 0.72 vs 0.88 | Load 0.72 → 0.88 | Runs | `paragraphId` | `source` | `document` getter |
| ----------------------------- | ----: | ----------------: | ---------------: | ---: | ------------: | -------: | ----------------: |
| `sample.docx` (corpus, POI)   |     1 |         **1.000** |      35 → 119 ms |  228 |             0 |      228 |               yes |
| `oversized-inline-image.docx` |     1 |             0.365 |       38 → 72 ms |    0 |             0 |        0 |               yes |

Readings:

1. **Fidelity holds on the corpus document**: the two engines paint
   `sample.docx` identically (SSIM 1.000 at 816 × 1056), so the fidelity
   snapshot and its 0.94 threshold would not move.
2. **The image fitting is gone without a pre-pass**: 0.88 draws the 21-inch
   inline picture at its declared extent (SSIM 0.365 against the fitted
   rendering), confirming that adjusting the model no longer reaches
   layout. The fitting must happen in the XML before the engine sees it.
3. **`w14:paraId` cannot be relied on**: the corpus document (written by
   Apache POI) has none, so 0 of 228 runs carry `paragraphId`. Every run
   does carry `source` (`{ story: "body", storyInstance: "body", path: [n] }`)
   and `sourceRunIndex`, but `path` indexes the engine's model, which
   unwraps `w:sdt`, hoists breaks and splits runs, so it is not an XML
   path. A generated id is needed.
4. **Load cost**: 0.88 takes about three times longer to open the one-page
   document (119 ms against 35 ms) because it lays the whole document out
   inside `load()`; the viewer's own load path hides most of it behind the
   first paint.
5. **XLSX moves with it**: the package is one; the spike did not render a
   workbook with 0.88 yet (T51 adds the compat workbook fixtures to the
   spike before the bump).

## In scope

- **The bump** (after approval): `@silurus/ooxml` → 0.88.0 for DOCX and
  XLSX; the `@silurus/ooxml-pptx` alias removed and every import of
  `@silurus/ooxml-pptx/pptx` turned back into `@silurus/ooxml/pptx`;
  `scripts/example.mjs` and `scripts/size-report.mjs` lose their second
  engine directory; `docs/architecture.md` and the notices follow.
- **The display pre-pass** (`src/adapters/docx-prepass.ts`), run on the
  bytes before `DocxDocument.load` and never on what the editor saves:
  - _Image fitting_: for every `wp:inline` whose `wp:extent` is wider than
    its section's content box (`w:pgSz` minus `w:pgMar`, the section being
    the `w:sectPr` that closes it), scale `wp:extent` and the picture's
    `a:ext` down with the aspect ratio kept, exactly as
    `fitInlineImagesToPage` does on the model today; anchored pictures keep
    their geometry.
  - _Paragraph ids_: every `w:p` of `word/document.xml` (and of the header,
    footer, footnote and endnote parts) without a `w14:paraId` gets one,
    generated from its position in document order (eight hex digits below
    `0x80000000`, as Word writes them, from a fixed base so the same file
    always gets the same ids), with the `w14` namespace declared on the
    root and listed in `mc:Ignorable` when missing. The engine then reports
    `paragraphId` on every run, and `docx-edit` computes the same ids from
    the original bytes with the same walk, so a rendered run maps to its
    `w:p` by id alone.
  - Both are patches through the package layer (module 03), so untouched
    bytes stay and malformed parts are refused the way they are today.
- **The run bridge**: the Office adapter keeps `paragraphId` on the
  `TextRun` values it exposes (a new optional `paragraphId` field on
  `TextRun`, documented), so hosts and module 06 can map selections to
  paragraphs without re-reading the engine.
- **Tests**: the spike spec becomes a regression (0.88 against the stored
  renderings, SSIM threshold as the fidelity gate's), the inline-image spec
  passes on the pre-pass, a unit test for the pre-pass on hand-built
  documents (sections, tables, headers, `w:sdt`, existing ids kept), the
  corpus renders (regression harness), the compat workbook fixtures through
  0.88.

## Out of scope

- Any editing of DOCX (module 06).
- Writing the generated `w14:paraId` values into saved files (module 06
  decides whether an edit that touches a paragraph also persists its id, as
  Word does on save).
- Changing the renderer's layout mode (`main` stays; `worker` mode is a
  separate decision).

## Work by layer

Proposed tasks for `tasks/plan.md` Phase 9, each a commit with tests:

- **T50 Spike.** `tests/e2e/docx-engine-spike.spec.ts` (done above); this
  spec with the results; the approval request.
- **T51 Pre-pass.** `docx-prepass.ts` on the package layer: image fitting
  and paragraph ids, unit tests, the inline-image spec against the pre-pass
  on the vendored 0.88 engine (still behind the alias), workbook fixtures
  added to the spike.
- **T52 Bump** (after approval). `package.json`, imports, scripts, docs,
  notices; fidelity snapshots re-checked; regression harness; size report
  (one engine copy); license gate.
- **T53 Run bridge and docs.** `TextRun.paragraphId`, `docs/api` updates,
  the roadmap, `npm run check`, Linear proofs.

## Definition of done

- `npm run test:e2e:matrix` and `npm run check` green with one
  `@silurus/ooxml` copy at 0.88.0; the fidelity gate's SSIM for the DOCX
  family at or above its threshold; the regression harness reports no new
  DOCX or XLSX failure.
- `oversized-inline-image.docx` still renders its picture inside the content
  box (the existing spec), now through the pre-pass.
- Every run the adapter reports for a DOCX page carries a `paragraphId`
  that names a `w:p` of the source XML, whether or not the file has
  `w14:paraId`.
- The size report shows the engine once; the license gate passes.
- The release and the monorepo verification are Leonid's.

## Decisions

Proposed in the draft; the bump itself awaits approval.

1. **Pre-pass over model patching.** 0.88 lays out inside `load()` from a
   sealed clone; the only supported way to change what it draws is the
   bytes it reads. The package layer already patches exact ranges, so the
   pre-pass costs a scan of `word/document.xml` (5 ms on the corpus) and
   nothing of the file's fidelity.
2. **Generated paragraph ids in the display copy, never in the saved
   file.** Word-written files keep their own `w14:paraId`; files without
   them get deterministic ids that `docx-edit` recomputes from the original
   bytes, so the bridge works for every producer and the saved file does
   not change for a read.
3. **XLSX moves with DOCX.** One package, one version; the spike renders the
   workbook fixtures before the bump so an XLSX regression is seen first.

## Open questions

- **Approval of the bump** (`@silurus/ooxml` 0.72.2 → 0.88.0, DOCX and
  XLSX): the data above is the case for it. T51 proceeds behind the alias
  meanwhile; T52 waits.
