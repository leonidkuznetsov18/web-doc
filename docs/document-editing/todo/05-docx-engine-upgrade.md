# Module 05. `docx-engine-upgrade` — one `@silurus/ooxml` engine, on the 0.88 line

**Status:** Bump done 2026-10-02 (T52); the run bridge (T53) follows
(Linear ACTION-813).
**Approval:** the ticket and the roadmap make the change of the pinned
`@silurus/ooxml` version an "ask first" item; the spike results below were
the case for it, and Leonid approved the bump on 2026-10-02. One
`@silurus/ooxml` 0.88.0 serves every format since T52.

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

`tests/e2e/docx-engine-spike.spec.ts` (since T52 the regression
`docx-engine.spec.ts`) renders each fixture twice in
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
3. **`w14:paraId` is not read by 0.88.0 at all**: the corpus document
   (written by Apache POI) has none, and a probe with a Word-style
   `w14:paraId="1A00ABCD"` on a hand-built paragraph came back without
   `paragraphId` on the run and without `paragraphId` on the model
   paragraph; the `paraId` the bundle reads belongs to comments. Every run
   does carry `source` (`{ story: "body", storyInstance: "body", path }`)
   and `sourceRunIndex`, where `path` indexes the engine's model: `[2, 0,
0, 0]` for the first paragraph of the first cell of a table, `w:sdt`
   unwrapped, and a paragraph that holds a page break split into two model
   paragraphs around a hoisted `pageBreak`. The model keeps bookmark names
   on each paragraph (`bookmarks: ["_wd1A000000"]`), which is the bridge
   the pre-pass uses: with it, all 228 runs of `sample.docx` resolve to
   their paragraph on 0.88.
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
  - _Paragraph ids_: every `w:p` of `word/document.xml` and of the header,
    footer, footnote and endnote parts gets a hidden bookmark pair right
    after its `w:pPr` (`<w:bookmarkStart w:id="7000000" w:name="_wd<id>"/>`
    and its end), the id being the file's `w14:paraId` when it has one and
    otherwise generated from the paragraph's position in document order
    (eight hex digits below `0x80000000`, as Word writes them, from a fixed
    base so the same file always gets the same ids; bookmark ids start
    above any the part uses). The engine keeps bookmark names on its model
    paragraphs, so a run's `source.path` leads to a paragraph whose `_wd…`
    bookmark names the `w:p`; a paragraph split around a page break takes
    the id of the nearest earlier paragraph of its container. `docx-edit`
    computes the same ids from the original bytes with the same walk, so a
    rendered run maps to its `w:p` by id alone.
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

- **T50 Spike.** `tests/e2e/docx-engine-spike.spec.ts` (done above; renamed
  to `docx-engine.spec.ts` as the regression in T52); this
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
2. **Paragraph ids ride on hidden bookmarks in the display copy, never in
   the saved file.** 0.88.0 ignores `w14:paraId` but keeps bookmark names
   on its model paragraphs, so the pre-pass marks every paragraph with a
   `_wd<id>` bookmark; Word-written files keep their own `w14:paraId` as
   the id, files without them get deterministic ids that `docx-edit`
   recomputes from the original bytes, and the saved file does not change
   for a read. Bookmarks are invisible to layout and to the text map.
3. **XLSX moves with DOCX.** One package, one version; the spike renders the
   workbook fixtures before the bump so an XLSX regression is seen first.

## Actual result

- **T51 (2026-10-02)**: `src/adapters/docx-prepass.ts` on the package layer
  fits oversized inline pictures section by section (`wp:extent` and the
  picture's `a:ext`, anchored pictures untouched) and marks every paragraph
  of the body, headers, footers, footnotes and endnotes with its `_wd<id>`
  bookmark; the Office adapter runs it on every DOCX before the engine
  loads (the old model patch stays as a no-op fallback until the bump).
  Unit tests (`docx-prepass.test.ts`, 6) cover sections, tables, `w:sdt`,
  authored ids, existing bookmarks, empty paragraphs, unreadable input and
  the fixture; `docx-inline-images.spec.ts` passes through the pre-pass on
  the shipped engine; after the review, ids and bookmark numbers are unique
  across every story part of a document (one counter threads body,
  headers, footers, footnotes and endnotes) and empty paragraphs are
  rewritten as one element patch; the spike now renders the pre-passed fixture
  identically on 0.88 (SSIM 1.000 for both fixtures) and resolves all 228
  runs of `sample.docx` to their paragraph through the bookmark bridge.
- **T52 (2026-10-02, after approval)**: `@silurus/ooxml` is 0.88.0 for
  DOCX, XLSX and PPTX; the `@silurus/ooxml-pptx` alias is gone from
  `package.json`, the lockfile, the adapter, `scripts/example.mjs` (one
  `vendor/ooxml` copy) and `scripts/size-report.mjs`; the adapter no longer
  touches the engine's model (`fitInlineImagesToPage` stays exported but
  deprecated for hosts that applied it to their own models, removal in the
  next major). Before the bump a pixel comparison of the two engines on
  `sample.docx` (deterministic fonts, the smoke and the fidelity set-ups)
  found 0 differing pixels in 816 × 1056, so the stored Linux snapshots
  hold. The spike spec became the regression `tests/e2e/docx-engine.spec.ts`:
  viewer against bare engine on the pre-passed bytes SSIM 1.000 for both
  fixtures, the pre-pass moves the oversized picture (original against
  pre-passed SSIM 0.365) and changes nothing on `sample.docx` (1.000), 228 of
  228 runs resolve through the bookmark bridge, layout complete. Gates:
  matrix 176/176 (Chromium, Chromium DPR 2, Firefox, WebKit), unit 360/360,
  `npm run licenses` pass (9 npm runtime packages). Size: the DOCX WASM goes
  from 892,883 to 1,947,090 bytes raw (297,932 → 589,532 Brotli) and the XLSX
  WASM from 815,854 to 1,725,510 (277,267 → 524,370 Brotli); the PPTX WASM
  was already 0.88; the base total grows by 547,068 bytes Brotli to 6.21 MiB,
  well under the 20 MiB target, and the example vendors one engine directory
  instead of two. The ticket's "the bundle does not grow" is therefore met
  only for the number of engine copies, not for the bytes a DOCX or XLSX
  viewer fetches; recorded here for the release note.

## Open questions

- None for the bump. The release and the monorepo verification on DOCX
  previews stay Leonid's.
