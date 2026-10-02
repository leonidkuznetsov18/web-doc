# Implementation plan: document editing — `edit-core` and `pdf-edit`

## Overview

This plan implements modules 01 and 02 of the `document-editing` package:
[`edit-core`](../docs/document-editing/todo/01-edit-core.md) and
[`pdf-edit`](../docs/document-editing/todo/02-pdf-edit.md), both approved on
2026-10-01. When it is done, `viewer.edit()` on a loaded PDF returns a
`PdfEditSession` with 15 typed methods backed by PDFium WASM in a worker; every
change is reopened by PDF.js in the viewer, and `save()` returns an incremental
update of the original file. The task list lives in [`todo.md`](./todo.md).

## Architecture decisions

1. **Edit state = original bytes + history.** The engine owns a working copy;
   undo, redo, reset and failure recovery call `engine.restore(prefix)`, which
   reloads the original and replays the batches. Element ids are derived from
   the original order or from (batch, operation, index), so a replay produces
   the same ids and the same bytes.
2. **Engines plug in through the adapter.** `DocumentAdapter` gains two optional
   members: `edit?: EditEngineProvider` and `reopen?(previous, bytes, context)`.
   The provider types live in `src/edit/engine.ts` and are not re-exported from
   the package root. The core never imports a format engine statically.
3. **Validation without a new dependency.** A JSON-only shape check plus a small
   in-repo validator for the JSON Schema 2020-12 keywords the operations use.
   The same schemas are exported through `session.schemas` for AI clients.
4. **Eager refresh inside `apply()`.** Engine apply → materialize → adapter
   reopen → swap handles → events. Any failure restores the engine to the
   previous prefix and keeps the old handle.
5. **Viewport keeps the view.** A new `replaceDocument(info)` path keeps zoom
   and scroll, adds the document revision to the render key, and drops slots
   so visible pages re-render. `pageToClient` and `clientToPage` are computed
   from the same page metrics the viewport lays out with.
6. **Environment-agnostic PDF engine.** `src/edit/pdf/engine/**` runs inside the
   worker (bundled by esbuild into `dist/workers/pdf-edit-worker.js`) and in
   Node unit tests (PDFium WASM read from `node_modules`). Platform services are
   injected: WASM loading, font fetching, image decoding.
7. **Coordinates mirror PDF.js.** Page space ↔ PDF user space is derived per
   page from the crop box clipped to the media box and from `/Rotate`, the way
   PDF.js builds its viewport; browser tests compare against PDF.js text
   positions.
8. **Parametric elements persist through marked content.** Text boxes and
   tables carry a `WebDoc` mark whose string parameter holds their JSON inputs;
   marks read from files are validated like operations.
9. **Fonts.** Standard fonts through `FPDFText_LoadStandardFont`; TrueType
   through `FPDFText_LoadFont` as CID fonts — host-registered fonts are fetched
   on the main thread and transferred, the bundled fallback TTF is fetched by
   the worker from the asset base.
10. **Deterministic fixtures.** A test-only minimal PDF builder (pages, media
    and crop boxes, rotation, Helvetica text, an image, a path) feeds unit tests
    and, from the Node side of Playwright, the browser tests. No binary fixtures
    are committed.

## Dependency graph

```text
T1 PDFium bridge and probes ────────────────────────────────┐
T2 Core contracts ── T3 Schema validator ── T4 Session core ┤
                                            T5 Headless viewer integration
                                            ├── T6 Viewport refresh and UI consistency
                                            ├── T7 Geometry helpers
                                            └── T8 Core docs                 ── Checkpoint A
T9 PDF worker, provider, session start (T1, T5)
├── T10 Release plumbing for PDF assets
└── T11 PDF document model and inspection
    └── T12 insertTextBox
        └── T13 Text box editing and persistence        ── Checkpoint B (first demo)
            ├── T14 Move, resize, delete
            ├── T15 Page operations
            ├── T16 Fonts: fallback TTF and registered fonts
            │   └── T17 Existing text edits                ── Checkpoint C
            ├── T18 Shapes
            ├── T19 Images
            └── T20 Tables                                  ── Checkpoint D
T21 Hardening (signatures, file id, crash recovery, performance, JSON-only client)
T22 PDF docs, browser matrix, full gate                     ── Checkpoint E (done)
T23 Contract revision 2: types and docs (approved spec)      [linear:ACTION-821]
├── T24 Session fixes: snapshot, listeners, aborted on end, save/markSaved, broken-session save
├── T25 Two-phase reopen, renderer page count, layoutchange, per-page render keys
├── T26 stateId ids, removedIds, same-batch references, applyJson, envelopes, read signals
├── T27 Checkpoints and the asset store
└── T28 PDF: save modes, mark staleness, feature warnings, findText ranges, subset-font and font-dedupe tests
    └── T29 apply() latency on 10/100/500 pages (PDF)
        └── T30 Docs, matrix, full gate                      ── Checkpoint F
```

T1 runs first because it is the riskiest assumption; it does not depend on the
core and can run in parallel with T2–T4. After Checkpoint B, T14–T20 depend only
on T13 and could be split between sessions; they touch separate engine files
but share `schemas.ts`, so they are planned sequentially to avoid conflicts.

## Phases

| Phase                      | Tasks   | Ends with                                                                                                                                                                                     |
| -------------------------- | ------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0. Risk first              | T1      | PDFium mechanics proven in Node, or the spec updated                                                                                                                                          |
| 1. `edit-core`             | T2–T8   | Checkpoint A: core done on all browsers, human review                                                                                                                                         |
| 2. First PDF slice         | T9–T13  | Checkpoint B: insert and edit a text box in a real PDF, save, reload — demo                                                                                                                   |
| 3. PDF method set          | T14–T20 | Checkpoints C and D: all 15 operations                                                                                                                                                        |
| 4. Hardening and readiness | T21–T22 | Checkpoint E: both definitions of done met, `npm run check` green, ready for merge review                                                                                                     |
| 5. Contract revision 2     | T23–T30 | Checkpoint F: `edit-core` R2 and `pdf-edit` R2 done, latency recorded (Linear ACTION-821)                                                                                                     |
| 6. PDF overlay primitives  | T32–T37 | Checkpoint G: layout, suppressed render, selection and range mapping, range-scoped `replaceText`, geometry cache, browser test (Linear ACTION-825)                                            |
| 7. OOXML package layer     | T38–T43 | Checkpoint H: ZIP reader and writer, OPC model, XML scanner, patches and transactions, corpus and browser reopen (Linear ACTION-810)                                                          |
| 8. PPTX editing            | T44–T49 | Checkpoint I: inspection, text, shapes, images and tables, slides, browser round trip and latency on 10/100/500 slides (Linear ACTION-812)                                                    |
| 9. DOCX engine upgrade     | T50–T53 | Checkpoint J: spike, XML pre-pass (image fitting, paragraph ids), the approved bump to one engine copy, the run bridge (Linear ACTION-813)                                                    |
| 10. DOCX editing           | T54–T58 | Checkpoint K: block index and inspection joined with the renderer's runs, text and formatting, structure and pictures, tables, round trip and latency on 10/100/500 pages (Linear ACTION-814) |

## Revision 2 decisions (ACTION-821)

11. **Ids from `stateId`.** The core passes each batch's history `stateId` to
    the engine; PDF ids become `p0:n<stateId>.<op>.<k>` and page keys
    `q<stateId>.<op>`. Replays reproduce them because `stateId`s are stored
    in the history.
12. **Two-phase reopen.** `prepareDocument` opens the edited bytes next to the
    current handle and may fail; `commitDocument` swaps synchronously and
    cannot; abort after the commit point is ignored.
13. **Checkpoints are retained `materialize("show")` outputs** at the fold
    boundary and every `maxEditHistory / 4`-th commit, within
    `maxEditCheckpointBytes`; `restore` takes `{ base, batches }`.
14. **Asset store.** Inline payloads are hashed (SHA-256) and interned before a
    batch enters the history; operations in history hold `asset:` references;
    the worker keeps the bytes once per session.
15. **Read envelopes without paging**; `sessionId` is a random 128-bit value.
16. **PDF saves default to full** for unsigned files; the viewer reopen keeps
    using the incremental form.

## Phase 6 decisions (ACTION-825)

16. **One worker request per primitive.** `edit-text-layout`,
    `edit-position-at`, `edit-range-rects` and `edit-render-without` join the
    protocol; the core session gains a `read()` hook that queues any read
    behind earlier calls and stamps the envelope, so the PDF session adds
    methods without touching the core's queue.
17. **Lines are objects.** A layout line is one PDFium text object: a text
    box's lines and a table's cells are separate objects already, and a plain
    text object is one line. Glyph geometry comes from the text page
    (`FPDFText_GetCharBox`, `GetLooseCharBox`, `GetCharOrigin`) mapped through
    `PageGeometry`, never from PDF.js.
18. **Suppressed render returns pixels, not a reopen.** `renderPageWithout`
    hands the host RGBA pixels at a requested scale; drawing them over the
    page is the host's job, and the session's bytes stay as they were.
19. **The selection ladder is GenOffice's.** Overlap ≥ 50 %, then containment,
    then an NFKC text match — attributed under Apache-2.0 in
    `THIRD_PARTY_NOTICES.md`; the code is web-doc's own.
20. **Range mapping walks the history.** `mapRange` derives the moved range
    from the entries after `fromRevision` (text length deltas, deletions,
    `remappedIds`), so it needs no extra bookkeeping in the engine.

## Phase 7 decisions (ACTION-810)

21. **The package layer is spec-bound.** `03-ooxml-package.md` (approved
    2026-10-02) fixes the API, the error codes and nine decisions; the tasks
    here only sequence the work: reader, writer, OPC model, scanner, patches
    and transactions, then the corpus and browser gate.
22. **Tests ship a ZIP builder of their own.** Hand-built archives (stored,
    deflated with Node's zlib, data descriptors, extra fields, ZIP64 markers,
    encryption flags, bad CRCs) exercise the reader before the writer exists;
    the writer later proves itself against the same builder and the corpus.

## Gate and configuration changes needing approval

These follow from the approved definitions of done but touch gates, so they are
listed here for explicit approval with the plan:

- `scripts/size-report.mjs`: also measure the PDFium WASM, the PDF edit worker
  and `.ttf` font files. Thresholds stay unchanged.
- `scripts/pack-test.mjs`: require the new `dist/` files.
- `playwright.matrix.config.ts`: add `edit-core.spec.ts` and `edit-pdf.spec.ts`
  to the matrix `testMatch` list.
- `packages/viewer/fonts/`: commit a TTF build of the bundled Noto Sans
  Latin/Cyrillic WOFF2 with a manifest entry (bytes and SHA-256), produced once
  with fontTools in a scratch virtual environment; no new project dependency.

## Risks and mitigations

| Risk                                                                                          | Impact | Mitigation                                                                                                                          |
| --------------------------------------------------------------------------------------------- | ------ | ----------------------------------------------------------------------------------------------------------------------------------- |
| Incremental save is not reachable through the wrapper (`PDFiumExt_SaveAsCopy` takes no flags) | High   | T1 calls `FPDF_SaveAsCopy` with `FPDF_INCREMENTAL` on the wrapper's file writer; if impossible, update the spec to full saves first |
| PDFium writes a random or time-based file id, breaking determinism                            | Medium | T1 measures it; normalize the id deterministically after saving (T21)                                                               |
| `addFunction` cannot grow the table, so inline JPEG loading fails                             | Low    | Decode JPEG to RGBA and embed losslessly; record the larger output in the spec                                                      |
| Regenerating a touched page's content changes how it renders                                  | Medium | Only touched pages are regenerated; read-back checks; visual comparison of untouched areas in browser tests                         |
| Coordinates disagree with PDF.js on rotated or cropped pages                                  | High   | Fixtures for all four rotations and an offset crop box; browser cross-check against PDF.js text positions                           |
| One reopen per `apply()` is slow for large PDFs                                               | Medium | Reuse the PDF.js worker on reopen (T9); batching guidance in docs; performance test with a 3-second ceiling (T21)                   |
| Subset fonts lack glyphs for replaced text                                                    | Medium | Safety rule, read-back and fallback with a warning (T17)                                                                            |
| PDFium heap growth across sessions                                                            | Low    | One worker per session, terminated when the session ends                                                                            |
| First browser run builds Rust/WASM and installs `wasm-bindgen-cli` (minutes)                  | Low    | One-time cost; the development loop uses Chromium only                                                                              |

## Phase 8 decisions (ACTION-812)

23. **The PPTX engine is spec-bound.** `04-pptx-edit.md` (draft 2026-10-02,
    written under the instruction to execute the Linear plan without
    stopping) fixes the operation table, the element model and thirteen
    decisions; the tasks implement it and record Spike results and the
    Actual result in it. Deviations go into the spec first.
24. **Inspection from the XML, not the renderer.** Bounds, inheritance and
    group transforms are computed by the engine in the worker; the
    renderer's `getElementBoundsByIds` is a browser-test oracle only
    (04-pptx-edit decision 1).
25. **One OOXML edit worker.** `src/ooxml-edit-worker.ts` serves PPTX now and
    DOCX in Phase 10; the PDF worker client's transport becomes a shared base
    class so the three clients differ only in their format reads.
26. **Ids `<slideKey>:<cNvPrId>` with next-free allocation**, so files look
    as if PowerPoint wrote them and replays reproduce ids; the reuse caveat
    after deleting a created slide is documented (04-pptx-edit decision 3).
27. **Full reopen per `apply()` with a three-second ceiling**, measured on
    synthetic 10-, 100- and 500-slide decks that the renderer opens; the
    spike also tries `progressiveLayout` on reopen.

## Phase 9 decisions (ACTION-813)

28. **The bump waits for approval; the pre-pass does not.** The spike (T50)
    and the XML pre-pass (T51) run behind the existing alias, so the only
    change that needs Leonid's word — `@silurus/ooxml` 0.72.2 → 0.88.0 for
    DOCX and XLSX — is isolated in T52 (05-docx-engine-upgrade). Approved
    and done on 2026-10-02: one engine copy, alias and model patch retired.
29. **Generated paragraph ids live in the display copy only.** Files without
    `w14:paraId` get deterministic ids before the engine reads them;
    `docx-edit` recomputes the same ids from the original bytes, so saved
    files do not change for a read (05-docx-engine-upgrade decision 2).

## Phase 10 decisions (ACTION-814)

30. **The DOCX engine is spec-bound and waits for module 05.** `06-docx-edit.md`
    (draft 2026-10-02) fixes the element model, the operation table and six
    decisions; T54 starts after T52's bump and T53's run bridge, which the
    inspection join depends on.

## Phase 11 decisions (`ai-edit`, module 07)

31. **The AI module is additive and host-agnostic.** `07-ai-edit.md` (draft
    2026-10-02) adds outline, description, target resolution, named
    checkpoints, a tool set with an executor, and DOCX tracked changes to the
    sessions of modules 01–06 without changing any existing shape; T59 starts
    after Leonid approves the spec and its recommended answers.

## Verification commands

```bash
npm run typecheck --workspace web-doc
npm run test --workspace web-doc                      # all viewer unit tests
node --test packages/viewer/.test-dist/test/<file>.test.js   # one compiled test file, after a full run compiled it
npm run test:e2e -- tests/e2e/edit-core.spec.ts       # Chromium, builds first
npm run test:e2e -- tests/e2e/edit-pdf.spec.ts
npm run test:e2e -- tests/e2e/edit-pptx.spec.ts
npm run test:e2e -- tests/e2e/edit-docx.spec.ts
npm run test:e2e -- tests/e2e/edit-ai.spec.ts
npm run test:e2e:matrix                               # at checkpoints
npm run licenses && npm run report:size && npm run test:pack
npm run check                                         # at the final checkpoint
```

## Open questions

None blocking. Phase 5 starts only after Leonid approves `01-edit-core.md`
revision 2; its commits carry `[linear:ACTION-821]`.
