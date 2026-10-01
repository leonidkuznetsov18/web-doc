# Tasks: `edit-core` and `pdf-edit`

Plan: [`plan.md`](./plan.md). Specs:
[`01-edit-core.md`](../docs/document-editing/todo/01-edit-core.md),
[`02-pdf-edit.md`](../docs/document-editing/todo/02-pdf-edit.md).

Every task also clears the standing bar: typecheck clean, the viewer unit suite
green, no weakened tests or gates, docs updated in the same change when public
behaviour changes.

## Phase 0 — risk first

### Task 1: PDFium bridge and feasibility probes

**Description:** Add `@embedpdf/pdfium@2.15.1` (exact pin) and a thin,
environment-agnostic bridge: load the module from WASM bytes, open and close
documents from bytes, move strings and buffers across the WASM boundary, save
through the wrapper's file writer. Node tests probe every PDFium mechanism the
`pdf-edit` spec relies on. Findings are written into the spec; a failed probe
updates the spec before T9.

**Acceptance criteria:**

- [x] Opening a PDF and saving it with `FPDF_INCREMENTAL` after adding a
      standard-font text object yields the original bytes followed by an
      appended update; reopening shows the text. Two runs of the same edit give
      identical bytes, or the cause of any difference is recorded.
- [x] A marked-content `WebDoc` tag with a string parameter survives save and
      reopen; page insert, delete, move and rotate survive save and reopen;
      inline JPEG loading through `addFunction` works, or the fallback is
      recorded.
- [x] `npm run licenses` passes and `THIRD_PARTY_NOTICES.md` lists PDFium
      (BSD-3-Clause) and the EmbedPDF wrapper (MIT).

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdfium-bridge.test.ts`)
- [x] `npm run typecheck --workspace web-doc` and `npm run licenses`
- [x] Spec `02-pdf-edit.md` has a "Spike results" section

**Dependencies:** None

**Files likely touched:**

- `packages/viewer/package.json`, `package-lock.json`
- `THIRD_PARTY_NOTICES.md`
- `packages/viewer/src/edit/pdf/engine/pdfium.ts`
- `packages/viewer/test/pdfium-bridge.test.ts`
- `docs/document-editing/todo/02-pdf-edit.md`

**Estimated scope:** Medium

## Phase 1 — `edit-core`

### Task 2: Core editing contracts

**Description:** Add the public types from the `edit-core` spec
(`src/edit/types.ts`), the engine interface (`src/edit/engine.ts`, not
re-exported), the new `ViewerApi` members (temporarily rejecting with
`edit-unsupported`), the two events, four error codes,
`DocumentCapabilities.editing` (always `false` for now), the optional
`DocumentAdapter.edit` and `reopen` members, and the `maxEditOperations` and
`maxEditHistory` limits.

**Acceptance criteria:**

- [x] All spec names are exported from `web-doc` and `web-doc/headless`; the
      engine interface is not.
- [x] `resolveLimits` validates the new limits; defaults are 500 and 200.
- [x] Existing unit tests pass unchanged; `edit()` on any document rejects with
      `edit-unsupported`.

**Verification:**

- [x] `npm run typecheck --workspace web-doc`
- [x] `npm run test --workspace web-doc` (new `edit-contracts.test.ts`)

**Dependencies:** None

**Files likely touched:**

- `packages/viewer/src/edit/types.ts`, `packages/viewer/src/edit/engine.ts`
- `packages/viewer/src/contracts.ts`, `packages/viewer/src/limits.ts`
- `packages/viewer/src/index.ts`, `packages/viewer/src/headless.ts`
- `packages/viewer/test/edit-contracts.test.ts`

**Estimated scope:** Medium

### Task 3: Operation shape checks and schema validator

**Description:** Implement the JSON-only check, `BinaryData` normalisation
(base64 string or `Uint8Array`), and a validator for the JSON Schema 2020-12
subset the operations use (`type`, `properties`, `required`,
`additionalProperties`, `enum`, `const`, numeric and length bounds, `pattern`,
`items`, `minItems`/`maxItems`, `oneOf`, local `$ref`, plus a binary marker).
It reports every issue with operation index, JSON pointer, code and message.

**Acceptance criteria:**

- [x] Functions, class instances, `undefined` in arrays, `NaN`, `Infinity`
      and cycles are rejected; `Uint8Array` is accepted only where the schema
      marks binary data.
- [x] All issues of a batch are collected (not just the first), with stable
      codes and correct JSON pointers.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `edit-schema.test.ts`)

**Dependencies:** Task 2

**Files likely touched:**

- `packages/viewer/src/edit/schema.ts`
- `packages/viewer/src/edit/operations.ts`
- `packages/viewer/test/edit-schema.test.ts`

**Estimated scope:** Small

### Task 4: Session core

**Description:** Implement `EditHistory` and the session controller independent
of the viewer: preconditions, shape and engine validation, dry run, apply,
commit, undo, redo, reset, save, `dirty`, `maxEditHistory` folding, the FIFO
call queue, abort handling and receipts. The viewer side is behind a host
interface (reopen and emit) so it can be tested with fakes.

**Acceptance criteria:**

- [x] Every rule in the spec sections "Applying a batch", "History, revisions
      and dirty state" and "Saving" has a passing test against a fake engine and
      a fake host.
- [x] A failure in engine apply, materialize or host reopen, or an abort during
      apply, restores the previous prefix and rejects with the right code and
      `stage`.
- [x] Undo to revision 0 materializes the original bytes; replaying a history
      materializes identical bytes.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `edit-history.test.ts`,
      `edit-session.test.ts`)

**Dependencies:** Task 3

**Files likely touched:**

- `packages/viewer/src/edit/history.ts`
- `packages/viewer/src/edit/session.ts`
- `packages/viewer/test/fixtures/fake-edit-engine.ts`
- `packages/viewer/test/edit-history.test.ts`
- `packages/viewer/test/edit-session.test.ts`

**Estimated scope:** Medium

### Task 5: Headless viewer integration

**Description:** Wire the session into `DocumentViewer`: `edit()` and
`getEditSession()`, `capabilities.editing`, lazy provider lookup on the
adapter, reopen through `adapter.reopen ?? adapter.open` with handle swap,
`ViewerState` and `DocumentInfo` updates, invalidation of text maps, fuzzy index,
search and selection, events, and ending the session on `load`, `close` and
`destroy`.

**Acceptance criteria:**

- [x] With a fake adapter and engine, `getDocumentInfo`, `getPageText`,
      `search` and `selectText` reflect the new content as soon as `apply()`
      resolves, including a changed page count.
- [x] `load`, `close` and `destroy` end the session: pending calls reject with
      `aborted`, later calls with `lifecycle-error`, and `editstatechange`
      reports `active: false`.
- [x] The provider is not touched before the first `edit()`.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `edit-viewer.test.ts`; existing
      suites unchanged)

**Dependencies:** Task 4

**Files likely touched:**

- `packages/viewer/src/viewer.ts`
- `packages/viewer/src/edit/session.ts`
- `packages/viewer/test/edit-viewer.test.ts`
- `packages/viewer/test/fixtures/fake-edit-engine.ts`

**Estimated scope:** Medium

### Task 6: Viewport refresh and built-in UI consistency

**Description:** Add a view-preserving document replacement to the viewport
(zoom, fit and scroll kept and clamped; revision in the render key; slots
re-rendered), and make the optional built-in UI refresh its page counter and
thumbnails on `documentchange`.

**Acceptance criteria:**

- [x] In the browser, after `apply`, `undo` and `redo` through a test adapter,
      the visible page's pixels change accordingly while zoom, fit and scroll
      are kept and search highlights disappear.
- [x] A change in page count updates the layout and the UI page counter.

**Verification:**

- [x] `npm run test:e2e -- tests/e2e/edit-core.spec.ts`

**Dependencies:** Task 5

**Files likely touched:**

- `packages/viewer/src/viewport.ts`
- `packages/viewer/src/viewer.ts`
- `packages/viewer/src/ui.ts`
- `tests/e2e/edit-core.spec.ts`

**Estimated scope:** Medium

### Task 7: View-geometry helpers

**Description:** Implement `pageToClient` and `clientToPage` from the
viewport's page metrics, returning `undefined` for headless viewers, unmounted
pages and spreadsheets.

**Acceptance criteria:**

- [x] Results agree with the rendered canvas within 1 CSS pixel at zoom 0.5, 1
      and 2, device pixel ratio 1 and 2, and after scrolling;
      `clientToPage(pageToClient(r))` round-trips.
- [x] Headless viewers, unmounted pages and spreadsheets return `undefined`.

**Verification:**

- [x] `npm run test:e2e -- tests/e2e/edit-core.spec.ts`
- [x] `npm run test --workspace web-doc`

**Dependencies:** Task 6

**Files likely touched:**

- `packages/viewer/src/viewport.ts`
- `packages/viewer/src/spreadsheet-viewport.ts`
- `packages/viewer/src/viewer.ts`
- `tests/e2e/edit-core.spec.ts`
- `packages/viewer/test/edit-viewer.test.ts`

**Estimated scope:** Small

### Task 8: Core documentation

**Description:** Write `docs/api/editing.md` (lifecycle, operations and
schemas, receipts, errors, events, page space, geometry helpers, guidance for AI
clients), update `docs/api/reference.md` and `docs/architecture.md`, and add the
page to the documentation sidebar.

**Acceptance criteria:**

- [x] Every public type, method, event, error code, limit and flag of the core
      is documented.
- [x] The documentation site builds.

**Verification:**

- [x] `npm run pages:build`

**Dependencies:** Task 7

**Files likely touched:**

- `docs/api/editing.md`, `docs/api/reference.md`, `docs/architecture.md`
- `docs/.vitepress/config.ts`

**Estimated scope:** Small

### Checkpoint A: `edit-core` complete

- [x] `npm run typecheck`, `npm test` and the existing browser suites pass
- [x] `edit-core.spec.ts` passes in the matrix (Chromium, Chromium DPR 2,
      Firefox, WebKit) after adding it to `playwright.matrix.config.ts`
- [x] The size report's `code` group grew by at most 20 KB Brotli
- [x] `01-edit-core.md` gets an "Actual result" section and status "Done"
- [x] Review with the human before PDF work starts

## Phase 2 — first PDF slice

### Task 9: PDF edit worker, provider and session start

**Description:** Add the worker entry and its RPC operations, a main-thread
engine client that implements the engine interface over RPC, the PDF adapter's
`edit` provider and `reopen` (reusing the PDF.js worker), the build steps that
bundle the worker and copy `pdfium.wasm`, and the `PdfEditSession` type with an
empty method set.

**Acceptance criteria:**

- [x] In the browser, `edit()` on a PDF loads the worker and PDFium; neither
      `pdf-edit-worker.js` nor `pdfium.wasm` is requested before `edit()`;
      `save()` without changes returns the original bytes; ending the session
      terminates the worker.
- [x] A crashed worker rejects pending calls with a typed error, and `edit()`
      can be called again afterwards.
- [x] `session.format` narrows the session union in a compile-only test.

**Verification:**

- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`
- [x] `npm run test --workspace web-doc`

**Dependencies:** Tasks 1 and 5

**Files likely touched:**

- `packages/viewer/src/pdf-edit-worker.ts`
- `packages/viewer/src/edit/pdf/provider.ts`, `packages/viewer/src/edit/pdf/types.ts`
- `packages/viewer/src/adapters/pdf.ts`
- `scripts/build-viewer.mjs`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Task 10: Release plumbing for PDF assets

**Description:** Make the gates see the new assets: the size report measures the
PDFium WASM and the edit worker; the pack test requires them.

**Acceptance criteria:**

- [x] `npm run report:size` lists the PDFium WASM and the edit worker and stays
      within the 20 MiB target.
- [x] `npm run test:pack` passes with the new files required.

**Verification:**

- [x] `npm run report:size` and `npm run test:pack`

**Dependencies:** Task 9

**Files likely touched:**

- `scripts/size-report.mjs`
- `scripts/pack-test.mjs`

**Estimated scope:** Small

### Task 11: PDF document model and inspection

**Description:** In the engine, build the document model over PDFium: page ids,
element ids, conversion between page space and user space, element kinds and
styles, `getElements`, `getElement`, `elementsAt` and `findText`, plus
`materialize` and `restore`. Add the test-only PDF builder.

**Acceptance criteria:**

- [x] On builder fixtures with rotations 0°, 90°, 180°, 270° and an offset crop
      box, element bounds match the expected page-space rectangles within
      0.5 pt.
- [x] Text, image, shape and other objects get the right kinds, text and
      styles; `findText` returns rectangles and element ids.
- [x] Ids are identical after `restore` of the same prefix.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-inspect.test.ts`)

**Dependencies:** Task 9

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/document.ts`
- `packages/viewer/src/edit/pdf/engine/geometry.ts`
- `packages/viewer/src/edit/pdf/engine/elements.ts`
- `packages/viewer/test/fixtures/pdf-builder.ts`
- `packages/viewer/test/pdf-edit-inspect.test.ts`

**Estimated scope:** Medium

### Task 12: `insertTextBox`

**Description:** Implement the first operation end to end in the engine: text
layout (paragraphs, greedy wrapping, alignment, line height), standard fonts,
colour, the `WebDoc` mark, the operation's JSON Schema and the typed
`insertTextBox` method.

**Acceptance criteria:**

- [x] After `insertTextBox`, save and reopen, PDFium extracts the text, the line
      positions follow the requested alignment and wrapping, and the element is
      listed as one `textBox` with its style.
- [x] Invalid input (empty rect, bad colour, out-of-range size, unknown font)
      fails with `invalid-operation` issues and changes nothing.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-textbox.test.ts`)

**Dependencies:** Task 11

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/text-layout.ts`
- `packages/viewer/src/edit/pdf/engine/text-box.ts`
- `packages/viewer/src/edit/pdf/engine/marks.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-textbox.test.ts`

**Estimated scope:** Medium

### Task 13: Text box editing and persistence

**Description:** `replaceText`, `setTextStyle` and `resizeElement` on text boxes
(rebuilt from their parameters), recognition of `WebDoc` tags when a saved file
is reopened, and the first browser round trip.

**Acceptance criteria:**

- [x] Editing a text box rebuilds it; undo and redo restore the exact earlier
      bytes.
- [x] In the browser, `insertTextBox` shows the text on the canvas;
      `getPageText` contains it after save and reload in a fresh viewer; the
      reloaded session lists one `textBox`.

**Verification:**

- [x] `npm run test --workspace web-doc`
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 12

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/text-box.ts`
- `packages/viewer/src/edit/pdf/engine/marks.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-textbox.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Checkpoint B: first PDF slice (demo)

- [x] Unit and `edit-pdf.spec.ts` (Chromium) pass; existing suites pass
- [ ] Demo for the human: open a PDF, insert and restyle a text box, undo/redo,
      save, reopen
- [ ] Decide with the human whether the remaining operations keep this order

## Phase 3 — PDF method set

### Task 14: Move, resize, delete

**Description:** `moveElement`, `resizeElement` and `deleteElement` for every
element kind, including groups.

**Acceptance criteria:**

- [x] After save and reopen, moved and resized elements have the expected
      bounds (±0.5 pt) on all four rotations; deleted elements are gone and
      untouched pages keep their content streams byte for byte.
- [x] `moveElement` with both or neither of `to` and `by` fails validation.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-transform.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 13

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/transform.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/src/edit/pdf/types.ts`
- `packages/viewer/test/pdf-edit-transform.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Task 15: Page operations

**Description:** `insertPage`, `deletePage`, `movePage` and `rotatePage`, with
element ids stable across them.

**Acceptance criteria:**

- [x] After each operation, save and reopen show the expected page count, order,
      sizes and rotation in PDF.js; the viewer's page count updates.
- [x] Deleting the last page fails validation; ids of elements on other pages
      survive page moves, undo and redo.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-pages.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 13

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/pages.ts`
- `packages/viewer/src/edit/pdf/engine/document.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-pages.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Task 16: Fonts — fallback TTF and registered fonts

**Description:** Produce the TTF build of the bundled Noto Sans Latin/Cyrillic
face, record it in the font manifest and notices, and implement font
resolution: standard fonts for WinAnsi text, host-registered TrueType/OpenType
fonts, then the fallback; `font-unavailable` issues; `font-substitution`
warnings for missing bold or italic faces. The size report counts `.ttf` fonts.

**Acceptance criteria:**

- [x] A Cyrillic text box inserted without host fonts is extracted correctly by
      PDF.js after save; with a registered TTF family, that font is embedded.
- [x] Text no available font covers fails with a `font-unavailable` issue; the
      license gate verifies the TTF's size and hash.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-fonts.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`
- [x] `npm run licenses` and `npm run report:size`

**Dependencies:** Task 13

**Files likely touched:**

- `packages/viewer/fonts/noto-sans-latin-cyrillic.ttf`,
  `packages/viewer/fonts/manifest.json`,
  `packages/viewer/fonts/THIRD_PARTY_NOTICES.md`
- `packages/viewer/src/edit/pdf/engine/fonts.ts`
- `scripts/size-report.mjs`
- `packages/viewer/test/pdf-edit-fonts.test.ts`

**Estimated scope:** Medium

### Task 17: Edits to existing text

**Description:** `replaceText` and `setTextStyle` (`color`, `fontSize`) on
existing text objects: the in-place safety rule, read-back verification,
fallback replacement with a `font-substitution` warning, and the
`unsupported-style` and `unsupported-script` issues.

**Acceptance criteria:**

- [x] With a non-subset font the text changes in place; with a subset font
      lacking a character the object is replaced in the fallback font at the
      same position, size and colour, with a warning.
- [x] Unsupported style fields and right-to-left text fail validation with the
      spec's issue codes.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-existing-text.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 16

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/existing-text.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-existing-text.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Checkpoint C: text complete

- [x] Unit and `edit-pdf.spec.ts` pass; the size report and license gate pass
- [x] Undo/redo byte identity holds across every text operation

### Task 18: Shapes

**Description:** `insertShape` (rectangle, ellipse, line) and `setShapeStyle`
(stroke and fill, `null` to remove).

**Acceptance criteria:**

- [x] Shapes appear with the requested geometry and colours after save and
      reopen; `setShapeStyle` changes them; a shape with neither stroke nor fill
      fails validation.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-shapes.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 14

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/shapes.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-shapes.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Small

### Task 19: Images

**Description:** `insertImage`: JPEG embedded as is, PNG decoded (in the worker
with `createImageBitmap` and `OffscreenCanvas`; injected decoder in Node) with
alpha as a soft mask; pixel and byte limits.

**Acceptance criteria:**

- [x] A JPEG and a transparent PNG appear at the requested rectangle after save
      and reopen; the JPEG stream is stored without re-encoding.
- [x] Oversized images fail with `resource-limit`; malformed data fails
      validation.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-images.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 14

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/images.ts`
- `packages/viewer/src/pdf-edit-worker.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-images.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Task 20: Tables

**Description:** `insertTable` (grid paths, wrapped cell text, row heights,
header fill) and `setTableCell` (rebuild from parameters), persisted through the
`WebDoc` mark.

**Acceptance criteria:**

- [x] A table appears with the requested rows, columns and widths; cell text is
      extracted in reading order after save; after reopen it is one `table`
      element with its rows.
- [x] `setTableCell` rebuilds the table; more than 100 rows or 20 columns fails
      validation.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `pdf-edit-tables.test.ts`)
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Tasks 16 and 18

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/tables.ts`
- `packages/viewer/src/edit/pdf/schemas.ts`
- `packages/viewer/test/pdf-edit-tables.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Checkpoint D: all 15 operations

- [x] Unit and `edit-pdf.spec.ts` pass; every operation has a typed method, a
      schema, a unit test and a browser round trip

## Phase 4 — hardening and readiness

### Task 21: Hardening

**Description:** Signed-PDF warning, deterministic file id, worker-crash
recovery, `dryRun` on PDF batches, the performance check (3-second ceiling for
a one-operation apply on a 10-page PDF), and a JSON-only client test that drives
every operation through `apply()` using only `session.schemas`.

**Acceptance criteria:**

- [x] A signed fixture can be edited, the first change warns with
      `fidelity-degraded`, and the signed revision's bytes stay intact.
- [x] Two independent sessions with the same history produce identical bytes;
      a dry run changes nothing and returns the receipt a real apply would.
- [x] A client that knows only the schemas performs all 15 operations and
      receives typed issues for invalid ones.

**Verification:**

- [x] `npm run test --workspace web-doc`
- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 20

**Files likely touched:**

- `packages/viewer/src/edit/pdf/engine/document.ts`
- `packages/viewer/src/edit/pdf/provider.ts`
- `packages/viewer/test/pdf-edit-hardening.test.ts`
- `tests/e2e/edit-pdf.spec.ts`

**Estimated scope:** Medium

### Task 22: PDF documentation, browser matrix, full gate

**Description:** Write the PDF section of `docs/api/editing.md`, add
`edit-pdf.spec.ts` to the matrix, run every gate, and record the actual results
in both specs.

**Acceptance criteria:**

- [x] The PDF docs cover every method and field, element kinds, page space,
      fonts and their limits, annotations and links, signatures, and
      `findText` versus `search()`.
- [x] `npm run test:e2e:matrix` and `npm run check` pass.

**Verification:**

- [x] `npm run pages:build`, `npm run test:e2e:matrix`, `npm run check`

**Dependencies:** Task 21

**Files likely touched:**

- `docs/api/editing.md`
- `playwright.matrix.config.ts`
- `docs/document-editing/todo/01-edit-core.md`,
  `docs/document-editing/todo/02-pdf-edit.md`,
  `docs/document-editing/00-roadmap.md`

**Estimated scope:** Small

### Checkpoint E: both modules done

- [x] Every item of both definitions of done is checked
- [x] Roadmap statuses updated; ready for merge review with the human

## Phase 5 — contract revision 2 (Linear ACTION-821)

Gate: Leonid approves `docs/document-editing/todo/01-edit-core.md` revision 2
before any task below starts. Every commit carries `[linear:ACTION-821]`.

### Task 23: Revision 2 types and docs

**Description:** Add the R2 contract types (`sessionId`, envelopes,
`SavedDocument`, `TextPosition`/`TextRange`, `fragments`, `story`, `frame`,
`EditColor`, `LayoutChange`, `removedIds`/`remappedIds`, `EngineBatch`,
`maxEditCheckpointBytes`, `edit-conflict` details), update the engine interface
and write the R2 parts of `docs/api/editing.md` and `reference.md`.

**Acceptance criteria:**

- [x] Every R2 type in the spec exists and is exported; the compile-only test
      narrows the union and calls `applyJson` on it.
- [x] Docs describe the interaction model, envelopes, `save`/`markSaved`,
      assets, `$n` references, text ranges, `layoutchange` and the id rules.

**Verification:**

- [x] `npm run typecheck --workspace web-doc`, `npm run pages:build`

**Dependencies:** Approved spec

**Files likely touched:** `packages/viewer/src/edit/types.ts`,
`src/edit/engine.ts`, `src/contracts.ts`, `src/limits.ts`,
`docs/api/editing.md`, `docs/api/reference.md`

**Estimated scope:** Medium

### Task 24: Session fixes

**Description:** Synchronous batch snapshot before queueing; listener
isolation; queued calls reject with `aborted` on `end()`; pure `save()` and
`markSaved()`; last committed bytes so a broken session still saves;
`expectedSessionId`.

**Acceptance criteria:**

- [x] A batch mutated after `apply()` is applied as it was; a throwing listener
      never rejects a committed call; `end()` rejects queued calls with
      `aborted`; `markSaved` with a stale token keeps `dirty`; a broken session
      returns the last committed bytes.

**Verification:**

- [x] `npm run test --workspace web-doc` (`edit-session.test.ts`)

**Dependencies:** Task 23

**Files likely touched:** `src/edit/session.ts`, `src/viewer.ts`,
`test/edit-session.test.ts`

**Estimated scope:** Medium

### Task 25: Two-phase reopen, renderer page count, layoutchange

**Description:** `prepareDocument`/`commitDocument`/`discardDocument` on the
session host; `pageCount` taken from the renderer; `layoutchange` after the
viewport paints; render keys change only for `changedPages`.

**Acceptance criteria:**

- [x] An abort or a throwing listener after the commit point leaves viewer and
      engine on the same state; `layoutchange` follows `documentchange` and
      the geometry helpers are exact once it fired; untouched pages keep their
      bitmaps.

**Verification:**

- [x] `npm run test --workspace web-doc`, `npm run test:e2e -- tests/e2e/edit-core.spec.ts`

**Dependencies:** Task 24

**Files likely touched:** `src/viewer.ts`, `src/viewport.ts`,
`src/edit/session.ts`, `test/edit-viewer.test.ts`, `tests/e2e/edit-core.spec.ts`

**Estimated scope:** Medium

### Task 26: stateId ids, removedIds, same-batch references, envelopes

**Description:** `EngineBatch` with `stateId`; PDF ids and page keys derived
from it; `removedIds` from delete operations and undo/redo; `$n` references
resolved while applying; `applyJson`; read envelopes with `AbortSignal`.

**Acceptance criteria:**

- [x] Ids after undo differ from the undone ones; `removedIds` lists deleted
      elements and the elements of deleted pages; `$n` resolves and fails as
      specified; every read carries `sessionId` and `revision`.

**Verification:**

- [x] `npm run test --workspace web-doc`, `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 24

**Files likely touched:** `src/edit/session.ts`, `src/edit/pdf/engine/document.ts`,
`src/edit/pdf/provider.ts`, `src/edit/pdf/session.ts`, PDF tests

**Estimated scope:** Large (many expected ids in tests change)

### Task 27: Checkpoints and the asset store

**Description:** Retained checkpoints at the fold boundary and every
`maxEditHistory / 4`-th commit within `maxEditCheckpointBytes`;
`restore({ base, batches })`; SHA-256 interning of inline payloads;
`addAsset()`; `asset:` references in the PDF image operation.

**Acceptance criteria:**

- [x] A checkpoint restore yields the same ids and content as a replay from
      the original (bytes differ in layout, see the spec); an interned payload
      crosses to the worker once; an undo replays at most a quarter of the
      history.

**Verification:**

- [x] `npm run test --workspace web-doc`

**Dependencies:** Task 26

**Files likely touched:** `src/edit/history.ts`, `src/edit/session.ts`,
`src/edit/assets.ts`, `src/edit/pdf/engine/handler.ts`, `src/worker-protocol.ts`

**Estimated scope:** Medium

### Task 28: PDF revision 2

**Description:** `save({ mode })` with full as the unsigned default; marked
groups checked against drawn text and bounds; DocMDP, tagged and PDF/A
warnings; `findText` ranges; subset-font and font-dedupe tests.

**Acceptance criteria:**

- [x] A full save drops deleted content and is identical with and without
      prior queries; a moved marked group degrades to plain objects; the three
      warnings fire on their fixtures; ranges round-trip through
      `EditElement.text` offsets; a subset font whose `cmap` maps to an
      emptied glyph is not used in place.

**Verification:**

- [x] `npm run test --workspace web-doc`, `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 26

**Files likely touched:** `src/edit/pdf/engine/document.ts`, `elements.ts`,
`existing-text.ts`, `src/edit/pdf/session.ts`, `types.ts`, PDF tests

**Estimated scope:** Medium

### Task 29: apply() latency measurement

**Description:** One `insertTextBox` on 10-, 100- and 500-page PDFs in
Chromium, with and without a warm PDF.js worker; numbers recorded under the
PDF spec's Actual result.

**Acceptance criteria:**

- [x] The three numbers are in the spec with the machine they were measured on.

**Verification:**

- [x] `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 28

**Files likely touched:** `tests/e2e/edit-pdf.spec.ts`,
`docs/document-editing/todo/02-pdf-edit.md`

**Estimated scope:** Small

### Task 30: Docs, matrix, full gate

**Description:** Final R2 docs pass, `npm run test:e2e:matrix`,
`npm run check`, results recorded in both specs and the roadmap; Linear
ACTION-821 proofs attached.

**Acceptance criteria:**

- [x] Matrix and `npm run check` pass; both specs' R2 definitions of done are
      ticked.

**Verification:**

- [x] `npm run pages:build`, `npm run test:e2e:matrix`, `npm run check`

**Dependencies:** Tasks 25, 27, 29

**Files likely touched:** docs and specs

**Estimated scope:** Small

### Task 31: Compaction hardening (peer review)

**Description:** The full-save compaction pass reads PDFium's output with a
PDF lexer instead of string searches; the corpus, the adversarial fuzz loop
and a 50 MB image-only file prove it; PDFium and PDF.js reopen every output;
the failure behaviour and the show-time save mode are recorded as open
questions for Leonid.

**Acceptance criteria:**

- [x] Hand-written edge cases (keywords inside names, strings and comments;
      split references; indirect lengths defined later; `null` objects) keep
      the reachable objects and drop the rest; the output is idempotent.
- [x] Every corpus PDF edited, saved in full and reopened in PDFium (page
      count, per-page text) and PDF.js 6.2.108 (page count, first-page text).
- [x] `compactPdf` is a fuzz target; a 51.6 MB 200-page raw-image PDF saves
      in full under a second; numbers in the PDF spec, compaction separately.
- [x] Open questions recorded in the PDF spec and the roadmap.

**Verification:**

- [x] `npm run test --workspace web-doc`, `npm run fuzz:js`, `npm run check`

**Dependencies:** Task 28

**Files likely touched:** `src/edit/pdf/engine/compact.ts`,
`test/pdf-edit-compact.test.ts`, `scripts/fuzz-js.mjs`, docs

**Estimated scope:** Small

### Checkpoint F: revision 2 done

- [x] Every R2 item of both definitions of done is checked
- [x] Linear: ACTION-821 Done with proofs; PR opened and ACTION-808/811 moved
      to Code Review with the human

## Phase 6 — PDF overlay primitives (Linear ACTION-825)

Blocked by ACTION-821 (Done 2026-10-01); blocks ACTION-815 (Operators PDF
UI). The API is fixed by `02-pdf-edit.md` → "Overlay primitives"; every
commit carries `[linear:ACTION-825]`. Plumbing shared by the tasks: the
engine client gains one worker request per primitive, the core session
exposes a `read()` hook that queues a read behind earlier calls and stamps
the envelope, and `PdfSession` wraps both. Types live in
`src/edit/pdf/types.ts`.

### Task 32: Text layout, positions and range rectangles

**Description:** `getTextLayout(elementId)` → `ReadItem<TextLayout>`,
`positionAt(pageIndex, point)` → `ReadItem<TextPosition>`,
`rangeRects(range)` → `ReadResult<PageRect>`. A new engine module reads the
text page once per call: characters are mapped to their object and element
(the `findText` char-offset mapping), grouped into lines in reading order
(one per text object; a text box's lines and a table's cells are separate
objects), each line with its baseline (`FPDFText_GetCharOrigin`), glyph boxes
(`FPDFText_GetCharBox`), advances (`FPDFText_GetLooseCharBox`), font family,
size, colour and the `TextRange` it covers. `positionAt` uses
`FPDFText_GetCharIndexAtPos` with a tolerance, then the nearest glyph box on
the page. `rangeRects` unions the glyph boxes per line for the offsets inside
the range, across the elements the range spans in reading order. All
geometry goes through `PageGeometry`, so rotation and crop boxes are
honoured.

**Acceptance criteria:**

- [x] On pages rotated 0/90/180/270 and on a cropped page, every glyph box
      lies inside its element's bounds and the line bounds equal the union of
      its glyph boxes; a text box reports one line per drawn line and a table
      one per cell, in reading order.
- [x] `positionAt` at the centre of a glyph box returns that glyph's offset;
      `rangeRects` of a range contain the points that `positionAt` resolves
      back into the range (round-trip on all four rotations).
- [x] Unknown element, non-text element, a point on an empty page and a
      range across pages give `item: undefined` / `items: []`, not errors.

**Verification:**

- [x] `npm run test --workspace web-doc` (new `test/pdf-edit-layout.test.ts`)

**Dependencies:** Task 28

**Files likely touched:** `src/edit/pdf/types.ts`, `src/edit/pdf/engine/layout.ts`
(new), `document.ts`, `pdfium.ts`, `handler.ts`, `provider.ts`, `session.ts`,
`src/edit/session.ts`, `src/edit/engine.ts`, `src/worker-protocol.ts`

**Estimated scope:** Medium

### Task 33: Suppressed render

**Description:** `renderPageWithout(pageIndex, elementIds, options)` →
`ReadItem<PageBitmap>`: the worker marks the elements' objects inactive
(`FPDFPageObj_SetIsActive`), renders the page with `FPDF_RenderPageBitmap`
into a BGRA bitmap at `options.scale` device pixels per point (default 1,
bounded by the raster limits), converts it to RGBA, restores the objects and
returns the pixels as a transferable buffer with the bitmap's size. Nothing
is reopened and the session's bytes do not change.

**Acceptance criteria:**

- [x] The suppressed render differs from a normal render inside the
      element's bounds and is identical outside; a normal render taken after
      it equals the one taken before (the objects are active again), and a
      save after it equals a save before it.
- [x] Unknown ids are ignored; a scale that exceeds the raster limit is
      refused with `resource-limit`.

**Verification:**

- [x] `npm run test --workspace web-doc` (`test/pdf-edit-render.test.ts`)

**Dependencies:** Task 32 (plumbing)

**Files likely touched:** `src/edit/pdf/engine/render.ts` (new), `pdfium.ts`,
`handler.ts`, `provider.ts`, `session.ts`, `types.ts`, `src/worker-protocol.ts`

**Estimated scope:** Small

### Task 34: Selection mapping and range mapping

**Description:** `elementsForSelection(selection)` → `ReadResult<TextRange>`
maps the viewer's `TextSelection` (PDF.js runs with rectangles in page space)
to elements and ranges through the overlap ladder: a run and a layout line
whose rectangles overlap by at least 50 % of the run; else the single text
object that contains the run's rectangle; else a text match after NFKC
folding and whitespace removal. Offsets inside a matched line come from the
run's characters against the line's glyphs. The ladder is GenOffice's
(Apache-2.0); the attribution goes into `THIRD_PARTY_NOTICES.md` and the
module header. `mapRange(range, fromRevision)` → `ReadItem<TextRange>` walks
the history entries after `fromRevision`: a `replaceText` on the element
shifts offsets after its range by the length difference (or by the whole
text when it has no range), a deleted element or page gives `undefined`,
`remappedIds` are followed, undo and redo are already reflected in the
entries.

**Acceptance criteria:**

- [x] A selection built from the fixture's PDF.js runs resolves to the right
      element and offsets on each ladder rung (overlap, containment, text
      match); a selection of nothing resolves to `items: []`.
- [x] After `replaceText` with a range before, inside and after a saved
      range, `mapRange` returns the moved range; after `deleteElement` it
      returns `undefined`; after `undo` the original range.

**Verification:**

- [x] `npm run test --workspace web-doc` (`test/pdf-edit-selection.test.ts`),
      `npm run licenses`

**Dependencies:** Task 32

**Files likely touched:** `src/edit/pdf/selection.ts` (new), `session.ts`,
`types.ts`, `src/edit/session.ts` (history access), `THIRD_PARTY_NOTICES.md`

**Estimated scope:** Medium

### Task 35: Range-scoped `replaceText`

**Description:** `ReplaceTextOperation.range?: TextRange` inside the target.
A text box splices its stored text and is laid out again. A text object is
rewritten in place with the spliced text when its font covers the new
characters; otherwise the object is split around the range into up to three
objects — the parts before and after keep the font, size, colour and
baseline, the middle part is drawn in the fallback font at the same baseline
and the following part is shifted by the advance difference. The split
object's id stays on the first part; the others are new ids in `createdIds`.
Validation: the range must lie inside the target's text (`invalid-range`),
and tables stay unsupported.

**Acceptance criteria:**

- [x] Replacing a word in the middle of a Helvetica text object keeps one
      object and the other words' glyph boxes; replacing it with Cyrillic
      produces three objects on one baseline whose texts concatenate to the
      expected string, with the tail moved by the advance difference.
- [x] A text box with a range reflows; a range outside the text or on a
      table is refused with the stated codes; the operation's JSON schema
      accepts `range`.

**Verification:**

- [x] `npm run test --workspace web-doc` (`test/pdf-edit-replace-range.test.ts`), `npm run test:e2e -- tests/e2e/edit-pdf.spec.ts`

**Dependencies:** Task 32

**Files likely touched:** `src/edit/pdf/engine/existing-text.ts`, `text-box.ts`,
`src/edit/pdf/types.ts`, `schema` of the operation, docs

**Estimated scope:** Medium

### Task 36: Main-thread geometry cache

**Description:** `PdfSession` keeps the last `getElements({ pageIndex })`
result per page and exposes `elementsAtSync(pageIndex, point)` →
`ReadResult<PdfElement>` read synchronously from it (empty with the cached
revision when the page has not been read; `cachedPages` says which have).
On every receipt the cache drops the changed pages and refetches the ones
that were cached, so hover never waits behind an `apply()`; `end()` clears
it.

**Acceptance criteria:**

- [x] After `getElements({ pageIndex: 0 })`, `elementsAtSync(0, point)`
      returns the same elements as `elementsAt` without awaiting; while an
      `apply()` is queued it still answers from the previous revision and
      reports that revision; after the receipt the changed page is refreshed.

**Verification:**

- [x] `npm run test --workspace web-doc` (`test/pdf-edit-session.test.ts`)

**Dependencies:** Task 32

**Files likely touched:** `src/edit/pdf/session.ts`, `types.ts`,
`src/edit/session.ts` (receipt hook)

**Estimated scope:** Small

### Task 37: Browser test, docs, matrix, proofs

**Description:** A Playwright test selects text in the viewer with the
mouse, resolves it with `elementsForSelection`, compares the layout's glyph
boxes with the PDF.js text layer within 1 CSS px on rotated and cropped
pages, replaces the range, maps it with `mapRange` after `documentchange`
and restores the selection with `selectText`; a suppressed render is drawn
over the page and the next normal render shows the element again.
`docs/api/editing.md` gets an "Overlay primitives" section with the
interaction model (overlay input → commit on blur or idle → re-selection);
the PDF spec's Actual result records Phase 6; matrix and `npm run check`
pass; proofs attached and ACTION-825 Done.

**Acceptance criteria:**

- [x] The browser test passes on the matrix; the docs describe every
      primitive with its envelope and the interaction model; the spec's
      overlay definition of done is ticked.

**Verification:**

- [x] `npm run test:e2e -- tests/e2e/edit-pdf-overlay.spec.ts`,
      `npm run test:e2e:matrix`, `npm run check`

**Dependencies:** Tasks 33–36

**Files likely touched:** `tests/e2e/edit-pdf-overlay.spec.ts` (new), docs, specs

**Estimated scope:** Medium

### Checkpoint G: overlay primitives done

- [x] ACTION-825 acceptance criteria met; Linear Done with proofs
- [ ] ACTION-815 (Operators) unblocked; parent ACTION-723 reviewed with
      Leonid for the PR / Code Review step

## Phase 7 — OOXML package layer (Linear ACTION-810)

Blocked by ACTION-808 (Done); blocks ACTION-812 (pptx-edit) and ACTION-813
(docx-engine-upgrade). The API, error codes and decisions are fixed by
`03-ooxml-package.md` (approved 2026-10-02: `store`, `warn`); every commit
carries `[linear:ACTION-810]`. Code lives in `packages/viewer/src/edit/ooxml/`.

### Task 38: ZIP reader

**Description:** `parseZip(bytes, limits)` reads the end-of-central-directory
record, the central directory and, on first use, each entry's local header;
`inflateEntry` streams an entry through `DecompressionStream("deflate-raw")`
with metering and a CRC-32 check; `OoxmlPackage.open` wraps it with OPC name
normalization, case-insensitive lookup and a part cache. ZIP64, encryption,
other methods and multi-disk archives are refused with `unsupported-package`;
lying directories, bad CRCs and missing records are `invalid-file`; the
existing limits apply. The four new error codes join `ViewerErrorCode`. Spike
recorded: compression streams in Node 22 and the matrix; corpus facts.

**Acceptance criteria:**

- [x] Hand-built archives (stored, deflated, data descriptors with and
      without signature, extra fields, archive comment) read back exactly;
      every refusal and every error code has a test.
- [x] Every corpus and fixture PPTX/DOCX opens, every entry inflates with a
      matching CRC, `[Content_Types].xml` is found; open time of the corpus
      deck is recorded.
- [x] `parseZip` is a fuzz target.

**Verification:**

- [x] `npm run test --workspace web-doc` (`test/ooxml-zip.test.ts`),
      `npm run fuzz:js`

**Dependencies:** none

**Files likely touched:** `src/edit/ooxml/zip.ts`, `names.ts`, `package.ts`
(new), `src/contracts.ts`, `test/fixtures/zip-builder.ts`, `scripts/fuzz-js.mjs`

**Estimated scope:** Medium

### Task 39: ZIP writer

**Description:** `writeZip` copies untouched entries verbatim (local record
bytes through the data descriptor, central record with the offset
rewritten), writes changed entries stored (or deflated on request) with fresh
headers, appends new entries with the fixed 1980 time, drops removed ones,
keeps the comment, refuses ZIP64 output. `OoxmlPackage.save` returns the
original bytes without changes.

**Acceptance criteria:**

- [x] No-change save and rebuild-with-no-changes are byte-identical to the
      original for every corpus and fixture package; after one changed part
      only that entry's bytes differ, entry by entry.
- [x] Stored and deflated outputs reopen through the reader with matching
      CRCs; an output that would need ZIP64 is refused before writing.

**Verification:**

- [x] `npm run test --workspace web-doc` (`test/ooxml-zip.test.ts` additions)

**Dependencies:** Task 38

**Files likely touched:** `src/edit/ooxml/zip.ts`, `package.ts`

**Estimated scope:** Medium

### Task 40: OPC model

**Description:** `[Content_Types].xml` (`Default`, `Override`) and `.rels`
parts as small models written back as patches; `contentTypeOf`,
`relationships`, `resolve`, `rId` allocation, content-type rules on add and
remove, indexed lookups by id and type.

**Acceptance criteria:**

- [ ] Target resolution covers relative, `..`, absolute, external and
      case-different targets; `rId` allocation fills the smallest gap; adding
      a part adds an Override only when no Default covers it; an untouched
      `[Content_Types].xml` keeps its bytes.

**Verification:**

- [ ] `npm run test --workspace web-doc` (`test/ooxml-opc.test.ts`)

**Dependencies:** Task 41 (patches are written through the scanner)

**Files likely touched:** `src/edit/ooxml/opc.ts` (new), `package.ts`

**Estimated scope:** Medium

### Task 41: XML scanner

**Description:** The offset-preserving tokenizer: declaration, opaque
regions, start and end tags, attributes, self-closing tags, character data;
element tree with exact ranges, namespaces in scope, `find`/`findAll`/`at`/
`textOf`/`attribute`; the UTF-8 round-trip guard; `malformed-xml`; fuzz
target.

**Acceptance criteria:**

- [ ] Every construct in the spec scans with correct ranges on hand-written
      parts and on every XML part of the corpus; `text.slice(start, end)` of
      every element reproduces its source; a part that fails the UTF-8 guard
      is `unsupported-part`; malformed input is `malformed-xml`.
- [ ] The largest corpus part scans in the time recorded under Actual result.

**Verification:**

- [ ] `npm run test --workspace web-doc` (`test/ooxml-xml.test.ts`),
      `npm run fuzz:js`

**Dependencies:** Task 38

**Files likely touched:** `src/edit/ooxml/xml.ts` (new), `package.ts`

**Estimated scope:** Medium

### Task 42: Patches and transactions

**Description:** Patch builders with expectations, apply-from-the-end,
re-scan and read-back verification, atomic `PackageTransaction` across parts
(patch, setPart, removePart, relationships, addMedia with hash
deduplication, uniquePartName, dangling-target warnings), O(1) snapshots and
restore.

**Acceptance criteria:**

- [ ] Every builder produces a patch whose read-back passes; overlapping,
      stale, malformed and mismatching patches are `invalid-patch` and leave
      the package unchanged, also when they are the last part of a
      multi-part transaction; snapshots restore exactly; media is stored once
      per content; removing a part leaves a warning per dangling relationship.

**Verification:**

- [ ] `npm run test --workspace web-doc` (`test/ooxml-patch.test.ts`)

**Dependencies:** Tasks 39, 40, 41

**Files likely touched:** `src/edit/ooxml/patch.ts`, `transaction.ts` (new),
`package.ts`

**Estimated scope:** Medium

### Task 43: Corpus, browser reopen, docs, gate

**Description:** Every corpus and fixture package through the full cycle
(open → no-change save identical → one-part patch → only that entry differs
→ every part well-formed → reopened by `@silurus/ooxml` in the browser);
`vbaProject.bin` survival on a macro-enabled fixture; adversarial manifest;
performance table; spec Actual result and Spike results; architecture doc
paragraph; `npm run check` and the matrix; Linear proofs; ACTION-810 Done.

**Acceptance criteria:**

- [ ] The definition of done of `03-ooxml-package.md` is met in full.

**Verification:**

- [ ] `npm run test:e2e -- tests/e2e/ooxml-package.spec.ts`,
      `npm run test:e2e:matrix`, `npm run check`

**Dependencies:** Task 42

**Files likely touched:** `tests/e2e/ooxml-package.spec.ts` (new), docs

**Estimated scope:** Medium

### Checkpoint H: package layer done

- [ ] ACTION-810 Done with proofs; ACTION-812 and ACTION-813 unblocked
