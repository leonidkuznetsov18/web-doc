# Roadmap: document editing API

> **Status, 2026-10-01:** the capability map below is approved, and so are the
> specs for `edit-core` and `pdf-edit`. Module specs are written one at a time,
> right before a module starts, so each one can use what the previous modules
> taught us.

## Purpose of this document

This is the index of the `document-editing` feature package. It records the
objective, the decisions that every module must respect, the capability map
with its build order, and the project-wide rules (commands, structure, code
style, testing, boundaries). Each module has its own spec in
[`todo/`](./todo/), named after its module id.

## Requirement sources

- Planning session of 2026-10-01: answers about the users of the API, the MVP
  scope, the PDF engine, the format order and the spec language. They are
  captured below as fixed decisions.
- Competitor and library research of 2026-10-01: GenOffice (Apache-2.0),
  SuperDoc (AGPL shell plus a proprietary engine), Docmentis (closed engine),
  PDF.js 6.2.108, `@silurus/ooxml` and PDFium WASM. The research report stays
  outside this public repository by decision of 2026-10-01; the specs carry the
  findings they rely on.
- Existing viewer contracts: [`../api/reference.md`](../api/reference.md),
  [`../architecture.md`](../architecture.md) and
  [`../universal-document-viewer/00-roadmap.md`](../universal-document-viewer/00-roadmap.md),
  which lists editing as deferred work that this package now picks up.

## Objective

web-doc gets a programmatic editing API for **PDF, PPTX and DOCX**. The host
application builds its own editing UI on top of it, and the host's AI agent uses
exactly the same API. Every edit is applied to the original file and drawn by the
engine that already renders the viewer, so what the user sees is what gets saved.

Who uses it:

- **Integrator** — the front-end developer of the host application. Builds the
  toolbars, inspectors and drag handles. Needs methods, element inspection with
  geometry, events, undo/redo and saving.
- **AI agent** of the host application. Reads the document structure, finds
  targets by text, sends validated batches of JSON operations, checks a batch
  with a dry run and is told when its view of the document is stale.
- **End user** of the host application. Expects an edit to appear at once and to
  survive saving and reopening in Word, PowerPoint or a PDF reader.

User stories for the MVP:

1. As an integrator, I start an edit session on a loaded PDF, PPTX or DOCX and
   call typed methods (insert text, change formatting, alignment and colours,
   insert a table, image or shape, move, resize or delete elements, work with
   pages or slides), and the viewer shows the result.
2. As an integrator, I list the elements of a page with their ids, kinds, bounds,
   text and style, and I ask which element lies under a point, so my UI can
   select and highlight it.
3. As an integrator, I undo and redo changes and save the edited file as bytes;
   saving without changes returns the original bytes.
4. As an AI agent (through the host), I send a batch of JSON operations with
   `expectedRevision`, receive either a typed validation error or a receipt, and
   can use `dryRun` to check a batch before applying it.

## Fixed decisions

1. **API only.** web-doc ships no editing UI: no toolbars, dialogs, caret or
   text-input overlays. It ships methods, element inspection, events and
   view-geometry helpers that let the host draw its own UI over the viewer.
2. **One operation vocabulary for the UI, the public API and AI.** Every typed
   method is a single-operation `apply()`. Operations are plain JSON objects with
   exported JSON Schemas, so an AI tool call and a button click take the same
   path.
3. **The original file is the source of truth.** The edit state is the original
   bytes plus an operation history. Undo and redo rebuild that state; saving
   without changes returns identical bytes; an edit rewrites only what it
   touches.
4. **Rendering stays with the existing engines.** After a change the viewer
   reopens the edited bytes through its normal adapter (PDF.js for PDF,
   `@silurus/ooxml` for Office), so the canvas always shows the file that
   `save()` would return.
5. **PDF changes are written by PDFium WASM** (`@embedpdf/pdfium` 2.15.1, an MIT
   wrapper around BSD-3-Clause PDFium). It is loaded lazily into a dedicated
   worker only when editing starts. PDF.js stays the PDF renderer.
6. **PPTX and DOCX are edited XML-first.** The original package is the source of
   truth, edits are patches to its XML parts, and untouched parts are copied as
   they are.
7. **Formats in the MVP:** `pdf`, `pptx`, `docx`. The `pptm`, `ppsx` and `docm`
   variants follow once the base formats pass; macro parts are copied byte for
   byte and never executed.
8. **Out of scope** for this package: a built-in editing UI; Markdown and plain
   text; spreadsheets (XLSX, CSV); legacy DOC, XLS and PPT (they stay view-only);
   real-time collaboration (operations stay serializable so it can be added
   later); password-protected files; printing; any server-side processing.
9. **Order:** PDF first, then PPTX, then DOCX; AI-specific tooling last.

## Capability map

| #   | Module id             | Responsibility                                                                                                                                                                                                                            | Depends on                             | Spec                                             | Status        |
| --- | --------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------- | ------------------------------------------------ | ------------- |
| 01  | `edit-core`           | Format-independent contract: edit session, JSON operations with schemas, dry run, revisions, undo/redo, `save()`, events, element inspection with geometry, view-geometry helpers, viewer refresh after a change, round-trip test harness | —                                      | [`todo/01-edit-core.md`](./todo/01-edit-core.md) | Spec approved |
| 02  | `pdf-edit`            | PDF methods on PDFium: text boxes (font, size, colour, alignment), replace and restyle existing text, images, shapes, simple tables, move/resize/delete objects, page operations; incremental save                                        | `edit-core`                            | [`todo/02-pdf-edit.md`](./todo/02-pdf-edit.md)   | Spec approved |
| 03  | `ooxml-package`       | Shared OOXML layer: unzip once, copy untouched ZIP entries as they are, offset-preserving XML scanner, patches with self-verification, relationships and content types                                                                    | `edit-core`                            | Written before the module starts                 | Not started   |
| 04  | `pptx-edit`           | PPTX methods: shape text, run formatting, paragraph alignment, text and fill colours, move/resize/delete shapes, insert text box, image and table, add/duplicate/delete/move slides                                                       | `ooxml-package`                        | Written before the module starts                 | Not started   |
| 05  | `docx-engine-upgrade` | Move DOCX rendering from `@silurus/ooxml` 0.72.2 to the 0.88 line (one engine copy), with image fitting done as an XML pre-pass; brings `w14:paraId` into text runs for mapping selections and geometry to paragraphs                     | `ooxml-package`                        | Written before the module starts                 | Not started   |
| 06  | `docx-edit`           | DOCX methods: paragraph text, run formatting, alignment and spacing, text colour and highlight, insert/delete/move paragraphs, tables, images                                                                                             | `docx-engine-upgrade`                  | Written before the module starts                 | Not started   |
| 07  | `ai-edit`             | AI tooling over the same operations: document outline, target resolution from text or citations, tool schemas, suggestion mode (tracked changes where the format has them), checkpoints                                                   | `edit-core`, plus formats as they land | Written before the module starts                 | Not started   |

Build order: `edit-core` → `pdf-edit` → `ooxml-package` → `pptx-edit` →
`docx-engine-upgrade` → `docx-edit` → `ai-edit`.

A module merges to `main` only when its definition of done is met, because
every push to `main` publishes a release.

## MVP method set by category

The request named these categories: text, formatting, alignment, colours,
tables, insertion and moving elements. The PDF column is fixed by
[`todo/02-pdf-edit.md`](./todo/02-pdf-edit.md); the PPTX and DOCX columns are
targets that their own specs will make exact.

| Category   | PDF (`pdf-edit`)                                               | PPTX (`pptx-edit`, target)                  | DOCX (`docx-edit`, target)                  |
| ---------- | -------------------------------------------------------------- | ------------------------------------------- | ------------------------------------------- |
| Text       | Insert a text box; replace the text of an existing text object | Replace or insert text in a shape           | Replace, insert or delete paragraph text    |
| Formatting | Font, size, bold, italic in text boxes; size of existing text  | Bold, italic, underline, size, font of runs | Bold, italic, underline, size, font of runs |
| Alignment  | Left, centre, right in text boxes                              | Paragraph alignment                         | Paragraph alignment and spacing             |
| Colours    | Text colour; shape stroke and fill                             | Text colour; shape fill and line            | Text colour; highlight                      |
| Tables     | Draw a simple table; edit its cell text                        | Insert a table; edit cell text              | Insert a table; edit cell text              |
| Insertion  | Image, shape, blank page                                       | Text box, image, slide                      | Paragraph, image, table                     |
| Moving     | Move, resize, delete objects; move, rotate, delete pages       | Move, resize, delete shapes; reorder slides | Move paragraphs                             |

## Tech stack

- TypeScript 7.0.2 with `strict`, `noUncheckedIndexedAccess`,
  `exactOptionalPropertyTypes`, `NodeNext` modules; ES2022 ESM; no UI framework.
- Build: `tsc` plus esbuild 0.28.1 through `scripts/build-viewer.mjs`; workers
  are bundled into `dist/workers/`, WASM and other assets are copied into
  `dist/assets/`.
- Renderers (unchanged): `pdfjs-dist` 6.2.108 for PDF; `@silurus/ooxml` 0.72.2
  for DOCX and 0.88.0 (npm alias `@silurus/ooxml-pptx`) for PPTX.
- New runtime dependency, approved on 2026-10-01: `@embedpdf/pdfium` 2.15.1,
  pinned exactly (MIT wrapper; PDFium itself is BSD-3-Clause). WASM is
  4,646,932 bytes raw, about 1.65 MB Brotli.
- Tests: `node:test` for unit tests (compiled with `tsconfig.test.json`),
  Playwright 1.61.1 for browser tests through the vanilla example, plus the
  existing visual, regression, license and size gates.
- Rust/WASM crates stay unchanged unless a module spec says otherwise.

## Commands

```bash
npm ci                                           # install (also sets the commit-msg hook)
npm run build                                    # WASM bindings, viewer dist and examples
npm run typecheck                                # all workspaces
npm run test --workspace web-doc                 # viewer unit tests only (fast loop)
npm test                                         # all unit tests plus cargo test
npm run test:e2e -- tests/e2e/edit-core.spec.ts  # one browser suite (Chromium); builds first
npm run test:e2e:matrix                          # Chromium, Firefox and WebKit, before a module is done
npm run format                                   # Prettier and rustfmt
npm run licenses                                 # runtime license gate
npm run report:size                              # size report against the 20 MiB Brotli target
npm run check                                    # full gate before merging
```

## Project structure

```text
packages/viewer/src/edit/             edit-core: session, history, validation, schemas, element and geometry types
packages/viewer/src/edit/pdf/         pdf-edit: PDF session methods, operation schemas, worker client
packages/viewer/src/pdf-edit-worker.ts PDFium worker entry, bundled to dist/workers/pdf-edit-worker.js
packages/viewer/src/edit/ooxml/       ooxml-package (later)
packages/viewer/src/edit/pptx/        pptx-edit (later)
packages/viewer/src/edit/docx/        docx-edit (later)
packages/viewer/test/edit-*.test.ts   unit tests
tests/e2e/edit-*.spec.ts              browser tests
tests/fixtures/edit/                  small generated or self-authored fixtures with recorded provenance
docs/api/editing.md                   public guide to the editing API
docs/document-editing/                this feature package: roadmap and module specs
```

The exact file list of a module is fixed in its plan, not here.

## Code style

The editing code follows the conventions of the existing viewer code:

```ts
export class EditHistory<TOperation extends EditOperation> {
  readonly #entries: HistoryEntry<TOperation>[] = [];
  #position = 0;

  get canUndo(): boolean {
    return this.#position > 0;
  }

  push(entry: HistoryEntry<TOperation>): void {
    // A new change after an undo drops the redo tail.
    this.#entries.length = this.#position;
    this.#entries.push(Object.freeze(entry));
    this.#position = this.#entries.length;
  }
}

export function assertExpectedRevision(
  expected: number | undefined,
  actual: number,
): void {
  if (expected !== undefined && expected !== actual)
    throw new ViewerError(
      "edit-conflict",
      "The document changed since it was read",
      {
        details: { expectedRevision: expected, revision: actual },
      },
    );
}
```

- ES private fields (`#name`), `readonly` by default; results handed to callers
  are frozen (`Object.freeze`, or `deepFreeze` for nested data).
- Optional properties are spread conditionally
  (`...(value === undefined ? {} : { value })`), as
  `exactOptionalPropertyTypes` requires.
- Failures are `ViewerError` instances with a stable code and structured
  `details`; nothing throws strings; unsupported input fails closed.
- Every asynchronous operation accepts an `AbortSignal` and respects
  `maxOperationMs`.
- Importing the package and creating a session stay SSR-safe: no DOM access
  unless a container is mounted.
- Comments explain why, not what, in the same density as the surrounding code.
- Prettier 3.9.5 formatting; Conventional Commits with a required scope
  (`viewer`, `docs`, `deps`, …) checked by commitlint.

## Testing strategy

1. **Unit tests** (`node:test`): operation validation and schemas, history,
   revision and conflict rules, receipts, events, coordinate conversion, OOXML
   patching. Engine logic that does not need a DOM also runs here; PDFium WASM
   runs in Node, so PDF operations can be checked without a browser.
2. **Browser tests** (Playwright; Chromium in the development loop, the full
   matrix before a module is done): each public method end to end — load a
   fixture, call the method, check that the viewer re-rendered, save, load the
   saved bytes into a fresh viewer, check text, geometry and page count. Also
   asset loading, laziness and worker failures.
3. **Round-trip invariants**, checked at both levels: `save()` without changes
   returns identical bytes; undoing back to revision 0 returns identical bytes;
   rebuilding the same history twice returns identical bytes.
4. **Independent read-back:** PDF output is reopened by PDF.js and by PDFium;
   OOXML output is reopened by `@silurus/ooxml` and checked for well-formed XML.
5. **Regression:** the existing unit, browser, visual and fuzz suites stay
   green; the license gate and the size report pass.

Fixtures are generated inside the tests or are small self-authored files in
`tests/fixtures/edit/` with their provenance recorded. Public corpus files
(`npm run corpus:fetch`) are used for robustness. Private or customer documents
are never added.

Expected coverage: every public method has at least one unit test and one
browser round-trip test, and every error code has a test that triggers it.

## Boundaries

**Always**

- Keep the original bytes immutable and prove the round-trip invariants with
  tests.
- Validate every operation before applying it; apply a batch completely or not
  at all.
- Load editing code and assets lazily; viewing a document without calling
  `edit()` must not fetch them.
- Add unit and browser tests with every method, and update
  `docs/api/editing.md` in the same change.
- Run `npm run typecheck`, the viewer unit tests and the affected browser suite
  before each commit; run `npm run check` before a module is marked done.
- Preserve content the editor does not understand.

**Ask first**

- Adding, upgrading or replacing a runtime dependency beyond the approved
  `@embedpdf/pdfium` 2.15.1.
- Any non-additive change to the existing public API or its semantics.
- Changing CI, release, license-gate or size-gate configuration.
- Adding bundled font or other large assets beyond the approved PDF fallback
  font (a TrueType build of the bundled Noto Sans Latin/Cyrillic face).
- Changing the pinned `@silurus/ooxml` versions (the `docx-engine-upgrade`
  module will ask with its own spec).
- Committing the research report or notes to this public repository.

**Never**

- Copy code from SuperDoc or Docmentis, or install `superdoc` 2.x or
  `@superdoc/docx-engine` anywhere in this repository, its CI or development
  setups. Code ideas from GenOffice may be reused only with Apache-2.0
  attribution.
- Add GPL, AGPL, LGPL, MPL or proprietary ("Pro", `LicenseRef-*`) runtime
  dependencies.
- Execute macros, call out to servers, or add telemetry.
- Commit private documents or customer data.
- Weaken a test or a gate to make a build pass (skip or delete tests, loosen
  thresholds, add `@ts-ignore`).
- Build editing UI, or add Markdown or plain-text support, as part of this
  package.

## Success criteria for the MVP

- For PDF, PPTX and DOCX, every method in the module's MVP list works in
  Chromium through the public API, is visible in the viewer after the call,
  survives save and reload, and has unit and browser tests.
- For every fixture, `save()` without changes returns identical bytes, and
  undoing to revision 0 returns identical bytes.
- A JSON-only client — no functions or class instances — can drive every MVP
  method through `apply()` using the exported schemas, and gets typed errors for
  invalid or stale batches.
- The existing suites, the license gate and the size budget pass. Loading and
  viewing a document without calling `edit()` fetches no editing assets.
- `docs/api/editing.md` documents every method, element kind, error code and
  event.

## Decisions log

| Date       | Decision                                                                                                   |
| ---------- | ---------------------------------------------------------------------------------------------------------- |
| 2026-10-01 | API only, shared by the host UI and AI; formats PDF, PPTX, DOCX; Markdown and plain text dropped entirely. |
| 2026-10-01 | PDF changes are written by PDFium WASM (`@embedpdf/pdfium` 2.15.1); PDF.js keeps rendering.                |
| 2026-10-01 | Order: PDF, then PPTX, then DOCX; specs are written module by module.                                      |
| 2026-10-01 | Specs are written in English.                                                                              |
| 2026-10-01 | `edit-core` and `pdf-edit` specs approved together with their recommended answers to open questions.       |
| 2026-10-01 | A TrueType build of the bundled Noto Sans Latin/Cyrillic face ships as a lazy PDF fallback font.           |
| 2026-10-01 | Signed PDFs may be edited, with a warning.                                                                 |
| 2026-10-01 | The competitor research report stays outside this public repository.                                       |

## Open questions

None at the moment. New questions are recorded in the spec of the module that
raises them.
