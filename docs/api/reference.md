# Public API reference

`web-doc` is an ESM, browser-side document viewer. Importing the package and creating a headless viewer are SSR-safe; a DOM is required only when `container` is supplied or `downloadOriginal()` initiates a browser download.

All page, slide, image, and sheet indices are zero-based. API snapshots and result arrays are immutable.

## Create a runtime and viewer

```ts
import { ViewerClient } from "web-doc";

const client = ViewerClient.create({
  assetBaseUrl: new URL("/zrimo/", location.href),
  limits: { maxInputBytes: 50 * 1024 * 1024 },
});

const viewer = client.createViewer({
  container: document.querySelector<HTMLElement>("#viewer")!,
  locale: "ru",
  layout: "continuous",
  fit: "width",
  overscan: 1,
});

await viewer.load(file, { fileName: file.name });
```

`ViewerClient.create()` and `createViewer()` are synchronous. Parsing and rendering assets are initialized lazily by `load()`/render operations. A client may own multiple independent viewers.

### `ViewerClientOptions`

| Option         | Type                         | Meaning                                                                                        |
| -------------- | ---------------------------- | ---------------------------------------------------------------------------------------------- |
| `assetBaseUrl` | `string \| URL`              | Base URL for package-owned workers, WASM, and future font assets.                              |
| `fetch`        | `ViewerFetch`                | Host-controlled fetch implementation used only for explicit URL sources.                       |
| `logger`       | `ViewerLogger`               | Optional debug/warning/error sink.                                                             |
| `limits`       | `Partial<ResourceLimits>`    | Runtime defaults for input/archive/pixel/SVG/CSV/text/unit/concurrency/operation-time budgets. |
| `adapters`     | `readonly DocumentAdapter[]` | Replace the built-in adapter set, primarily for testing or custom formats.                     |
| `fontPolicy`   | `FontPolicy`                 | Select `auto`, `offline`, or host `custom` font resolution.                                    |
| `fonts`        | `readonly RegisteredFont[]`  | App-provided URL/byte faces resolved before system/package fallback.                           |

`client.registerAdapter(adapter)` registers an additional adapter and returns an unregister function. `await client.destroy()` destroys every owned viewer and shared adapter.

### `ViewerOptions`

| Option         | Default         | Meaning                                                                                                       |
| -------------- | --------------- | ------------------------------------------------------------------------------------------------------------- |
| `container`    | none            | `HTMLElement` receiving the managed viewport; omit for headless use.                                          |
| `layout`       | `"continuous"`  | Continuous virtual list or one visible unit with `"single"`.                                                  |
| `overscan`     | `1`             | Extra units on either side of the visible range, clamped to `0…5`.                                            |
| `initialZoom`  | `1`             | Initial scale, clamped to `0.1…8`.                                                                            |
| `fit`          | `"none"`        | Initial `"none"`, `"width"`, or `"page"` fit mode.                                                            |
| `locale`       | `"en"`          | Built-in locale identifier (`"en"` or `"ru"`).                                                                |
| `ui`           | `false`         | Mount the optional localized toolbar/panels around the viewport.                                              |
| `useWorker`    | adapter default | Compatibility switch for integrations; built-in parsing backends select their safe worker path.               |
| `translations` | none            | Partial locale dictionary override with English fallback.                                                     |
| `search`       | none            | Defaults for every `search()` call, currently `{ fuzzy }`; see [Search and selection](#search-and-selection). |

The host must give the container a resolvable height. The viewport fills its container and removes all DOM/listeners during `destroy()`.

## Lifecycle and state

`load(source, options?)` accepts a URL string, `URL`, `Blob`/`File`, `ArrayBuffer`, or `Uint8Array`. `open()` is an alias. A new load cancels the previous load and closes the old handle. `close()` returns the viewer to `idle`; `destroy()` is idempotent and terminal.

```ts
interface ViewerState {
  readonly status: "idle" | "loading" | "ready" | "error" | "destroyed";
  readonly format?: DocumentFormat;
  readonly pageIndex: number;
  readonly pageCount: number;
  readonly zoom: number;
  readonly fit: "none" | "page" | "width";
  readonly panX: number;
  readonly panY: number;
}
```

`viewer.state` is replaced, not mutated. `getDocumentInfo()` returns format,
render unit, count, optional sheet geometry/warnings, and capabilities
(`textSelection`, `cellSelection`, `search`, and `thumbnails`). Page-oriented
backends may also return `pageSizes`, an ordered `{ width, height }[]` in natural
CSS pixels at zoom 1; the managed viewport uses it for mixed-size PDF and Office
documents instead of coercing every page to A4. The image adapter reports the
decoded bitmap size (TIFF per page) and the SVG adapter the root
`width`/`height` or `viewBox`, so fit-to-width and fit-to-page scale a picture
against its own geometry.

An inline DOCX picture declared wider or taller than its section's text area
(a generated document often embeds a chart at its native size) is scaled down
to fit before the page layout runs, aspect ratio preserved; Word would draw it
clipped at the page edge. Anchored (floating) pictures keep their geometry.

## View and navigation

| Method                     | Behavior                                                                            |
| -------------------------- | ----------------------------------------------------------------------------------- |
| `setZoom(zoom)`            | Set a clamped absolute zoom and leave fit mode.                                     |
| `zoomIn()` / `zoomOut()`   | Multiply or divide zoom by `1.2`.                                                   |
| `setFit(mode)`             | Apply `none`, `width`, or `page`.                                                   |
| `fitWidth()` / `fitPage()` | Derive scale from the attached viewport; headless viewers keep their current scale. |
| `panBy(dx, dy)`            | Scroll the attached viewport or update headless pan state.                          |
| `goToPage(index)`          | Navigate to a clamped page/slide/image/sheet index.                                 |
| `next()` / `previous()`    | Navigate one unit.                                                                  |
| `setSheet(index)`          | Navigate a sheet document; throws `lifecycle-error` for non-sheet documents.        |

The managed viewport also supports pointer drag, Ctrl/Cmd+wheel, two-pointer
pinch, Page Up/Down, and arrow-key pan. Spreadsheet documents automatically use
a virtual surface with variable/hidden row and column geometry, fixed
headers/frozen panes, one bounded canvas, and enough trailing blank bands to
fill the viewport. Column borders in the header can be dragged to apply
view-only per-sheet width overrides without modifying the source file. Sheet
`fitWidth` fits the used columns, `fitPage` fits both axes, and zoom preserves
the pointer or top-left viewport anchor. Spreadsheet cells support pointer drag,
Shift+click/Shift+arrow range extension, Ctrl/Cmd+click or drag for
non-contiguous ranges, and Ctrl/Cmd+C copying of the current selection.

The optional controls, shortcuts, localization, and CSS variables are documented in [Basic UI](../ui.md). Font policies and resolver types are documented in [Fonts](../fonts.md).

## Search and selection

`search(query, { caseSensitive?, pageRange?, nearPage?, fuzzy? })` searches logical text in page order. The default applies Unicode NFKC normalization and locale-independent case folding while mapping matches back to original UTF-16 offsets. It does not strip Arabic diacritics or reorder RTL text.

`pageRange` is an inclusive 0-based `[first, last]` window. It bounds both the pages whose text is read and the matches the result carries, so a caller that already knows where a passage lives — a citation with a page number, for example — highlights only that page instead of every incidental match in the document, and reads one page of text instead of all of them. Endpoint order does not matter; an out-of-range endpoint rejects with `render-failed` and leaves the standing result and its highlights untouched.

`nearPage` is an approximate 0-based page hint for callers whose page numbers come from another pagination of the same file (a citation produced from a server-side render, while the viewer paginates DOCX itself). It is clamped into the scanned window rather than rejected. The result's `activeIndex` — and the page the viewer navigates to — becomes the match closest to the hint, and the fuzzy fallback scans pages nearest to it first.

`fuzzy` enables a fallback that runs only when the exact search finds nothing. [Fuse.js](https://www.fusejs.io/) selects candidate pages and a whole-query edit alignment locates one contiguous passage: the query is compared to each page's text with a bounded edit budget, so spacing, line breaks, list bullets, table separators and typographic punctuation may differ from the source, and every hit is mapped back to the verbatim page text so highlights land on the original. Pages are compared nearest to `nearPage` first, a batch at a time, and the scan stops at the first batch that holds the passage. `true` uses the viewer's defaults (`ViewerOptions.search.fuzzy`, so an integration can turn it on once at `createViewer`), an object enables it and overrides them, `false` disables it for one call:

| Option              | Default | Meaning                                                                                                               |
| ------------------- | ------- | --------------------------------------------------------------------------------------------------------------------- |
| `threshold`         | `0.3`   | Fuse.js edit budget per 32-character chunk of the query, `0` exact to `1` anything.                                   |
| `maxScore`          | `0.4`   | Highest Fuse.js score and whole-passage edit ratio allowed (`0` perfect); raise it to accept a passage that spans a page break.   |
| `maxQueryLength`    | `600`   | Query characters considered; the matcher's cost grows with the query and a passage is identified well before its end. |
| `maxPageTextLength` | `20000` | Characters of each page's text considered.                                                                            |
| `pagesPerBatch`     | `2`     | Main-thread fallback only: pages compared per batch, yielding to the event loop between batches.                      |
| `pageWindow`        | `12`    | With `nearPage`, how many pages nearest to the hint are scanned before giving up; without a hint every page is.       |
| `worker`            | `true`  | Match in a Web Worker that keeps the document's Fuse.js index; falls back to the main thread when unavailable.        |

With `worker` on, the first fuzzy search of a document ships its page texts to `workers/fuzzy-search-worker.js` (served next to the other worker assets), which builds one Fuse.js index and answers every later search off the main thread — a citation lookup tries several anchors in a row, and a long document never blocks the page while they run. A worker that cannot start (no `Worker`, a missing asset) is logged through `ViewerLogger.warn` and the main-thread matcher takes over for the rest of the document's life.

`SearchResult.strategy` reports how the matches were found (`exact` or `fuzzy`) and is absent when there are none.

A search, `searchNext()` and `searchPrevious()` land on the active match itself, not only on its page: once the page's text runs are known the viewport scrolls so the match sits about a third of the way down, so a hit low on a page taller than the viewport is visible without a manual scroll. A fuzzy result carries one match per page, covering the best contiguous occurrence. Fuse character masks are not used as offsets: repeated occurrences and unrelated chunks cannot stretch the range across a page. Equal-cost occurrences prefer the earliest end. The `maxScore` ceiling also bounds whole-passage edit distance divided by normalized query length; a candidate that fails this check produces no match, allowing callers to try a shorter citation anchor.

```ts
const result = await viewer.search("привет");
viewer.searchNext();
viewer.searchPrevious();
viewer.clearSearch();
```

`SearchResult` contains the original query, immutable `{ pageIndex, start, end, text }` matches, `activeIndex` (`-1` when empty), and `strategy`. Starting a new search cancels an earlier one.

Use `selectText({ startPageIndex, startOffset, endPageIndex, endOffset })` for logical cross-page ranges. Use `selectCells({ sheetIndex, startRow, startColumn, endRow, endColumn })` for one spreadsheet range or `selectCellRanges(ranges)` for a non-contiguous selection; merged cells expand the returned ranges. `copySelection()` writes to the Clipboard API when permission is available and always resolves to the deterministic text/TSV value. Multi-range TSV is emitted in row-major order, with tabs preserving gaps between selected cells on the same row. `getSelection()` and `clearSelection()` expose the current model.

Text/search offsets are UTF-16 code-unit offsets. Programmatic ranges preserve
that contract exactly; native drag/double-click endpoints are grapheme-aware so
the UI does not split emoji, combining sequences, or Indic clusters. DOCX
selectable spans use the renderer's exact font, pitch, rotation, and vertical
metadata. Search highlights live in a separate `pointer-events:none` overlay and
cannot change native selection geometry.

## Headless rendering

| Method                                                | Result                                                                                       |
| ----------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| `renderPage(index, canvas, options?)`                 | Render one page/slide/sheet/image into an existing `HTMLCanvasElement` or `OffscreenCanvas`. |
| `renderThumbnail(index, canvas, options?)`            | Render using a scale bounded by `maxWidth`/`maxHeight`.                                      |
| `renderSheetViewport(index, canvas, range, options?)` | Render a bounded row/column region.                                                          |
| `getPageText(index, signal?)`                         | Return logical text for one render unit.                                                     |

Render options are `zoom`, `devicePixelRatio`, optional `width`/`height`, and
`signal`. Sheet-region rendering additionally accepts unscaled
`scrollOffsetX/Y` for a partially clipped first column/row. Operations reject
stale completion after close/reload and normalize cancellation to `ViewerError`
with code `aborted`. See [headless rendering](./headless.md).

## Original bytes

`getOriginalBytes()` returns a defensive copy while a document is open. `downloadOriginal(fileName?)` returns a `Blob` containing the exact original bytes and, in a browser document, also starts a download. Legacy Office conversion output is never substituted for the original.

## Events

`viewer.on(type, listener)` returns an unsubscribe function.

| Event             | Payload                                                               |
| ----------------- | --------------------------------------------------------------------- |
| `statechange`     | Complete immutable `ViewerState`.                                     |
| `ready`           | Ready `ViewerState`.                                                  |
| `pagechange`      | `{ pageIndex, pageCount }`.                                           |
| `zoomchange`      | `{ zoom, fit }`.                                                      |
| `viewchange`      | Coalesced `ViewerState`, at most once per animation frame.            |
| `progress`        | Loading/detecting/parsing/converting/rendering progress.              |
| `searchchange`    | `SearchResult \| null`.                                               |
| `selectionchange` | `TextSelection \| CellRange \| null`.                                 |
| `warning`         | Non-fatal `ViewerWarning`.                                            |
| `error`           | Serializable `ViewerErrorData`. Promise-returning calls still reject. |

## Errors and warnings

Catch `ViewerError` and branch on its stable `code`: `unsupported-format`, `fidelity-unsupported`, `invalid-file`, `encrypted-document`, `resource-limit`, `network-error`, `aborted`, `font-unavailable`, `render-failed`, `worker-crashed`, `lifecycle-error`, or `internal`. `fidelity-unsupported` remains available for a recognized format/revision that no qualified backend can render without inventing or losing material structure. Qualified Word 97–2003 DOC uses the structured conversion path; older or malformed revisions fail typed conversion instead of falling back to plain text.

Warnings use `format-hint-mismatch`, `unsupported-feature`, `font-substitution`, `external-resource-blocked`, or `fidelity-degraded`. Warnings report explicit degradation and do not silently enable active content.
