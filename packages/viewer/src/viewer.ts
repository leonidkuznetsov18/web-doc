import type {
  AdapterOpenContext,
  CellRange,
  CellSelection,
  DocumentAdapter,
  DocumentFormat,
  DocumentInfo,
  DocumentSource,
  FitMode,
  HeadlessRenderOptions,
  OpenDocumentOptions,
  ResourceLimits,
  SearchMatch,
  SearchOptions,
  SearchStrategy,
  SearchResult,
  SpreadsheetViewportRange,
  TextRun,
  TextSelection,
  TextSelectionRange,
  ThumbnailRenderOptions,
  ViewerApi,
  ViewerEventListener,
  ViewerEventMap,
  ViewerFetch,
  ViewerLogger,
  ViewerOptions,
  ViewerState,
} from "./contracts.js";
import type {
  EditableFormat,
  EditOptions,
  PageHit,
  PageRect,
  ViewportRect,
} from "./edit/types.js";
import type { EditEngineProvider } from "./edit/engine.js";
import type { EditSession } from "./edit/sessions.js";
import { EditSessionController, type EditSessionHost } from "./edit/session.js";
import { linkedAbortController } from "./abort.js";
import { detectFormat } from "./detect.js";
import { abortError, normalizeError, ViewerError } from "./errors.js";
import {
  cellRangeToTsv,
  cellRangesToTsv,
  findNormalizedMatches,
  normalizeCellRange,
} from "./interaction.js";
import {
  findFuzzyPageMatches,
  nearestMatchIndex,
  pagesNearestFirst,
  resolveFuzzySearchOptions,
  type FuzzyPageText,
  type ResolvedFuzzySearchOptions,
} from "./fuzzy-search.js";
import { FuzzyWorkerClient } from "./fuzzy-worker-client.js";
import { enforceContainerLimits, resolveLimits } from "./limits.js";
import type { AdapterRegistry } from "./registry.js";
import { loadDocumentSource } from "./source.js";
import { AdaptiveViewport } from "./viewport.js";
import type { FontManager } from "./fonts.js";
import { BasicViewerUi } from "./ui.js";
import { resolveTranslations } from "./i18n.js";
import { RenderScheduler } from "./render-scheduler.js";

const MIN_ZOOM = 0.1;
const MAX_ZOOM = 8;
const THUMBNAIL_BASE_WIDTH = 816;
const THUMBNAIL_BASE_HEIGHT = 1056;

export interface ViewerRuntime {
  readonly registry: AdapterRegistry;
  readonly fetch: ViewerFetch;
  readonly logger?: ViewerLogger;
  readonly assetBaseUrl?: URL;
  readonly limits: ResourceLimits;
  readonly fonts: FontManager;
  readonly renderScheduler: RenderScheduler;
  readonly release: (viewer: DocumentViewer) => void;
}

export class DocumentViewer implements ViewerApi {
  readonly #options: ViewerOptions;
  readonly #runtime: ViewerRuntime;
  readonly #listeners = new Map<
    keyof ViewerEventMap,
    Set<(event: never) => void>
  >();
  readonly #operations = new Set<AbortController>();
  readonly #textMaps = new Map<number, readonly TextRun[]>();
  readonly #viewport: AdaptiveViewport | undefined;
  readonly #ui: BasicViewerUi | undefined;
  #state: ViewerState;
  #adapter: DocumentAdapter | undefined;
  #handle: unknown;
  #info: DocumentInfo | undefined;
  #original: Uint8Array | undefined;
  #originalFileName: string | undefined;
  #originalContentType: string | undefined;
  /** Limits the current document was opened with; reused when it is reopened. */
  #limits: ResourceLimits;
  #session:
    | { readonly core: EditSessionController; readonly session: EditSession }
    | undefined;
  /** Shared by concurrent `edit()` calls while the engine starts. */
  #sessionStart: Promise<EditSession> | undefined;
  #activeLoad: AbortController | undefined;
  #activeSearch: AbortController | undefined;
  #selection: TextSelection | CellRange | CellSelection | null = null;
  #searchResult: SearchResult | null = null;
  /** Off-thread fuzzy matcher holding the current document's index. */
  #fuzzyWorker: FuzzyWorkerClient | undefined;
  /** Which pages/options the worker's index was built from; rebuilt on change. */
  #fuzzyIndexKey: string | undefined;
  /** Set once the worker failed to start, so the main thread takes over for good. */
  #fuzzyWorkerUnavailable = false;
  #generation = 0;
  #searchGeneration = 0;
  #viewEventScheduled = false;
  #textMapBytes = 0;

  constructor(options: ViewerOptions, runtime: ViewerRuntime) {
    this.#options = options;
    this.#runtime = runtime;
    this.#limits = runtime.limits;
    this.#state = freezeState({
      status: "idle",
      pageIndex: 0,
      pageCount: 0,
      zoom: clampZoom(options.initialZoom ?? 1),
      fit: options.fit ?? "none",
      panX: 0,
      panY: 0,
    });
    let viewportContainer = options.container;
    if (options.container && options.ui) {
      this.#ui = new BasicViewerUi(
        options.container,
        this,
        resolveTranslations(options.locale, options.translations),
      );
      this.#ui.root.lang = options.locale ?? "en";
      viewportContainer = this.#ui.viewportContainer;
    }
    if (viewportContainer) {
      const thisViewer = this;
      this.#viewport = new AdaptiveViewport(
        viewportContainer,
        {
          get state() {
            return thisViewer.state;
          },
          renderPage: (pageIndex, target, renderOptions) =>
            this.renderPage(pageIndex, target, renderOptions),
          renderSheetViewport: (sheetIndex, target, range, renderOptions) =>
            this.renderSheetViewport(sheetIndex, target, range, renderOptions),
          getTextRuns: (pageIndex, signal) =>
            this.#getTextRuns(pageIndex, signal),
          getSearchMatches: (pageIndex) => this.#matchesForPage(pageIndex),
          getActiveSearchMatch: () => this.#activeSearchMatch(),
          onVisiblePage: (pageIndex) => this.#goToPage(pageIndex, false),
          onPan: (panX, panY) => this.#setPan(panX, panY),
          onZoom: (zoom) => this.setZoom(zoom),
          onTextSelection: (range) => void this.selectText(range),
          onCellSelection: (selection) =>
            this.#applyInteractiveCellSelection(selection),
          getSelectionText: () => this.#getCachedSelectionText(),
        },
        options,
      );
    }
  }

  get state(): ViewerState {
    return this.#state;
  }

  async load(
    source: DocumentSource,
    options: OpenDocumentOptions = {},
  ): Promise<void> {
    this.#assertAlive();
    const generation = ++this.#generation;
    this.#abortOperations();
    this.#activeLoad?.abort();
    const controller = linkedAbortController(options.signal);
    this.#activeLoad = controller;
    await this.#closeDocument();
    if (controller.signal.aborted || generation !== this.#generation)
      throw abortError();
    this.#update({
      status: "loading",
      pageIndex: 0,
      pageCount: 0,
      panX: 0,
      panY: 0,
    });

    try {
      const limits = resolveLimits(this.#runtime.limits, options.limits);
      const loaded = await loadDocumentSource(source, {
        fetch: this.#runtime.fetch,
        signal: controller.signal,
        maxBytes: limits.maxInputBytes,
        reportProgress: (progress) => this.#emit("progress", progress),
      });
      this.#emit("progress", {
        phase: "detecting",
        loaded: loaded.bytes.byteLength,
      });
      const fileName = options.fileName ?? loaded.fileName;
      const contentType = options.contentType ?? loaded.contentType;
      const detection = detectFormat(loaded.bytes, {
        ...(options.format ? { format: options.format } : {}),
        ...(fileName ? { fileName } : {}),
        ...(contentType ? { contentType } : {}),
      });
      for (const warning of detection.warnings) this.#emit("warning", warning);
      enforceContainerLimits(loaded.bytes, detection.format, limits);
      const adapter = this.#runtime.registry.resolve(detection.format);
      const handle = await adapter.open(loaded.bytes, {
        format: detection.format,
        signal: controller.signal,
        limits,
        reportProgress: (progress) => this.#emit("progress", progress),
        reportWarning: (warning) => this.#emit("warning", warning),
        ...(fileName ? { fileName } : {}),
        ...(contentType ? { contentType } : {}),
        ...(this.#runtime.assetBaseUrl
          ? { assetBaseUrl: this.#runtime.assetBaseUrl }
          : {}),
      });
      if (controller.signal.aborted || generation !== this.#generation) {
        await adapter.close(handle);
        throw abortError();
      }
      let backendInfo: DocumentInfo;
      try {
        backendInfo = await adapter.getInfo(handle);
      } catch (error) {
        await adapter.close(handle);
        throw error;
      }
      if (generation !== this.#generation) {
        await adapter.close(handle);
        throw abortError();
      }
      let info: DocumentInfo;
      try {
        info = describeDocument(adapter, backendInfo, detection.format, limits);
      } catch (error) {
        await adapter.close(handle);
        throw error;
      }
      this.#adapter = adapter;
      this.#handle = handle;
      this.#info = info;
      this.#original = loaded.bytes;
      this.#originalFileName = fileName;
      this.#originalContentType = contentType;
      this.#limits = limits;
      this.#update({
        status: "ready",
        format: detection.format,
        pageIndex: 0,
        pageCount: info.pageCount,
      });
      for (const warning of info.warnings ?? []) this.#emit("warning", warning);
      this.#viewport?.setDocument(info);
      this.#emit("ready", this.#state);
      if (this.#options.fit === "width") this.fitWidth();
      else if (this.#options.fit === "page") this.fitPage();
    } catch (cause) {
      const error = normalizeError(cause, "invalid-file");
      if (generation === this.#generation) {
        this.#update({ status: "error" });
        this.#emit("error", error.toJSON());
        this.#runtime.logger?.error?.(error.message, { code: error.code });
      }
      throw error;
    } finally {
      if (this.#activeLoad === controller) this.#activeLoad = undefined;
    }
  }

  open(
    source: DocumentSource,
    options: OpenDocumentOptions = {},
  ): Promise<void> {
    return this.load(source, options);
  }

  async close(): Promise<void> {
    if (this.#state.status === "destroyed") return;
    this.#generation += 1;
    this.#activeLoad?.abort();
    this.#activeLoad = undefined;
    this.#abortOperations();
    await this.#closeDocument();
    const { format: _format, ...state } = this.#state;
    this.#state = freezeState({
      ...state,
      status: "idle",
      pageIndex: 0,
      pageCount: 0,
      panX: 0,
      panY: 0,
    });
    this.#viewport?.setDocument(undefined);
    this.#emit("statechange", this.#state);
    this.#scheduleViewEvent();
  }

  setZoom(zoom: number): void {
    this.#assertAlive();
    const next = clampZoom(zoom);
    this.#update({ zoom: next, fit: "none" });
  }

  zoomIn(): void {
    this.setZoom(this.#state.zoom * 1.2);
  }

  zoomOut(): void {
    this.setZoom(this.#state.zoom / 1.2);
  }

  setFit(fit: FitMode): void {
    this.#assertAlive();
    if (fit === "width") this.fitWidth();
    else if (fit === "page") this.fitPage();
    else this.#update({ fit: "none" });
  }

  fitWidth(): void {
    this.#assertAlive();
    const zoom = this.#viewport?.fitWidth() ?? this.#state.zoom;
    this.#update({ zoom: clampZoom(zoom), fit: "width" });
  }

  fitPage(): void {
    this.#assertAlive();
    const zoom = this.#viewport?.fitPage() ?? this.#state.zoom;
    this.#update({ zoom: clampZoom(zoom), fit: "page" });
  }

  panBy(deltaX: number, deltaY: number): void {
    this.#assertAlive();
    if (!Number.isFinite(deltaX) || !Number.isFinite(deltaY)) return;
    if (this.#viewport) this.#viewport.panBy(deltaX, deltaY);
    else this.#setPan(this.#state.panX + deltaX, this.#state.panY + deltaY);
  }

  goToPage(pageIndex: number): void {
    this.#goToPage(pageIndex, true);
  }

  next(): void {
    this.goToPage(this.#state.pageIndex + 1);
  }

  previous(): void {
    this.goToPage(this.#state.pageIndex - 1);
  }

  setSheet(sheetIndex: number): void {
    this.#assertReady();
    if (this.#info?.unit !== "sheet")
      throw new ViewerError(
        "lifecycle-error",
        "The current document has no sheets",
      );
    this.goToPage(sheetIndex);
  }

  async renderPage(
    pageIndex: number,
    target: HTMLCanvasElement | OffscreenCanvas,
    options: HeadlessRenderOptions = {},
  ): Promise<void> {
    const { adapter, handle, info } = this.#assertReady();
    assertPageIndex(pageIndex, info.pageCount);
    const operation = this.#startOperation(options.signal);
    const generation = this.#generation;
    try {
      await this.#runtime.renderScheduler.run(
        options.priority ?? "visible",
        operation.signal,
        async () => {
          const zoom = clampZoom(options.zoom ?? this.#state.zoom);
          const devicePixelRatio = positiveNumber(
            options.devicePixelRatio,
            globalThis.devicePixelRatio ?? 1,
          );
          assertRenderBudget(
            options.width ?? THUMBNAIL_BASE_WIDTH * zoom,
            options.height ?? THUMBNAIL_BASE_HEIGHT * zoom,
            devicePixelRatio,
            this.#runtime.limits.maxDecodedPixels,
          );
          this.#emit("progress", { phase: "rendering", loaded: 0, total: 1 });
          if (adapter.getTextMap) {
            const runs = await this.#getTextRuns(pageIndex, operation.signal);
            await this.#runtime.fonts.ensureRuns(
              runs,
              (warning) => this.#emit("warning", warning),
              operation.signal,
            );
          }
          await adapter.render(
            handle,
            target,
            {
              pageIndex,
              zoom,
              devicePixelRatio,
              ...(options.width === undefined ? {} : { width: options.width }),
              ...(options.height === undefined
                ? {}
                : { height: options.height }),
            },
            operation.signal,
          );
          if (generation !== this.#generation || operation.signal.aborted)
            throw abortError();
          this.#emit("progress", {
            phase: "rendering",
            loaded: 1,
            total: 1,
            ratio: 1,
          });
        },
      );
    } catch (error) {
      throw normalizeError(error, "render-failed");
    } finally {
      this.#finishOperation(operation);
    }
  }

  async renderThumbnail(
    pageIndex: number,
    target: HTMLCanvasElement | OffscreenCanvas,
    options: ThumbnailRenderOptions = {},
  ): Promise<void> {
    const scale = Math.min(
      1,
      positiveNumber(options.maxWidth, 180) / THUMBNAIL_BASE_WIDTH,
      positiveNumber(options.maxHeight, 240) / THUMBNAIL_BASE_HEIGHT,
    );
    await this.renderPage(pageIndex, target, {
      zoom: scale,
      priority: "background",
      ...(options.devicePixelRatio === undefined
        ? {}
        : { devicePixelRatio: options.devicePixelRatio }),
      ...(options.signal ? { signal: options.signal } : {}),
    });
  }

  async renderSheetViewport(
    sheetIndex: number,
    target: HTMLCanvasElement | OffscreenCanvas,
    range: SpreadsheetViewportRange,
    options: HeadlessRenderOptions = {},
  ): Promise<void> {
    const { adapter, handle, info } = this.#assertReady();
    if (info.unit !== "sheet")
      throw new ViewerError(
        "lifecycle-error",
        "The current document has no sheets",
      );
    assertPageIndex(sheetIndex, info.pageCount);
    const operation = this.#startOperation(options.signal);
    const generation = this.#generation;
    try {
      await this.#runtime.renderScheduler.run(
        options.priority ?? "visible",
        operation.signal,
        async () => {
          const zoom = clampZoom(options.zoom ?? this.#state.zoom);
          const devicePixelRatio = positiveNumber(
            options.devicePixelRatio,
            globalThis.devicePixelRatio ?? 1,
          );
          assertRenderBudget(
            options.width ?? THUMBNAIL_BASE_WIDTH * zoom,
            options.height ?? THUMBNAIL_BASE_HEIGHT * zoom,
            devicePixelRatio,
            this.#runtime.limits.maxDecodedPixels,
          );
          if (adapter.getTextMap) {
            const runs = await this.#getTextRuns(sheetIndex, operation.signal);
            await this.#runtime.fonts.ensureRuns(
              runs,
              (warning) => this.#emit("warning", warning),
              operation.signal,
            );
          }
          await adapter.render(
            handle,
            target,
            {
              pageIndex: sheetIndex,
              zoom,
              devicePixelRatio,
              sheetRange: normalizeSheetViewport(range),
              ...(options.width === undefined ? {} : { width: options.width }),
              ...(options.height === undefined
                ? {}
                : { height: options.height }),
              ...(options.scrollOffsetX === undefined
                ? {}
                : { scrollOffsetX: options.scrollOffsetX }),
              ...(options.scrollOffsetY === undefined
                ? {}
                : { scrollOffsetY: options.scrollOffsetY }),
              ...(options.columnWidths === undefined
                ? {}
                : { columnWidths: options.columnWidths }),
            },
            operation.signal,
          );
          if (generation !== this.#generation || operation.signal.aborted)
            throw abortError();
        },
      );
    } catch (error) {
      throw normalizeError(error, "render-failed");
    } finally {
      this.#finishOperation(operation);
    }
  }

  async getPageText(pageIndex: number, signal?: AbortSignal): Promise<string> {
    return (await this.#getTextRuns(pageIndex, signal))
      .map((run) => run.text)
      .join("");
  }

  getDocumentInfo(): DocumentInfo {
    return this.#assertReady().info;
  }

  async search(
    query: string,
    options: SearchOptions = {},
  ): Promise<SearchResult> {
    const { info } = this.#assertReady();
    // Validated before the in-flight search is cancelled so a rejected range
    // leaves the current result and highlights untouched.
    const [firstPage, lastPage] = resolveSearchPageRange(
      info.pageCount,
      options.pageRange,
    );
    const nearPage = resolveNearPage(firstPage, lastPage, options.nearPage);
    const fuzzy = resolveFuzzySearchOptions(
      this.#options.search?.fuzzy,
      options.fuzzy,
    );
    const caseSensitive = options.caseSensitive ?? false;
    this.#activeSearch?.abort();
    const controller = new AbortController();
    this.#activeSearch = controller;
    const generation = ++this.#searchGeneration;
    const cleanQuery = query;
    if (!cleanQuery) {
      this.clearSearch();
      return immutableSearchResult({ query, matches: [], activeIndex: -1 });
    }
    const matches: SearchMatch[] = [];
    let strategy: SearchStrategy = "exact";
    try {
      const texts = new Map<number, string>();
      for (let pageIndex = firstPage; pageIndex <= lastPage; pageIndex += 1) {
        if (controller.signal.aborted) throw abortError();
        const text = await this.getPageText(pageIndex, controller.signal);
        texts.set(pageIndex, text);
        matches.push(
          ...findNormalizedMatches(text, cleanQuery, pageIndex, caseSensitive),
        );
      }
      if (matches.length === 0 && fuzzy) {
        // With a hint the passage sits near it, so only that neighbourhood is
        // worth the fuzzy cost; without one every page is a candidate.
        const order = pagesNearestFirst(firstPage, lastPage, nearPage).slice(
          0,
          nearPage === undefined ? undefined : fuzzy.pageWindow,
        );
        const pattern = cleanQuery.slice(0, fuzzy.maxQueryLength);
        const pages: FuzzyPageText[] = [];
        for (const [pageIndex, text] of texts) pages.push({ pageIndex, text });
        const offThread = fuzzy.worker
          ? await this.#searchFuzzyInWorker(
              pages,
              pattern,
              order,
              fuzzy,
              caseSensitive,
              controller.signal,
            )
          : undefined;
        if (offThread) matches.push(...offThread);
        else
          for (
            let offset = 0;
            offset < order.length && matches.length === 0;
            offset += fuzzy.pagesPerBatch
          ) {
            // The page texts are already in memory, so the fallback is CPU
            // only. Pages are compared nearest to the hint first, a batch at
            // a time, and the scan stops at the first batch that holds the
            // passage; a yield between batches keeps a long document from
            // freezing the UI.
            if (offset > 0) await yieldToEventLoop();
            if (controller.signal.aborted) throw abortError();
            const batch = order
              .slice(offset, offset + fuzzy.pagesPerBatch)
              .map((pageIndex) => ({
                pageIndex,
                text: texts.get(pageIndex) ?? "",
              }));
            matches.push(
              ...findFuzzyPageMatches(batch, pattern, fuzzy, caseSensitive),
            );
          }
        if (matches.length > 0) strategy = "fuzzy";
      }
      if (generation !== this.#searchGeneration) throw abortError();
      const result = immutableSearchResult({
        query,
        matches,
        activeIndex: nearestMatchIndex(matches, nearPage),
        ...(matches.length > 0 ? { strategy } : {}),
      });
      this.#searchResult = result;
      if (result.activeIndex >= 0)
        this.#revealSearchMatch(result.matches[result.activeIndex]!);
      this.#emit("searchchange", result);
      this.#viewport?.update();
      return result;
    } finally {
      if (this.#activeSearch === controller) this.#activeSearch = undefined;
    }
  }

  /**
   * Fuzzy scan in the worker. The document's pages are indexed once per
   * (range, options) and reused by every search until the document closes;
   * `undefined` hands the scan back to the main thread when the worker is
   * unavailable or failed, so a missing worker asset degrades to slowness,
   * never to a lost match.
   */
  async #searchFuzzyInWorker(
    pages: readonly FuzzyPageText[],
    pattern: string,
    pageIndices: readonly number[],
    fuzzy: ResolvedFuzzySearchOptions,
    caseSensitive: boolean,
    signal: AbortSignal,
  ): Promise<readonly SearchMatch[] | undefined> {
    if (this.#fuzzyWorkerUnavailable) return undefined;
    try {
      this.#fuzzyWorker ??= FuzzyWorkerClient.create(this.#fuzzyWorkerUrl());
      if (!this.#fuzzyWorker) {
        this.#fuzzyWorkerUnavailable = true;
        return undefined;
      }
      const key = `${pages.map((page) => page.pageIndex).join(",")}|${fuzzy.threshold}|${fuzzy.maxPageTextLength}|${caseSensitive}`;
      if (this.#fuzzyIndexKey !== key) {
        this.#fuzzyIndexKey = undefined;
        await this.#fuzzyWorker.index(pages, {
          threshold: fuzzy.threshold,
          maxPageTextLength: fuzzy.maxPageTextLength,
          caseSensitive,
        });
        this.#fuzzyIndexKey = key;
      }
      const matches = await this.#fuzzyWorker.search(
        pattern,
        fuzzy.maxScore,
        pageIndices,
      );
      if (signal.aborted) throw abortError();
      return matches;
    } catch (error) {
      if (signal.aborted) throw error;
      this.#runtime.logger?.warn?.(
        "Fuzzy search worker unavailable; matching on the main thread",
        { message: error instanceof Error ? error.message : String(error) },
      );
      this.#dropFuzzyWorker();
      this.#fuzzyWorkerUnavailable = true;
      return undefined;
    }
  }

  #fuzzyWorkerUrl(): URL {
    return this.#runtime.assetBaseUrl
      ? new URL("workers/fuzzy-search-worker.js", this.#runtime.assetBaseUrl)
      : new URL("./workers/fuzzy-search-worker.js", import.meta.url);
  }

  #dropFuzzyWorker(): void {
    this.#fuzzyWorker?.terminate();
    this.#fuzzyWorker = undefined;
    this.#fuzzyIndexKey = undefined;
  }

  searchNext(): SearchResult | null {
    return this.#moveSearch(1);
  }

  searchPrevious(): SearchResult | null {
    return this.#moveSearch(-1);
  }

  clearSearch(): void {
    this.#activeSearch?.abort();
    this.#activeSearch = undefined;
    this.#searchGeneration += 1;
    this.#searchResult = null;
    this.#emit("searchchange", null);
    this.#viewport?.update();
  }

  getSelection(): TextSelection | CellRange | CellSelection | null {
    return this.#selection;
  }

  async selectText(range: TextSelectionRange): Promise<TextSelection> {
    const { info } = this.#assertReady();
    if (!info.capabilities?.textSelection)
      throw new ViewerError("lifecycle-error", "Text selection is unavailable");
    const normalized = normalizeTextSelectionRange(range);
    assertPageIndex(normalized.startPageIndex, info.pageCount);
    assertPageIndex(normalized.endPageIndex, info.pageCount);
    const selectedRuns: TextRun[] = [];
    const pageTexts: string[] = [];
    for (
      let pageIndex = normalized.startPageIndex;
      pageIndex <= normalized.endPageIndex;
      pageIndex += 1
    ) {
      const runs = await this.#getTextRuns(pageIndex);
      const text = runs.map((run) => run.text).join("");
      const start =
        pageIndex === normalized.startPageIndex ? normalized.startOffset : 0;
      const end =
        pageIndex === normalized.endPageIndex
          ? normalized.endOffset
          : text.length;
      const safeStart = Math.max(0, Math.min(text.length, start));
      const safeEnd = Math.max(safeStart, Math.min(text.length, end));
      pageTexts.push(text.slice(safeStart, safeEnd));
      selectedRuns.push(...sliceTextRuns(runs, safeStart, safeEnd));
    }
    const selection = Object.freeze({
      pageIndex: normalized.startPageIndex,
      ...(normalized.endPageIndex === normalized.startPageIndex
        ? {}
        : { endPageIndex: normalized.endPageIndex }),
      startOffset: normalized.startOffset,
      endOffset: normalized.endOffset,
      text: pageTexts.join("\n"),
      runs: Object.freeze(selectedRuns),
    });
    this.#selection = selection;
    this.#emit("selectionchange", selection);
    return selection;
  }

  selectCells(range: CellRange): CellRange {
    const { info } = this.#assertReady();
    if (info.unit !== "sheet")
      throw new ViewerError("lifecycle-error", "Cell selection is unavailable");
    assertPageIndex(range.sheetIndex, info.pageCount);
    const selection = Object.freeze(
      expandMergedRange(normalizeCellRange(range), info),
    );
    this.#selection = selection;
    this.#emit("selectionchange", selection);
    return selection;
  }

  selectCellRanges(ranges: readonly CellRange[]): CellSelection {
    const { info } = this.#assertReady();
    if (info.unit !== "sheet")
      throw new ViewerError("lifecycle-error", "Cell selection is unavailable");
    if (ranges.length === 0)
      throw new ViewerError(
        "lifecycle-error",
        "At least one cell range is required",
      );
    const sheetIndex = ranges[0]!.sheetIndex;
    assertPageIndex(sheetIndex, info.pageCount);
    const normalized = ranges.map((range) => {
      if (range.sheetIndex !== sheetIndex)
        throw new ViewerError(
          "lifecycle-error",
          "Cell ranges must belong to one sheet",
        );
      return Object.freeze(expandMergedRange(normalizeCellRange(range), info));
    });
    const selection = Object.freeze({
      sheetIndex,
      ranges: Object.freeze(normalized),
    });
    this.#selection = selection;
    this.#emit("selectionchange", selection);
    return selection;
  }

  clearSelection(): void {
    this.#selection = null;
    this.#emit("selectionchange", null);
  }

  async copySelection(): Promise<string> {
    const selection = this.#selection;
    if (!selection) return "";
    let text = this.#getCachedSelectionText();
    if (text === undefined) {
      if ("text" in selection) text = selection.text;
      else {
        const runs = await this.#getTextRuns(selection.sheetIndex);
        text = cellSelectionToTsv(selection, runs);
      }
    }
    try {
      await globalThis.navigator?.clipboard?.writeText(text);
    } catch {
      // Clipboard permission is host/browser controlled; returning text is deterministic.
    }
    return text;
  }

  #getCachedSelectionText(): string | undefined {
    const selection = this.#selection;
    if (!selection) return undefined;
    if ("text" in selection) return selection.text;
    const runs = this.#textMaps.get(selection.sheetIndex);
    return runs ? cellSelectionToTsv(selection, runs) : undefined;
  }

  #applyInteractiveCellSelection(
    selection: CellRange | CellSelection | null,
  ): void {
    if (!selection) this.clearSelection();
    else if ("ranges" in selection) this.selectCellRanges(selection.ranges);
    else this.selectCells(selection);
  }

  getOriginalBytes(): Uint8Array | undefined {
    this.#assertAlive();
    return this.#original?.slice();
  }

  downloadOriginal(fileName = this.#originalFileName ?? "document"): Blob {
    this.#assertAlive();
    if (!this.#original)
      throw new ViewerError("lifecycle-error", "No document is open");
    const blob = new Blob([this.#original.slice()], {
      type: "application/octet-stream",
    });
    if (typeof document !== "undefined") {
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = fileName;
      link.hidden = true;
      document.body.append(link);
      link.click();
      link.remove();
      queueMicrotask(() => URL.revokeObjectURL(url));
    }
    return blob;
  }

  async edit(options: EditOptions = {}): Promise<EditSession> {
    const { adapter, info } = this.#assertReady();
    if (this.#session?.core.usable) return this.#session.session;
    if (this.#sessionStart) return this.#sessionStart;
    const format = editableFormat(adapter, info.format);
    if (!format || !adapter.edit)
      throw new ViewerError(
        "edit-unsupported",
        "Editing is not available for this document",
        { details: { format: info.format, reason: "no-engine" } },
      );
    const start = this.#startSession(adapter.edit, format, options);
    this.#sessionStart = start;
    try {
      return await start;
    } finally {
      if (this.#sessionStart === start) this.#sessionStart = undefined;
    }
  }

  getEditSession(): EditSession | undefined {
    this.#assertAlive();
    return this.#session?.session;
  }

  pageToClient(pageIndex: number, rect: PageRect): ViewportRect | undefined {
    this.#assertAlive();
    if (
      this.#state.status !== "ready" ||
      !Number.isInteger(pageIndex) ||
      ![rect.x, rect.y, rect.width, rect.height].every(Number.isFinite)
    )
      return undefined;
    return this.#viewport?.pageToClient(pageIndex, rect);
  }

  clientToPage(clientX: number, clientY: number): PageHit | undefined {
    this.#assertAlive();
    if (
      this.#state.status !== "ready" ||
      !Number.isFinite(clientX) ||
      !Number.isFinite(clientY)
    )
      return undefined;
    return this.#viewport?.clientToPage(clientX, clientY);
  }

  on<K extends keyof ViewerEventMap>(
    type: K,
    listener: ViewerEventListener<K>,
  ): () => void {
    this.#assertAlive();
    const listeners =
      this.#listeners.get(type) ?? new Set<(event: never) => void>();
    listeners.add(listener as (event: never) => void);
    this.#listeners.set(type, listeners);
    return () => listeners.delete(listener as (event: never) => void);
  }

  async destroy(): Promise<void> {
    if (this.#state.status === "destroyed") return;
    this.#generation += 1;
    this.#activeLoad?.abort();
    this.#activeLoad = undefined;
    this.#abortOperations();
    await this.#closeDocument();
    this.#viewport?.destroy();
    this.#ui?.destroy();
    this.#state = freezeState({
      ...this.#state,
      status: "destroyed",
      pageIndex: 0,
      pageCount: 0,
      panX: 0,
      panY: 0,
    });
    this.#emit("statechange", this.#state);
    this.#listeners.clear();
    this.#runtime.release(this);
  }

  async #startSession(
    provider: EditEngineProvider,
    format: EditableFormat,
    options: EditOptions,
  ): Promise<EditSession> {
    const generation = this.#generation;
    const operation = this.#startOperation(options.signal);
    // A session that broke (failed recovery) makes way for a fresh engine.
    const stale = this.#session;
    this.#session = undefined;
    if (stale) await stale.core.end();
    try {
      const engine = await provider.load(this.#original!.slice(), {
        format,
        limits: this.#limits,
        fonts: this.#runtime.fonts.registered,
        signal: operation.signal,
        ...(this.#originalFileName ? { fileName: this.#originalFileName } : {}),
        ...(this.#runtime.assetBaseUrl
          ? { assetBaseUrl: this.#runtime.assetBaseUrl }
          : {}),
      });
      if (generation !== this.#generation || operation.signal.aborted) {
        await engine.dispose().catch(() => undefined);
        throw abortError();
      }
      const host: EditSessionHost = {
        format,
        limits: this.#limits,
        replaceDocument: (bytes, signal) =>
          this.#replaceDocument(bytes, signal),
        emit: (type, event) => this.#emit(type, event),
      };
      const core = new EditSessionController(
        engine,
        host,
        this.#original!,
        this.#info!.pageCount,
      );
      const session = provider.createSession(core);
      this.#session = { core, session };
      this.#emit("editstatechange", {
        ...core.state,
        active: true,
        format,
      });
      return session;
    } catch (error) {
      const normalized = normalizeError(error, "edit-failed");
      throw normalized.code === "edit-failed" && !normalized.details
        ? new ViewerError("edit-failed", normalized.message, {
            cause: error,
            details: { stage: "load" },
          })
        : normalized;
    } finally {
      this.#finishOperation(operation);
    }
  }

  /**
   * Shows edited bytes as the current document: opens them through the
   * adapter, swaps the handle, drops every cache derived from the old content
   * and keeps the view where it was. Returns the new page count.
   */
  async #replaceDocument(
    bytes: Uint8Array,
    signal: AbortSignal,
  ): Promise<number> {
    const { adapter, handle, info } = this.#assertReady();
    const generation = this.#generation;
    enforceContainerLimits(bytes, info.format, this.#limits);
    const context: AdapterOpenContext = {
      format: info.format,
      signal,
      limits: this.#limits,
      reportProgress: (progress) => this.#emit("progress", progress),
      reportWarning: (warning) => this.#emit("warning", warning),
      ...(this.#originalFileName ? { fileName: this.#originalFileName } : {}),
      ...(this.#originalContentType
        ? { contentType: this.#originalContentType }
        : {}),
      ...(this.#runtime.assetBaseUrl
        ? { assetBaseUrl: this.#runtime.assetBaseUrl }
        : {}),
    };
    const next = adapter.reopen
      ? await adapter.reopen(handle, bytes, context)
      : await adapter.open(bytes, context);
    let nextInfo: DocumentInfo;
    try {
      if (generation !== this.#generation || signal.aborted) throw abortError();
      const backendInfo = await adapter.getInfo(next);
      if (generation !== this.#generation || signal.aborted) throw abortError();
      nextInfo = describeDocument(
        adapter,
        backendInfo,
        info.format,
        this.#limits,
      );
    } catch (error) {
      await adapter.close(next);
      throw error;
    }
    this.#handle = next;
    this.#info = nextInfo;
    this.#invalidateContentCaches();
    this.#update({
      pageCount: nextInfo.pageCount,
      pageIndex: Math.min(this.#state.pageIndex, nextInfo.pageCount - 1),
    });
    for (const warning of nextInfo.warnings ?? [])
      this.#emit("warning", warning);
    this.#viewport?.replaceDocument(nextInfo);
    await adapter.close(handle);
    return nextInfo.pageCount;
  }

  /** Forgets everything derived from the document's content, not from its view. */
  #invalidateContentCaches(): void {
    this.#textMaps.clear();
    this.#textMapBytes = 0;
    // The worker keeps its index; a stale key makes the next search rebuild it.
    this.#fuzzyIndexKey = undefined;
    this.#activeSearch?.abort();
    this.#activeSearch = undefined;
    this.#searchGeneration += 1;
    if (this.#searchResult) {
      this.#searchResult = null;
      this.#emit("searchchange", null);
    }
    if (this.#selection) {
      this.#selection = null;
      this.#emit("selectionchange", null);
    }
  }

  async #getTextRuns(
    pageIndex: number,
    signal?: AbortSignal,
  ): Promise<readonly TextRun[]> {
    const { adapter, handle, info } = this.#assertReady();
    assertPageIndex(pageIndex, info.pageCount);
    if (signal?.aborted) throw abortError();
    const cached = this.#textMaps.get(pageIndex);
    if (cached) return cached;
    if (!adapter.getTextMap) return [];
    const operation = this.#startOperation(signal);
    const generation = this.#generation;
    try {
      const runs = await adapter.getTextMap(
        handle,
        pageIndex,
        operation.signal,
      );
      if (generation !== this.#generation || operation.signal.aborted)
        throw abortError();
      const immutableRuns = Object.freeze(
        runs.map((run) => Object.freeze({ ...run })),
      );
      const bytes = immutableRuns.reduce(
        (total, run) =>
          total + new TextEncoder().encode(run.text).byteLength + 96,
        0,
      );
      if (this.#textMapBytes + bytes > this.#runtime.limits.maxTextMapBytes)
        throw new ViewerError(
          "resource-limit",
          "Document text maps exceed the configured memory limit",
          {
            details: {
              bytes: this.#textMapBytes + bytes,
              limit: this.#runtime.limits.maxTextMapBytes,
            },
          },
        );
      this.#textMaps.set(pageIndex, immutableRuns);
      this.#textMapBytes += bytes;
      return immutableRuns;
    } catch (error) {
      throw normalizeError(error, "render-failed");
    } finally {
      this.#finishOperation(operation);
    }
  }

  #matchesForPage(pageIndex: number): readonly SearchMatch[] {
    return (
      this.#searchResult?.matches.filter(
        (match) => match.pageIndex === pageIndex,
      ) ?? []
    );
  }

  #activeSearchMatch(): SearchMatch | undefined {
    const result = this.#searchResult;
    return result && result.activeIndex >= 0
      ? result.matches[result.activeIndex]
      : undefined;
  }

  #moveSearch(direction: 1 | -1): SearchResult | null {
    const current = this.#searchResult;
    if (!current || current.matches.length === 0) return current;
    const activeIndex =
      (current.activeIndex + direction + current.matches.length) %
      current.matches.length;
    const result = immutableSearchResult({ ...current, activeIndex });
    this.#searchResult = result;
    this.#revealSearchMatch(result.matches[activeIndex]!);
    this.#emit("searchchange", result);
    this.#viewport?.update();
    return result;
  }

  /** Land on the match's page, then bring the match itself into view. */
  #revealSearchMatch(match: SearchMatch): void {
    this.#goToPage(match.pageIndex, true);
    void this.#viewport?.revealMatch(match);
  }

  #goToPage(pageIndex: number, scrollViewport: boolean): void {
    this.#assertAlive();
    const upperBound = Math.max(0, this.#state.pageCount - 1);
    const next = Math.min(Math.max(0, Math.trunc(pageIndex)), upperBound);
    this.#update({ pageIndex: next });
    if (scrollViewport) this.#viewport?.goToPage(next);
  }

  #setPan(panX: number, panY: number): void {
    const x = Math.max(0, Number.isFinite(panX) ? panX : 0);
    const y = Math.max(0, Number.isFinite(panY) ? panY : 0);
    if (x === this.#state.panX && y === this.#state.panY) return;
    this.#update({ panX: x, panY: y });
  }

  async #closeDocument(): Promise<void> {
    const session = this.#session;
    this.#session = undefined;
    this.#sessionStart = undefined;
    if (session) await session.core.end();
    const adapter = this.#adapter;
    const handle = this.#handle;
    this.#adapter = undefined;
    this.#handle = undefined;
    this.#info = undefined;
    this.#original = undefined;
    this.#originalFileName = undefined;
    this.#originalContentType = undefined;
    this.#activeSearch?.abort();
    this.#activeSearch = undefined;
    this.#searchResult = null;
    this.#dropFuzzyWorker();
    this.#selection = null;
    this.#textMaps.clear();
    this.#textMapBytes = 0;
    if (adapter && handle !== undefined) await adapter.close(handle);
  }

  #assertReady(): {
    adapter: DocumentAdapter;
    handle: unknown;
    info: DocumentInfo;
  } {
    this.#assertAlive();
    if (
      this.#state.status !== "ready" ||
      !this.#adapter ||
      this.#handle === undefined ||
      !this.#info
    )
      throw new ViewerError("lifecycle-error", "No document is ready");
    return { adapter: this.#adapter, handle: this.#handle, info: this.#info };
  }

  #assertAlive(): void {
    if (this.#state.status === "destroyed")
      throw new ViewerError("lifecycle-error", "DocumentViewer is destroyed");
  }

  #startOperation(signal?: AbortSignal): AbortController {
    const controller = linkedAbortController(signal);
    this.#operations.add(controller);
    return controller;
  }

  #finishOperation(controller: AbortController): void {
    this.#operations.delete(controller);
  }

  #abortOperations(): void {
    for (const controller of this.#operations) controller.abort();
    this.#operations.clear();
    this.#activeSearch?.abort();
  }

  #update(patch: Partial<ViewerState>): void {
    const previous = this.#state;
    this.#state = freezeState({ ...this.#state, ...patch });
    this.#emit("statechange", this.#state);
    if (previous.pageIndex !== this.#state.pageIndex)
      this.#emit("pagechange", {
        pageIndex: this.#state.pageIndex,
        pageCount: this.#state.pageCount,
      });
    if (previous.zoom !== this.#state.zoom || previous.fit !== this.#state.fit)
      this.#emit("zoomchange", {
        zoom: this.#state.zoom,
        fit: this.#state.fit,
      });
    this.#viewport?.update();
    this.#scheduleViewEvent();
  }

  #scheduleViewEvent(): void {
    if (this.#viewEventScheduled) return;
    this.#viewEventScheduled = true;
    const emit = (): void => {
      this.#viewEventScheduled = false;
      if (this.#state.status !== "destroyed")
        this.#emit("viewchange", this.#state);
    };
    if (typeof requestAnimationFrame === "function")
      requestAnimationFrame(emit);
    else queueMicrotask(emit);
  }

  #emit<K extends keyof ViewerEventMap>(
    type: K,
    event: ViewerEventMap[K],
  ): void {
    for (const listener of this.#listeners.get(type) ?? [])
      try {
        listener(event as never);
      } catch (error) {
        // A listener's bug is its own; it never fails the call that emitted.
        const report = (
          globalThis as { reportError?: (error: unknown) => void }
        ).reportError;
        if (report) report(error);
        else console.error(error);
      }
  }
}

/** The edit format of `format` when `adapter` can edit it. */
function editableFormat(
  adapter: DocumentAdapter,
  format: DocumentFormat,
): EditableFormat | undefined {
  const formats: readonly string[] = adapter.edit?.formats ?? [];
  return formats.includes(format) ? (format as EditableFormat) : undefined;
}

/** Validates backend info and freezes it with the viewer's capability flags. */
function describeDocument(
  adapter: DocumentAdapter,
  backendInfo: DocumentInfo,
  format: DocumentFormat,
  limits: ResourceLimits,
): DocumentInfo {
  if (
    !Number.isInteger(backendInfo.pageCount) ||
    backendInfo.pageCount < 1 ||
    backendInfo.pageCount > limits.maxDocumentUnits
  )
    throw new ViewerError(
      "resource-limit",
      "Document unit count exceeds the configured limit",
      {
        details: {
          pageCount: backendInfo.pageCount,
          limit: limits.maxDocumentUnits,
        },
      },
    );
  return immutableInfo({
    ...backendInfo,
    capabilities: {
      textSelection: Boolean(adapter.getTextMap),
      cellSelection: backendInfo.unit === "sheet",
      search: Boolean(adapter.getTextMap),
      thumbnails: backendInfo.unit !== "sheet",
      editing: editableFormat(adapter, format) !== undefined,
    },
  });
}

function clampZoom(value: number): number {
  if (!Number.isFinite(value)) return 1;
  return Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value));
}

function positiveNumber(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isFinite(value) && value > 0
    ? value
    : fallback;
}

function assertRenderBudget(
  cssWidth: number,
  cssHeight: number,
  devicePixelRatio: number,
  limit: number,
): void {
  const width = Math.max(
    1,
    Math.ceil(positiveNumber(cssWidth, 1) * devicePixelRatio),
  );
  const height = Math.max(
    1,
    Math.ceil(positiveNumber(cssHeight, 1) * devicePixelRatio),
  );
  const pixels = width * height;
  if (!Number.isSafeInteger(pixels) || pixels > limit)
    throw new ViewerError(
      "resource-limit",
      "Render target exceeds the configured pixel limit",
      { details: { width, height, pixels, limit } },
    );
}

/**
 * Resolve the inclusive 0-based page window a search scans. Omitting the range
 * searches the whole document; a reversed range is normalized so callers can
 * pass either endpoint order.
 */
function resolveSearchPageRange(
  pageCount: number,
  pageRange: readonly [number, number] | undefined,
): [number, number] {
  if (!pageRange) return [0, pageCount - 1];
  const [first, last] = pageRange;
  assertPageIndex(first, pageCount);
  assertPageIndex(last, pageCount);
  return first <= last ? [first, last] : [last, first];
}

/**
 * Clamp the approximate page hint into the scanned window. A hint that comes
 * from another pagination of the same file is allowed to overshoot; rejecting
 * it would defeat its purpose.
 */
function resolveNearPage(
  firstPage: number,
  lastPage: number,
  nearPage: number | undefined,
): number | undefined {
  if (nearPage === undefined || !Number.isFinite(nearPage)) return undefined;
  return Math.min(lastPage, Math.max(firstPage, Math.trunc(nearPage)));
}

function yieldToEventLoop(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function assertPageIndex(pageIndex: number, pageCount: number): void {
  if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pageCount)
    throw new ViewerError("render-failed", "Page index is out of range", {
      details: { pageIndex, pageCount },
    });
}

function normalizeSheetViewport(
  range: SpreadsheetViewportRange,
): SpreadsheetViewportRange {
  return {
    row: Math.max(1, Math.trunc(range.row)),
    column: Math.max(1, Math.trunc(range.column)),
    rowCount: Math.max(1, Math.trunc(range.rowCount)),
    columnCount: Math.max(1, Math.trunc(range.columnCount)),
  };
}

function normalizeTextSelectionRange(
  range: TextSelectionRange,
): TextSelectionRange {
  const forward =
    range.startPageIndex < range.endPageIndex ||
    (range.startPageIndex === range.endPageIndex &&
      range.startOffset <= range.endOffset);
  return forward
    ? {
        startPageIndex: Math.trunc(range.startPageIndex),
        startOffset: Math.max(0, Math.trunc(range.startOffset)),
        endPageIndex: Math.trunc(range.endPageIndex),
        endOffset: Math.max(0, Math.trunc(range.endOffset)),
      }
    : {
        startPageIndex: Math.trunc(range.endPageIndex),
        startOffset: Math.max(0, Math.trunc(range.endOffset)),
        endPageIndex: Math.trunc(range.startPageIndex),
        endOffset: Math.max(0, Math.trunc(range.startOffset)),
      };
}

function sliceTextRuns(
  runs: readonly TextRun[],
  start: number,
  end: number,
): TextRun[] {
  const selected: TextRun[] = [];
  let offset = 0;
  for (const run of runs) {
    const runStart = offset;
    const runEnd = offset + run.text.length;
    offset = runEnd;
    const sliceStart = Math.max(start, runStart);
    const sliceEnd = Math.min(end, runEnd);
    if (sliceStart >= sliceEnd || run.text.length === 0) continue;
    const localStart = sliceStart - runStart;
    const localEnd = sliceEnd - runStart;
    const fractionStart = localStart / run.text.length;
    const fractionWidth = (localEnd - localStart) / run.text.length;
    selected.push({
      ...run,
      text: run.text.slice(localStart, localEnd),
      x: run.x + run.width * fractionStart,
      width: run.width * fractionWidth,
    });
  }
  return selected;
}

function expandMergedRange(range: CellRange, info: DocumentInfo): CellRange {
  const sheet = info.sheets?.[range.sheetIndex];
  if (!sheet) return range;
  let result = range;
  let changed = true;
  while (changed) {
    changed = false;
    for (const merged of sheet.mergedRanges) {
      if (
        merged.endRow < result.startRow ||
        merged.startRow > result.endRow ||
        merged.endColumn < result.startColumn ||
        merged.startColumn > result.endColumn
      )
        continue;
      const expanded = {
        ...result,
        startRow: Math.min(result.startRow, merged.startRow),
        startColumn: Math.min(result.startColumn, merged.startColumn),
        endRow: Math.max(result.endRow, merged.endRow),
        endColumn: Math.max(result.endColumn, merged.endColumn),
      };
      changed =
        expanded.startRow !== result.startRow ||
        expanded.startColumn !== result.startColumn ||
        expanded.endRow !== result.endRow ||
        expanded.endColumn !== result.endColumn;
      result = expanded;
    }
  }
  return result;
}

function cellSelectionToTsv(
  selection: CellRange | CellSelection,
  runs: readonly TextRun[],
): string {
  const cells = new Map<string, string>();
  for (const run of runs)
    if (run.row !== undefined && run.column !== undefined)
      cells.set(`${run.row}:${run.column}`, run.text);
  return "ranges" in selection
    ? cellRangesToTsv(selection.ranges, cells)
    : cellRangeToTsv(selection, cells);
}

function freezeState(state: ViewerState): ViewerState {
  return Object.freeze(state);
}

function immutableInfo(info: DocumentInfo): DocumentInfo {
  return deepFreeze(structuredClone(info));
}

function immutableSearchResult(result: SearchResult): SearchResult {
  return Object.freeze({
    ...result,
    matches: Object.freeze(
      result.matches.map((match) => Object.freeze({ ...match })),
    ),
  });
}

function deepFreeze<T>(value: T): T {
  if (value && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value as Record<string, unknown>))
      deepFreeze(child);
  }
  return value;
}
