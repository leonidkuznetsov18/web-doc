import type { TextSelection } from "../../contracts.js";
import { ViewerError } from "../../errors.js";
import { readDescription, readOutline } from "../ai/outline.js";
import type {
  DescribeOptions,
  DocumentDescription,
  OutlineOptions,
  OutlineResult,
} from "../ai/types.js";
import type { EditEngine, EditSessionCore } from "../engine.js";
import { reportError } from "../session.js";
import { rectContains } from "./engine/geometry.js";
import { mapRangeThrough, type MutationRecord } from "./range-map.js";
import { resolveSelection } from "./selection.js";
import type {
  ApplyOptions,
  AssetOptions,
  EditFindOptions,
  EditOperation,
  EditReceipt,
  EditState,
  ElementQuery,
  HistoryOptions,
  OperationSchemaSet,
  PagePoint,
  PageRect,
  ReadItem,
  ReadOptions,
  ReadResult,
  SavedDocument,
  TextPosition,
  TextRange,
  TextTarget,
} from "../types.js";
import type { PdfEngineReads } from "./provider.js";
import type {
  DeleteElementOperation,
  DeletePageOperation,
  Fields,
  InsertImageOperation,
  InsertPageOperation,
  InsertShapeOperation,
  InsertTableOperation,
  InsertTextBoxOperation,
  MoveElementOperation,
  MovePageOperation,
  PdfEditSession,
  PdfElement,
  PdfOperation,
  PdfSaveOptions,
  ReplaceTextOperation,
  ResizeElementOperation,
  RotatePageOperation,
  SetShapeStyleOperation,
  SetTableCellOperation,
  PageBitmap,
  PageLayout,
  RenderOptions,
  SetTextStyleOperation,
  TextLayout,
} from "./types.js";

/**
 * The PDF session: the core session narrowed to PDF operations and elements.
 * Typed methods are added here as their operations ship; each is `apply()`
 * with a single operation.
 */
export class PdfSession implements PdfEditSession {
  readonly format = "pdf" as const;
  readonly #core: EditSessionCore;
  /** Committed calls, oldest first, for `mapRange`; bounded by `LOG_LIMIT`. */
  readonly #log: MutationRecord[] = [];
  /** Batches applied and not undone, and those undone and not redone. */
  #applied: (readonly EditOperation[])[] = [];
  #undone: (readonly EditOperation[])[] = [];
  /** The last elements read per page, with the revision they describe. */
  readonly #geometry = new Map<
    number,
    { readonly revision: number; readonly elements: readonly PdfElement[] }
  >();

  constructor(core: EditSessionCore) {
    this.#core = core;
  }

  get sessionId(): string {
    return this.#core.sessionId;
  }

  get state(): EditState {
    return this.#core.state;
  }

  get schemas(): OperationSchemaSet {
    return this.#core.schemas;
  }

  async apply(
    operations: readonly PdfOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    const receipt = await this.#core.apply(operations, options);
    this.#record("apply", operations, receipt);
    return receipt;
  }

  async applyJson(
    operations: readonly EditOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    const receipt = await this.#core.applyJson(operations, options);
    this.#record("apply", operations, receipt);
    return receipt;
  }

  async undo(options?: HistoryOptions): Promise<EditReceipt> {
    const receipt = await this.#core.undo(options);
    this.#record("undo", [], receipt);
    return receipt;
  }

  async redo(options?: HistoryOptions): Promise<EditReceipt> {
    const receipt = await this.#core.redo(options);
    this.#record("redo", [], receipt);
    return receipt;
  }

  async reset(options?: HistoryOptions): Promise<EditReceipt> {
    const receipt = await this.#core.reset(options);
    this.#record("reset", [], receipt);
    return receipt;
  }

  /** Keeps the mutation log in step with the core after a committed call. */
  #record(
    kind: MutationRecord["kind"],
    operations: readonly EditOperation[],
    receipt: EditReceipt,
  ): void {
    const last = this.#log.at(-1)?.revision ?? 0;
    // A dry run or a no-op leaves the revision alone and changes nothing.
    if (receipt.dryRun || receipt.revision === last) return;
    this.#refreshGeometry(receipt);
    let involved: readonly EditOperation[] = operations;
    switch (kind) {
      case "apply":
        this.#applied.push(operations);
        this.#undone = [];
        break;
      case "undo":
        involved = this.#applied.pop() ?? [];
        this.#undone.push(involved);
        break;
      case "redo":
        involved = this.#undone.pop() ?? [];
        this.#applied.push(involved);
        break;
      case "reset":
        involved = this.#applied.flat();
        this.#applied = [];
        this.#undone = [];
        break;
    }
    this.#log.push({
      revision: receipt.revision,
      kind,
      operations: involved,
      receipt,
    });
    if (this.#log.length > LOG_LIMIT) this.#log.splice(0, 1);
  }

  save(options?: PdfSaveOptions): Promise<SavedDocument> {
    return this.#core.save(options);
  }

  markSaved(stateToken: string): void {
    this.#core.markSaved(stateToken);
  }

  addAsset(data: Uint8Array, options?: AssetOptions): Promise<string> {
    return this.#core.addAsset(data, options);
  }

  async getElements(
    query?: ElementQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<PdfElement>> {
    const result = (await this.#core.getElements(
      query,
      options,
    )) as ReadResult<PdfElement>;
    // A whole page's elements feed the geometry cache.
    if (
      query?.pageIndex !== undefined &&
      query.kinds === undefined &&
      query.intersects === undefined
    )
      this.#remember(query.pageIndex, result);
    return result;
  }

  elementsAtSync(pageIndex: number, point: PagePoint): ReadResult<PdfElement> {
    const cached = this.#geometry.get(pageIndex);
    return Object.freeze({
      sessionId: this.sessionId,
      revision: cached?.revision ?? this.state.revision,
      items: Object.freeze(
        (cached?.elements ?? [])
          .filter((element) => rectContains(element.bounds, point))
          .reverse(),
      ),
    });
  }

  get cachedPages(): readonly number[] {
    return [...this.#geometry.keys()].sort((a, b) => a - b);
  }

  #remember(pageIndex: number, result: ReadResult<PdfElement>): void {
    const cached = this.#geometry.get(pageIndex);
    if (cached && cached.revision > result.revision) return;
    this.#geometry.set(pageIndex, {
      revision: result.revision,
      elements: result.items,
    });
  }

  /**
   * After a committed change the changed pages' entries are stale: they are
   * dropped and, while the page still exists, read again in the background
   * so the next hover answers from the new revision.
   */
  #refreshGeometry(receipt: EditReceipt): void {
    for (const pageIndex of [...this.#geometry.keys()]) {
      const changed =
        pageIndex >= receipt.pageCount ||
        receipt.changedPages.includes(pageIndex);
      if (!changed) continue;
      this.#geometry.delete(pageIndex);
      if (pageIndex < receipt.pageCount)
        this.getElements({ pageIndex }).catch(reportError);
    }
  }

  getElement(id: string, options?: ReadOptions): Promise<ReadItem<PdfElement>> {
    return this.#core.getElement(id, options) as Promise<ReadItem<PdfElement>>;
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadResult<PdfElement>> {
    return this.#core.elementsAt(pageIndex, point, options) as Promise<
      ReadResult<PdfElement>
    >;
  }

  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<ReadResult<TextTarget>> {
    return this.#core.findText(query, options);
  }

  getOutline(options?: OutlineOptions): Promise<OutlineResult> {
    return readOutline(this, this.#core.limits, options);
  }

  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>> {
    return readDescription(this, this.#core.limits, options);
  }

  getTextLayout(
    elementId: string,
    options?: ReadOptions,
  ): Promise<ReadItem<TextLayout>> {
    return this.#core.readItem(options, (engine, signal) =>
      pdfReads(engine).textLayout(elementId, signal),
    );
  }

  positionAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadItem<TextPosition>> {
    return this.#core.readItem(options, (engine, signal) =>
      pdfReads(engine).positionAt(pageIndex, point, signal),
    );
  }

  rangeRects(
    range: TextRange,
    options?: ReadOptions,
  ): Promise<ReadResult<PageRect>> {
    return this.#core.readItems(options, (engine, signal) =>
      pdfReads(engine).rangeRects(range, signal),
    );
  }

  getPageLayout(
    pageIndex: number,
    options?: ReadOptions,
  ): Promise<ReadItem<PageLayout>> {
    return this.#core.readItem(options, (engine, signal) =>
      pdfReads(engine).pageLayout(pageIndex, signal),
    );
  }

  elementsForSelection(
    selection: TextSelection,
    options?: ReadOptions,
  ): Promise<ReadResult<TextRange>> {
    return this.#core.readItems(options, async (engine, signal) => {
      const first = selection.pageIndex;
      const last = selection.endPageIndex ?? first;
      const pages: PageLayout[] = [];
      for (let pageIndex = first; pageIndex <= last; pageIndex += 1) {
        const layout = await pdfReads(engine).pageLayout(pageIndex, signal);
        if (layout) pages.push(layout);
      }
      return resolveSelection(selection, pages);
    });
  }

  mapRange(
    range: TextRange,
    fromRevision: number,
    options?: ReadOptions,
  ): Promise<ReadItem<TextRange>> {
    // Queued like any read, so every call committed before it is in the log.
    return this.#core.readItem(options, async (engine, signal) => {
      const current = this.state.revision;
      if (fromRevision > current || fromRevision < 0) return undefined;
      const records = this.#log.filter(
        (record) => record.revision > fromRevision,
      );
      // A gap means the log no longer reaches back that far.
      if (records.length !== current - fromRevision) return undefined;
      const mapped = mapRangeThrough(range, records);
      if (!mapped) return undefined;
      // The element as it is now bounds the offsets; a missing one ends it.
      const clamp = async (
        position: TextPosition,
      ): Promise<TextPosition | undefined> => {
        const element = engine.getElement
          ? await engine.getElement(position.elementId, signal)
          : (await engine.getElements({}, signal)).find(
              (entry) => entry.id === position.elementId,
            );
        if (!element) return undefined;
        const length = element.text?.length ?? 0;
        return { ...position, offset: Math.min(position.offset, length) };
      };
      const start = await clamp(mapped.start);
      const end = await clamp(mapped.end);
      return start && end ? { start, end } : undefined;
    });
  }

  renderPageWithout(
    pageIndex: number,
    elementIds: readonly string[],
    options: RenderOptions = {},
  ): Promise<ReadItem<PageBitmap>> {
    return this.#core.readItem(options, (engine, signal) =>
      pdfReads(engine).renderWithout(
        pageIndex,
        [...elementIds],
        options.scale ?? 1,
        signal,
      ),
    );
  }

  insertTextBox(
    fields: Fields<InsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTextBox", ...fields }], options);
  }

  replaceText(
    fields: Fields<ReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "replaceText", ...fields }], options);
  }

  setTextStyle(
    fields: Fields<SetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTextStyle", ...fields }], options);
  }

  resizeElement(
    fields: Fields<ResizeElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "resizeElement", ...fields }], options);
  }

  moveElement(
    fields: Fields<MoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "moveElement", ...fields }], options);
  }

  deleteElement(
    fields: Fields<DeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deleteElement", ...fields }], options);
  }

  insertPage(
    fields: Fields<InsertPageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertPage", ...fields }], options);
  }

  deletePage(
    fields: Fields<DeletePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deletePage", ...fields }], options);
  }

  movePage(
    fields: Fields<MovePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "movePage", ...fields }], options);
  }

  rotatePage(
    fields: Fields<RotatePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "rotatePage", ...fields }], options);
  }

  insertShape(
    fields: Fields<InsertShapeOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply(
      [{ op: "insertShape", ...fields } as InsertShapeOperation],
      options,
    );
  }

  setShapeStyle(
    fields: Fields<SetShapeStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setShapeStyle", ...fields }], options);
  }

  insertImage(
    fields: Fields<InsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertImage", ...fields }], options);
  }

  insertTable(
    fields: Fields<InsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTable", ...fields }], options);
  }

  setTableCell(
    fields: Fields<SetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTableCell", ...fields }], options);
  }
}

/** How many committed calls `mapRange` can look back over. */
const LOG_LIMIT = 512;

/** The engine behind a PDF session answers the overlay reads; a stand-in may not. */
function pdfReads(engine: EditEngine): PdfEngineReads {
  const reads = engine as Partial<PdfEngineReads>;
  if (
    !reads.textLayout ||
    !reads.positionAt ||
    !reads.rangeRects ||
    !reads.renderWithout ||
    !reads.pageLayout
  )
    throw new ViewerError(
      "edit-unsupported",
      "This engine does not implement the PDF overlay primitives",
      { details: { format: "pdf", reason: "no-overlay-primitives" } },
    );
  return reads as PdfEngineReads;
}
