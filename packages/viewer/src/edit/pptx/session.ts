import { ViewerError } from "../../errors.js";
import { readDescription, readOutline } from "../ai/outline.js";
import { resolveTargets } from "../ai/targets.js";
import { buildToolSet, callTool as runTool } from "../ai/tools.js";
import type {
  DescribeOptions,
  DocumentDescription,
  EditCheckpoint,
  OutlineOptions,
  OutlineResult,
  TargetCandidate,
  TargetQuery,
  ToolCall,
  ToolCallOptions,
  ToolResult,
  ToolSet,
} from "../ai/types.js";
import type { EditEngine, EditSessionCore } from "../engine.js";
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
  ReadItem,
  ReadOptions,
  ReadResult,
  SavedDocument,
  TextTarget,
} from "../types.js";
import type { PptxEngineReads } from "./engine.js";
import type {
  PptxDeleteElementOperation,
  PptxDeleteSlideOperation,
  PptxDuplicateSlideOperation,
  PptxEditSession,
  PptxElement,
  PptxFields,
  PptxInsertImageOperation,
  PptxInsertSlideOperation,
  PptxInsertTableOperation,
  PptxInsertTextBoxOperation,
  PptxLayoutInfo,
  PptxMoveElementOperation,
  PptxMoveSlideOperation,
  PptxOperation,
  PptxReplaceTextOperation,
  PptxResizeElementOperation,
  PptxSaveOptions,
  PptxSetShapeStyleOperation,
  PptxSetTableCellOperation,
  PptxSetTextStyleOperation,
  PptxSlideInfo,
} from "./types.js";

/**
 * The PPTX session: the core session narrowed to PPTX operations and
 * elements, plus the slide and layout reads. Each typed method is `apply()`
 * with a single operation.
 */
export class PptxSession implements PptxEditSession {
  readonly format = "pptx" as const;
  readonly #core: EditSessionCore;
  #tools: ToolSet | undefined;

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

  apply(
    operations: readonly PptxOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.#core.apply(operations, options);
  }

  applyJson(
    operations: readonly EditOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.#core.applyJson(operations, options);
  }

  undo(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.undo(options);
  }

  redo(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.redo(options);
  }

  reset(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.reset(options);
  }

  save(options?: PptxSaveOptions): Promise<SavedDocument> {
    return this.#core.save(options);
  }

  markSaved(stateToken: string): void {
    this.#core.markSaved(stateToken);
  }

  addAsset(data: Uint8Array, options?: AssetOptions): Promise<string> {
    return this.#core.addAsset(data, options);
  }

  getElements(
    query?: ElementQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<PptxElement>> {
    return this.#core.getElements(query, options) as Promise<
      ReadResult<PptxElement>
    >;
  }

  getElement(
    id: string,
    options?: ReadOptions,
  ): Promise<ReadItem<PptxElement>> {
    return this.#core.getElement(id, options) as Promise<ReadItem<PptxElement>>;
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
    options?: ReadOptions,
  ): Promise<ReadResult<PptxElement>> {
    return this.#core.elementsAt(pageIndex, point, options) as Promise<
      ReadResult<PptxElement>
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

  resolveTargets(
    query: TargetQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<TargetCandidate>> {
    return resolveTargets(this, query, options);
  }

  createCheckpoint(label?: string): Promise<EditCheckpoint> {
    return this.#core.createCheckpoint(label);
  }

  listCheckpoints(): readonly EditCheckpoint[] {
    return this.#core.listCheckpoints();
  }

  restoreCheckpoint(
    id: string,
    options?: HistoryOptions,
  ): Promise<EditReceipt> {
    return this.#core.restoreCheckpoint(id, options);
  }

  dropCheckpoint(id: string): void {
    this.#core.dropCheckpoint(id);
  }

  get tools(): ToolSet {
    return (this.#tools ??= buildToolSet(this.format, this.schemas));
  }

  callTool(call: ToolCall, options?: ToolCallOptions): Promise<ToolResult> {
    return runTool(this, this.tools, call, options);
  }

  getSlides(options?: ReadOptions): Promise<ReadResult<PptxSlideInfo>> {
    return this.#core.readItems(options, (engine, signal) =>
      pptxReads(engine).slides(signal),
    );
  }

  getLayouts(options?: ReadOptions): Promise<ReadResult<PptxLayoutInfo>> {
    return this.#core.readItems(options, (engine, signal) =>
      pptxReads(engine).layouts(signal),
    );
  }

  replaceText(
    fields: PptxFields<PptxReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "replaceText", ...fields }], options);
  }

  setTextStyle(
    fields: PptxFields<PptxSetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTextStyle", ...fields }], options);
  }

  setShapeStyle(
    fields: PptxFields<PptxSetShapeStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setShapeStyle", ...fields }], options);
  }

  moveElement(
    fields: PptxFields<PptxMoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "moveElement", ...fields }], options);
  }

  resizeElement(
    fields: PptxFields<PptxResizeElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "resizeElement", ...fields }], options);
  }

  deleteElement(
    fields: PptxFields<PptxDeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deleteElement", ...fields }], options);
  }

  insertTextBox(
    fields: PptxFields<PptxInsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTextBox", ...fields }], options);
  }

  insertImage(
    fields: PptxFields<PptxInsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertImage", ...fields }], options);
  }

  insertTable(
    fields: PptxFields<PptxInsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTable", ...fields }], options);
  }

  setTableCell(
    fields: PptxFields<PptxSetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "setTableCell", ...fields }], options);
  }

  insertSlide(
    fields: PptxFields<PptxInsertSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertSlide", ...fields }], options);
  }

  duplicateSlide(
    fields: PptxFields<PptxDuplicateSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "duplicateSlide", ...fields }], options);
  }

  deleteSlide(
    fields: PptxFields<PptxDeleteSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "deleteSlide", ...fields }], options);
  }

  moveSlide(
    fields: PptxFields<PptxMoveSlideOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "moveSlide", ...fields }], options);
  }
}

/** The engine's PPTX reads; an engine without them cannot serve this session. */
function pptxReads(engine: EditEngine): PptxEngineReads {
  const candidate = engine as Partial<PptxEngineReads>;
  if (
    typeof candidate.slides !== "function" ||
    typeof candidate.layouts !== "function"
  )
    throw new ViewerError(
      "edit-unsupported",
      "The engine does not provide the PPTX reads",
      { details: { format: "pptx", reason: "no-reads" } },
    );
  return candidate as PptxEngineReads;
}
