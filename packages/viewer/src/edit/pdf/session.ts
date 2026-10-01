import type { EditSessionCore } from "../engine.js";
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
  SetTextStyleOperation,
} from "./types.js";

/**
 * The PDF session: the core session narrowed to PDF operations and elements.
 * Typed methods are added here as their operations ship; each is `apply()`
 * with a single operation.
 */
export class PdfSession implements PdfEditSession {
  readonly format = "pdf" as const;
  readonly #core: EditSessionCore;

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
    operations: readonly PdfOperation[],
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

  save(options?: PdfSaveOptions): Promise<SavedDocument> {
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
  ): Promise<ReadResult<PdfElement>> {
    return this.#core.getElements(query, options) as Promise<
      ReadResult<PdfElement>
    >;
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
