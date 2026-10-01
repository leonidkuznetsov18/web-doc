import type { EditSessionCore } from "../engine.js";
import type {
  ApplyOptions,
  EditFindOptions,
  EditReceipt,
  EditState,
  ElementQuery,
  HistoryOptions,
  OperationSchemaSet,
  PagePoint,
  SaveOptions,
  TextTarget,
} from "../types.js";
import type {
  Fields,
  InsertTextBoxOperation,
  PdfEditSession,
  PdfElement,
  PdfOperation,
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

  undo(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.undo(options);
  }

  redo(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.redo(options);
  }

  reset(options?: HistoryOptions): Promise<EditReceipt> {
    return this.#core.reset(options);
  }

  save(options?: SaveOptions): Promise<Uint8Array> {
    return this.#core.save(options);
  }

  getElements(query?: ElementQuery): Promise<readonly PdfElement[]> {
    return this.#core.getElements(query) as Promise<readonly PdfElement[]>;
  }

  getElement(id: string): Promise<PdfElement | undefined> {
    return this.#core.getElement(id) as Promise<PdfElement | undefined>;
  }

  elementsAt(
    pageIndex: number,
    point: PagePoint,
  ): Promise<readonly PdfElement[]> {
    return this.#core.elementsAt(pageIndex, point) as Promise<
      readonly PdfElement[]
    >;
  }

  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<readonly TextTarget[]> {
    return this.#core.findText(query, options);
  }

  insertTextBox(
    fields: Fields<InsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt> {
    return this.apply([{ op: "insertTextBox", ...fields }], options);
  }
}
