import type {
  ApplyOptions,
  BinaryData,
  EditColor,
  EditElement,
  EditReceipt,
  EditSessionBase,
  SavedDocument,
  SaveOptions,
  TextRange,
} from "../types.js";

/*
 * The DOCX editing contract: the elements of the body story (paragraphs,
 * tables, inline pictures), their text and styles, and the operations that
 * change them. Geometry is page space (CSS pixels at 96 dpi) and comes from
 * the renderer's text runs; font sizes and spacing are points. Positions
 * in a flow document are other elements, never page points.
 */

export type DocxElementKind =
  | "paragraph" // w:p of the body or a table cell
  | "table" // w:tbl
  | "image" // an inline picture (w:drawing/wp:inline) inside a paragraph
  | "other"; // anchored drawings, OLE objects, equations: listed, not edited

export type DocxTextAlign = "left" | "center" | "right" | "justify";

/** Word's sixteen highlight colours (`w:highlight/@w:val`). */
export type DocxHighlight =
  | "yellow"
  | "green"
  | "cyan"
  | "magenta"
  | "blue"
  | "red"
  | "darkBlue"
  | "darkCyan"
  | "darkGreen"
  | "darkMagenta"
  | "darkRed"
  | "darkYellow"
  | "darkGray"
  | "lightGray"
  | "black"
  | "white";

/** The style of a paragraph's first run with text, resolved through the style chain and the theme. */
export interface DocxTextStyle {
  /** Theme fonts resolved ("minorHAnsi" becomes the minor Latin face). */
  readonly fontFamily: string;
  /** Points. */
  readonly fontSize: number;
  readonly bold: boolean;
  readonly italic: boolean;
  readonly underline: boolean;
  /** `#RRGGBB`, `{ theme }` for a theme colour, "auto" for Word's automatic colour. */
  readonly color: EditColor;
  readonly highlight?: DocxHighlight;
}

export interface DocxParagraphSpacing {
  /** Points before the paragraph; absent when inherited. */
  readonly before?: number;
  /** Points after the paragraph; absent when inherited. */
  readonly after?: number;
  /**
   * Line spacing: points for `exact` and `atLeast`, a multiple of single
   * spacing (1 = single, 1.5, 2) for `auto`; absent when inherited.
   */
  readonly line?: number;
  readonly lineRule?: "auto" | "exact" | "atLeast";
}

export interface DocxParagraphStyle {
  /** `w:pStyle`, when the paragraph names a style. */
  readonly styleId?: string;
  readonly align: DocxTextAlign;
  readonly spacing: DocxParagraphSpacing;
  /** `w:numPr`: the paragraph belongs to a list. */
  readonly numbering?: { readonly numId: number; readonly level: number };
}

/** Why a paragraph's text cannot be edited in place. */
export type DocxReadOnlyReason =
  "tracked-changes" | "section-break" | "unsupported-content";

/**
 * An element of the body story. Geometry comes from the renderer: `bounds`
 * and `fragments` are the unions of the element's text runs per page, so
 * an element whose pages the viewer has not laid out yet, or which draws
 * no run (an empty paragraph), has `pageIndex` −1, empty `bounds` and an
 * empty `fragments` list until a query names its page.
 */
export interface DocxElement extends EditElement {
  readonly kind: DocxElementKind;
  /** Present for a paragraph: resolved style of its first run with text. */
  readonly textStyle?: DocxTextStyle;
  /** Present for a paragraph. */
  readonly paragraphStyle?: DocxParagraphStyle;
  /** Present for a table: cell text by row. */
  readonly table?: { readonly rows: readonly (readonly string[])[] };
  /** Present for a paragraph whose text cannot be edited in place. */
  readonly readOnlyReason?: DocxReadOnlyReason;
  /** Present for an inline picture: its declared extent in page space (CSS pixels). */
  readonly imageSize?: { readonly width: number; readonly height: number };
}

export interface DocxTextStyleChange {
  /** `w:rFonts` ascii and hAnsi faces. */
  readonly fontFamily?: string;
  /** 1–400 points. */
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** `w:u` single or none. */
  readonly underline?: boolean;
  /** `#RRGGBB` writes `w:color`; `{ theme }` writes the theme colour and its resolved value. */
  readonly color?: EditColor;
  /** `w:highlight`; "none" removes it. */
  readonly highlight?: DocxHighlight | "none";
}

export interface DocxParagraphStyleChange {
  readonly align?: DocxTextAlign;
  /** Points; a field absent keeps its bytes. */
  readonly spacing?: {
    readonly before?: number;
    readonly after?: number;
    /** A multiple of single spacing, written with `w:lineRule="auto"`. */
    readonly line?: number;
  };
}

export interface DocxReplaceTextOperation {
  readonly op: "replaceText";
  /** A `paragraph` element. */
  readonly target: string;
  readonly text: string;
  /** Both ends on the target; absent, the whole text; collapsed, an insertion. */
  readonly range?: TextRange;
}

export interface DocxSetTextStyleOperation {
  readonly op: "setTextStyle";
  readonly target: string;
  readonly range?: TextRange;
  readonly style: DocxTextStyleChange;
}

export interface DocxSetParagraphStyleOperation {
  readonly op: "setParagraphStyle";
  readonly target: string;
  readonly style: DocxParagraphStyleChange;
}

export interface DocxInsertParagraphOperation {
  readonly op: "insertParagraph";
  /** Exactly one of `before` and `after`: a paragraph or table id. */
  readonly before?: string;
  readonly after?: string;
  readonly text: string;
  readonly style?: DocxTextStyleChange;
}

export interface DocxDeleteElementOperation {
  readonly op: "deleteElement";
  /** A paragraph, a table or an inline picture. */
  readonly target: string;
}

export interface DocxMoveElementOperation {
  readonly op: "moveElement";
  /** A paragraph or a table. */
  readonly target: string;
  /** Exactly one of `before` and `after`: a paragraph or table id of the same container. */
  readonly before?: string;
  readonly after?: string;
}

export interface DocxInsertTableOperation {
  readonly op: "insertTable";
  readonly before?: string;
  readonly after?: string;
  /** Cell text by row; every row has the same number of cells. 1–100 rows, 1–20 columns. */
  readonly rows: readonly (readonly string[])[];
  /** Relative column weights, one per column; equal when omitted. */
  readonly columnWidths?: readonly number[];
}

export interface DocxSetTableCellOperation {
  readonly op: "setTableCell";
  /** A `table` element. */
  readonly target: string;
  readonly row: number;
  readonly column: number;
  /** The cell's new text; empty clears the cell. */
  readonly text: string;
}

export interface DocxInsertImageOperation {
  readonly op: "insertImage";
  readonly before?: string;
  readonly after?: string;
  /** PNG or JPEG bytes; base64 over pure-JSON transports. */
  readonly data: BinaryData;
  readonly mimeType: "image/png" | "image/jpeg";
  /** Points. */
  readonly size: { readonly width: number; readonly height: number };
}

/** The DOCX operation union; operations are added as they ship. */
export type DocxOperation =
  | DocxReplaceTextOperation
  | DocxSetTextStyleOperation
  | DocxSetParagraphStyleOperation
  | DocxInsertParagraphOperation
  | DocxDeleteElementOperation
  | DocxMoveElementOperation
  | DocxInsertTableOperation
  | DocxSetTableCellOperation
  | DocxInsertImageOperation;

/** DOCX saves have no fields of their own. */
export type DocxSaveOptions = SaveOptions;

/** An operation's fields without its `op`, as the typed methods take them. */
export type DocxFields<T extends DocxOperation> = T extends unknown
  ? Omit<T, "op">
  : never;

export interface DocxEditSession extends EditSessionBase<
  DocxOperation,
  DocxElement
> {
  readonly format: "docx";
  save(options?: DocxSaveOptions): Promise<SavedDocument>;
  /** Replaces the whole text of a paragraph, or the part a range covers. */
  replaceText(
    fields: DocxFields<DocxReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes run properties; unspecified ones keep their bytes. */
  setTextStyle(
    fields: DocxFields<DocxSetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes alignment and spacing of a paragraph. */
  setParagraphStyle(
    fields: DocxFields<DocxSetParagraphStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a paragraph next to another element; its id is in `createdIds`. */
  insertParagraph(
    fields: DocxFields<DocxInsertParagraphOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes a paragraph, a table or an inline picture. */
  deleteElement(
    fields: DocxFields<DocxDeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Moves a paragraph or a table next to another element of the same container. */
  moveElement(
    fields: DocxFields<DocxMoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a table with the given cell text. */
  insertTable(
    fields: DocxFields<DocxInsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes one cell's text. */
  setTableCell(
    fields: DocxFields<DocxSetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a paragraph holding an inline PNG or JPEG. */
  insertImage(
    fields: DocxFields<DocxInsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
}
