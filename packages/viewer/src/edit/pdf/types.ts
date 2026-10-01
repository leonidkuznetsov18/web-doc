import type {
  ApplyOptions,
  BinaryData,
  EditElement,
  EditReceipt,
  EditSessionBase,
  PagePoint,
  PageRect,
} from "../types.js";

export type PdfElementKind =
  | "text" // one text object as stored in the file (often a word or a line)
  | "image"
  | "shape" // a path object
  | "textBox" // a text box created by insertTextBox
  | "table" // a table created by insertTable
  | "other"; // shadings, form XObjects and anything else

export type PdfTextAlign = "left" | "center" | "right";

export interface PdfTextStyle {
  readonly fontFamily: string;
  /** Points. */
  readonly fontSize: number;
  readonly bold: boolean;
  readonly italic: boolean;
  /** `#RRGGBB`. */
  readonly color: string;
  /** Text boxes only. */
  readonly align?: PdfTextAlign;
  /** Text boxes only; a multiple of the font size. */
  readonly lineHeight?: number;
}

export interface PdfShapeStyle {
  readonly stroke?: { readonly color: string; readonly width: number };
  readonly fill?: { readonly color: string };
}

export interface PdfElement extends EditElement {
  readonly kind: PdfElementKind;
  /** Present for `text` and `textBox`. */
  readonly textStyle?: PdfTextStyle;
  /** Present for `shape` and `table`. */
  readonly shapeStyle?: PdfShapeStyle;
  /** Present for `table`. */
  readonly table?: { readonly rows: readonly (readonly string[])[] };
}

export interface PdfTextBoxStyle {
  /** "Helvetica" (default), "Times" or "Courier"; registered families follow. */
  readonly fontFamily?: string;
  /** 1–500 points, default 12. */
  readonly fontSize?: number;
  readonly bold?: boolean;
  readonly italic?: boolean;
  /** `#RRGGBB`, default "#000000". */
  readonly color?: string;
  readonly align?: PdfTextAlign;
  /** Multiple of the font size, 0.5–5, default 1.2. */
  readonly lineHeight?: number;
}

export interface InsertTextBoxOperation {
  readonly op: "insertTextBox";
  readonly pageIndex: number;
  readonly rect: PageRect;
  readonly text: string;
  readonly style?: PdfTextBoxStyle;
}

export interface ReplaceTextOperation {
  readonly op: "replaceText";
  /** A `textBox` or `text` element. */
  readonly target: string;
  readonly text: string;
}

export interface SetTextStyleOperation {
  readonly op: "setTextStyle";
  /** A `textBox` or `text` element. */
  readonly target: string;
  readonly style: PdfTextBoxStyle;
}

export interface ResizeElementOperation {
  readonly op: "resizeElement";
  readonly target: string;
  /** New bounds in page space; a text box is laid out again inside them. */
  readonly rect: PageRect;
}

export interface MoveElementOperation {
  readonly op: "moveElement";
  readonly target: string;
  /** New top-left corner of the element's bounds, in page space. */
  readonly to?: PagePoint;
  /** Offset in page space. Exactly one of `to` and `by` is given. */
  readonly by?: { readonly dx: number; readonly dy: number };
}

export interface DeleteElementOperation {
  readonly op: "deleteElement";
  readonly target: string;
}

export interface InsertPageOperation {
  readonly op: "insertPage";
  /** Position of the new page, 0 to the current page count. */
  readonly index: number;
  /** Points; defaults to the size of the page before, else after, the position. */
  readonly size?: { readonly width: number; readonly height: number };
}

export interface DeletePageOperation {
  readonly op: "deletePage";
  readonly pageIndex: number;
}

export interface MovePageOperation {
  readonly op: "movePage";
  readonly from: number;
  /** The page's index after the move. */
  readonly to: number;
}

export interface PdfStroke {
  readonly color: string;
  /** Points; 0 draws the thinnest line the device can. */
  readonly width: number;
}

export interface PdfFill {
  readonly color: string;
}

export type InsertShapeOperation =
  | {
      readonly op: "insertShape";
      readonly pageIndex: number;
      readonly shape: "rectangle" | "ellipse";
      readonly rect: PageRect;
      readonly stroke?: PdfStroke;
      readonly fill?: PdfFill;
    }
  | {
      readonly op: "insertShape";
      readonly pageIndex: number;
      readonly shape: "line";
      readonly from: PagePoint;
      readonly to: PagePoint;
      readonly stroke?: PdfStroke;
    };

export interface InsertImageOperation {
  readonly op: "insertImage";
  readonly pageIndex: number;
  readonly rect: PageRect;
  /** PNG or JPEG bytes; base64 over pure-JSON transports. */
  readonly data: BinaryData;
  readonly mimeType: "image/png" | "image/jpeg";
}

export interface PdfTableStyle {
  /** "Helvetica" (default), "Times", "Courier" or a registered family. */
  readonly fontFamily?: string;
  /** 1–500 points, default 10. */
  readonly fontSize?: number;
  /** Text colour, `#RRGGBB`, default "#000000". */
  readonly color?: string;
  /** Grid colour, default "#000000". */
  readonly borderColor?: string;
  /** Grid stroke in points, 0–20, default 0.75. */
  readonly borderWidth?: number;
  /** Space between a cell's edges and its text, 0–100, default 4. */
  readonly cellPadding?: number;
  /** Fill of the first row; none by default. */
  readonly headerFill?: string;
}

export interface InsertTableOperation {
  readonly op: "insertTable";
  readonly pageIndex: number;
  /** Top-left corner in page space. */
  readonly at: PagePoint;
  /** Total width in points; the height follows the wrapped cell text. */
  readonly width: number;
  /** Cell text by row; every row has the same number of cells. 1–100 rows, 1–20 columns. */
  readonly rows: readonly (readonly string[])[];
  /** Relative column weights, one per column; equal when omitted. */
  readonly columnWidths?: readonly number[];
  readonly style?: PdfTableStyle;
}

export interface SetTableCellOperation {
  readonly op: "setTableCell";
  /** A `table` element. */
  readonly target: string;
  readonly row: number;
  readonly column: number;
  /** The cell's new text; empty clears the cell. */
  readonly text: string;
}

export interface SetShapeStyleOperation {
  readonly op: "setShapeStyle";
  readonly target: string;
  /** A new stroke, `null` to remove it, or absent to keep it. */
  readonly stroke?: PdfStroke | null;
  /** A new fill, `null` to remove it, or absent to keep it. */
  readonly fill?: PdfFill | null;
}

export interface RotatePageOperation {
  readonly op: "rotatePage";
  readonly pageIndex: number;
  /** Absolute clockwise rotation in degrees. */
  readonly rotation: 0 | 90 | 180 | 270;
}

/** The PDF operation union; operations are added as they ship. */
export type PdfOperation =
  | InsertTextBoxOperation
  | ReplaceTextOperation
  | SetTextStyleOperation
  | ResizeElementOperation
  | MoveElementOperation
  | DeleteElementOperation
  | InsertPageOperation
  | DeletePageOperation
  | MovePageOperation
  | RotatePageOperation
  | InsertShapeOperation
  | SetShapeStyleOperation
  | InsertImageOperation
  | InsertTableOperation
  | SetTableCellOperation;

/**
 * An operation's fields without its `op`, as the typed methods take them.
 * Distributes over unions so `insertShape` keeps its per-shape fields.
 */
export type Fields<T extends PdfOperation> = T extends unknown
  ? Omit<T, "op">
  : never;

export interface PdfEditSession extends EditSessionBase<
  PdfOperation,
  PdfElement
> {
  readonly format: "pdf";
  /** Lays `text` out inside `rect` as new text objects; the box's id is in `createdIds`. */
  insertTextBox(
    fields: Fields<InsertTextBoxOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Replaces the text of a text box (laid out again) or of a text object. */
  replaceText(
    fields: Fields<ReplaceTextOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes style fields; unspecified fields keep their value. */
  setTextStyle(
    fields: Fields<SetTextStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Gives an element new bounds; text boxes reflow, other elements stretch. */
  resizeElement(
    fields: Fields<ResizeElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Moves an element to a point or by an offset, in page space. */
  moveElement(
    fields: Fields<MoveElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes an element and, for a group, every object in it. */
  deleteElement(
    fields: Fields<DeleteElementOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Adds a blank page. */
  insertPage(
    fields: Fields<InsertPageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Removes a page; the last page cannot be removed. */
  deletePage(
    fields: Fields<DeletePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Reorders pages. */
  movePage(
    fields: Fields<MovePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Sets a page's rotation. */
  rotatePage(
    fields: Fields<RotatePageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Draws a rectangle, ellipse or line as a path object. */
  insertShape(
    fields: Fields<InsertShapeOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes or removes a shape's stroke and fill. */
  setShapeStyle(
    fields: Fields<SetShapeStyleOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Places a PNG or JPEG inside `rect`; JPEG data is embedded as it is. */
  insertImage(
    fields: Fields<InsertImageOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Draws a grid with wrapped cell text; rows grow to fit their tallest cell. */
  insertTable(
    fields: Fields<InsertTableOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  /** Changes one cell's text and lays the table out again. */
  setTableCell(
    fields: Fields<SetTableCellOperation>,
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
}
