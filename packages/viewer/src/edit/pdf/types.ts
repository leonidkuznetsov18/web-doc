import type {
  ApplyOptions,
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

/** The PDF operation union; operations are added as they ship. */
export type PdfOperation =
  | InsertTextBoxOperation
  | ReplaceTextOperation
  | SetTextStyleOperation
  | ResizeElementOperation
  | MoveElementOperation
  | DeleteElementOperation;

/** An operation's fields without its `op`, as the typed methods take them. */
export type Fields<T extends PdfOperation> = Omit<T, "op">;

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
}
