import type {
  EditableFormat,
  ReadItem,
  ReadOptions,
  ReadResult,
} from "../types.js";

/*
 * The AI-facing reads of an edit session: the document as a prompt sees
 * it. Everything here is built over the session's own reads, so a host that
 * drives a model gets the same ids, operations and revisions as any client.
 */

export interface OutlineOptions extends ReadOptions {
  /** Inclusive 0-based page range; a node outside it is left out unless a descendant is inside. */
  readonly pageRange?: readonly [number, number];
  /** Element kinds to keep; a container of a kept node stays as its parent. */
  readonly kinds?: readonly string[];
  /** Characters of text kept per node; the rest is cut and `truncated` says so. Default 160. */
  readonly maxTextChars?: number;
}

/** One element as a prompt sees it: no geometry, no styles, nested where the format nests. */
export interface OutlineNode {
  /** The element id, as every operation takes it. */
  readonly id: string;
  /** The reading-order path: "3", "3.2" (a cell paragraph under its table). */
  readonly ordinal: string;
  /** The format's element kind. */
  readonly kind: string;
  /** −1 while the page is not laid out (a DOCX page the viewer has not cached). */
  readonly pageIndex: number;
  /** The first `maxTextChars` characters of the element's text. */
  readonly text?: string;
  /** Length of the whole text, in UTF-16 code units. */
  readonly textLength: number;
  /** True when `text` is shorter than the element's text. */
  readonly truncated: boolean;
  /** What a person calls it: "Title 1", "Heading 2", "Table (3×4)". */
  readonly label?: string;
  /** Why the element only accepts insertions next to it, when it does. */
  readonly readOnlyReason?: string;
  /** Listed and editable, but not drawn (a hidden PowerPoint shape). */
  readonly hidden?: boolean;
  /** Names of the operations that accept the element as their target. */
  readonly operations: readonly string[];
  /** A table's cell paragraphs, a group's members, a paragraph's inline pictures. */
  readonly children?: readonly OutlineNode[];
}

/** `getOutline()`'s read: the nodes, how many there are and whether the limit cut them. */
export interface OutlineResult extends ReadResult<OutlineNode> {
  /** Nodes in the result, nested ones included. */
  readonly nodeCount: number;
  /** True when `maxOutlineNodes` cut the outline. */
  readonly truncated: boolean;
}

export interface DescribeOptions extends OutlineOptions {
  /** Characters the description may take; the tail is cut and `truncated` says so. Default 50 000. */
  readonly maxChars?: number;
}

/** The outline rendered as plain text for a prompt. */
export interface DocumentDescription {
  readonly format: EditableFormat;
  readonly pageCount: number;
  /** Elements the description covers, nested ones included. */
  readonly elementCount: number;
  /**
   * A header line, then one line per node in the documented grammar:
   * `[sld2:7] slide 2 shape "Title 1": Quarterly review`.
   */
  readonly text: string;
  /** True when `maxChars` or `maxOutlineNodes` cut the description. */
  readonly truncated: boolean;
}

/** The AI-facing reads every edit session has. */
export interface EditSessionReads {
  /** The body elements in reading order, shaped for a prompt. */
  getOutline(options?: OutlineOptions): Promise<OutlineResult>;
  /** The outline as plain text, one line per element, within a character budget. */
  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>>;
}
