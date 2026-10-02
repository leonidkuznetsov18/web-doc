import type {
  ChangeMode,
  EditableFormat,
  EditReceipt,
  HistoryOptions,
  JsonSchema,
  OperationIssue,
  ReadItem,
  ReadOptions,
  ReadResult,
  TextRange,
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

/** What a model names: a quoted text, a citation with its page, a kind, a place. */
export interface TargetQuery {
  /** Text to find; whitespace-insensitive, case-insensitive, tolerant to small differences. */
  readonly text?: string;
  /** A citation as the viewer's `search()` gets them: the passage and the 1-based page it is expected on. */
  readonly citation?: { readonly text: string; readonly pageNumber?: number };
  readonly pageIndex?: number;
  readonly kinds?: readonly string[];
  /** Restrict to an element and its descendants (a table, a group). */
  readonly within?: string;
  /** Default 5. */
  readonly maxResults?: number;
}

/** One way to read a query, with how it was found and how far to trust it. */
export interface TargetCandidate {
  readonly elementId: string;
  /** The matched part, for ranged operations; absent when the match is the element's whole text. */
  readonly range?: TextRange;
  readonly pageIndex: number;
  /** 1 for an exact match, down to 0.5 for the loosest accepted one. */
  readonly score: number;
  /** The pass that matched: the engine's exact search, folded text, fuzzy alignment, or a kind lookup. */
  readonly reason: "exact" | "normalized" | "fuzzy" | "kind-only";
  /** The matched text with a little context, on one line. */
  readonly snippet: string;
}

/** A state the host named, to come back to; session state, gone with the session. */
export interface EditCheckpoint {
  /** 22 URL-safe characters, unique in the session. */
  readonly id: string;
  readonly label?: string;
  /** The revision it names. */
  readonly revision: number;
  /** ISO 8601. */
  readonly createdAt: string;
}

/** One tool as a model provider takes it: a name, a description and a JSON Schema (draft 2020-12). */
export interface ToolDefinition {
  readonly name: string;
  readonly description: string;
  readonly inputSchema: JsonSchema;
}

/** The tools of a session; the same names on every format, the operations of its own. */
export interface ToolSet {
  /** Raised when a tool's shape changes incompatibly. */
  readonly version: number;
  readonly format: EditableFormat;
  readonly definitions: readonly ToolDefinition[];
}

/** A model's call: the tool's name and its JSON arguments, validated against the tool's schema. */
export interface ToolCall {
  readonly name: string;
  readonly arguments: unknown;
}

/** What a tool call produced, for the model and the chat. */
export interface ToolResult {
  readonly ok: boolean;
  /** JSON for the model: a description, an outline, candidates, a receipt, or the issues of a refused call. */
  readonly content: unknown;
  /** One or two sentences for the model and the chat: "Replaced the title of slide 2." */
  readonly text: string;
  /** Why a call was refused, in the shape `apply()` reports. */
  readonly issues?: readonly OperationIssue[];
}

export interface ToolCallOptions {
  /** Checked before any tool that changes the document, like `ApplyOptions.expectedRevision`. */
  readonly expectedRevision?: number;
  /** How `document_apply` writes: in place, or as tracked changes where the format has them. */
  readonly changeMode?: ChangeMode;
  /** The author of tracked changes; required with `changeMode: "tracked"`. */
  readonly author?: string;
  readonly signal?: AbortSignal;
}

/** The AI-facing reads, the checkpoints and the tools every edit session has. */
export interface EditSessionReads {
  /** The body elements in reading order, shaped for a prompt. */
  getOutline(options?: OutlineOptions): Promise<OutlineResult>;
  /** The outline as plain text, one line per element, within a character budget. */
  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>>;
  /**
   * Elements and ranges a query names, best first: an exact match, else a
   * match with whitespace, case, quotes and compatibility forms folded,
   * else the viewer's fuzzy citation match, else (without text) the
   * elements of the asked kinds in reading order.
   */
  resolveTargets(
    query: TargetQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<TargetCandidate>>;
  /**
   * Pins the current state under a new id. Rejects with `resource-limit`
   * past `maxEditCheckpoints`.
   */
  createCheckpoint(label?: string): Promise<EditCheckpoint>;
  /** Every checkpoint alive, in creation order. */
  listCheckpoints(): readonly EditCheckpoint[];
  /**
   * Back to the checkpoint's content as one history entry: undoable, a new
   * revision, `documentchange` with reason `restore`. Rejects with
   * `invalid-operation` for an unknown id.
   */
  restoreCheckpoint(id: string, options?: HistoryOptions): Promise<EditReceipt>;
  /** Forgets a checkpoint; unknown ids are ignored. */
  dropCheckpoint(id: string): void;
  /** The tool definitions of this session, for a model provider's tool list. */
  readonly tools: ToolSet;
  /**
   * Runs one tool call. A model's mistake (an unknown tool, bad arguments, a
   * refused batch, a stale revision) comes back as `ok: false` with issues
   * and a text to act on; only the session's own errors (`lifecycle-error`,
   * `aborted`) throw.
   */
  callTool(call: ToolCall, options?: ToolCallOptions): Promise<ToolResult>;
}
