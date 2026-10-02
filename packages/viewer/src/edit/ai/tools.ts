import { ViewerError } from "../../errors.js";
import { binaryFields, isAssetReference } from "../assets.js";
import { validateSchema } from "../schema.js";
import type {
  ApplyOptions,
  EditableFormat,
  EditElement,
  EditOperation,
  EditReceipt,
  EditState,
  HistoryOptions,
  JsonSchema,
  OperationIssue,
  OperationSchemaSet,
  ReadItem,
  ReadOptions,
  ReadResult,
} from "../types.js";
import { foldText } from "./outline.js";
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
  ToolDefinition,
  ToolResult,
  ToolSet,
} from "./types.js";

/*
 * The tool set is the editing API as a model provider lists tools: plain
 * JSON Schema a host maps to its provider's format, and one executor that
 * validates a call, runs it on the session and answers with JSON for the
 * model and a sentence for the chat. A model's mistake never throws: it
 * comes back as `ok: false` with the issues `apply()` would report.
 */

export const TOOL_SET_VERSION = 1;

/** What the tools need of a session. */
export interface ToolSource {
  readonly format: EditableFormat;
  readonly state: EditState;
  readonly schemas: OperationSchemaSet;
  describe(options?: DescribeOptions): Promise<ReadItem<DocumentDescription>>;
  getOutline(options?: OutlineOptions): Promise<OutlineResult>;
  resolveTargets(
    query: TargetQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<TargetCandidate>>;
  getElement(id: string, options?: ReadOptions): Promise<ReadItem<EditElement>>;
  applyJson(
    operations: readonly EditOperation[],
    options?: ApplyOptions,
  ): Promise<EditReceipt>;
  undo(options?: HistoryOptions): Promise<EditReceipt>;
  createCheckpoint(label?: string): Promise<EditCheckpoint>;
  listCheckpoints(): readonly EditCheckpoint[];
  restoreCheckpoint(id: string, options?: HistoryOptions): Promise<EditReceipt>;
}

const ASSET_PATTERN = "^asset:[0-9a-f]{64}$";
const ID: JsonSchema = {
  type: "string",
  minLength: 1,
  maxLength: 200,
  description: "An element id as the outline lists it; copy it verbatim.",
};
const PAGE_RANGE: JsonSchema = {
  type: "array",
  items: { type: "integer", minimum: 0 },
  minItems: 2,
  maxItems: 2,
  description: "Inclusive 0-based [first, last] page range.",
};
const KINDS: JsonSchema = {
  type: "array",
  items: { type: "string", minLength: 1, maxLength: 40 },
  maxItems: 20,
  description: "Element kinds to keep.",
};
const MAX_TEXT_CHARS: JsonSchema = {
  type: "integer",
  minimum: 0,
  maximum: 20000,
  description: "Characters of text per element; default 160.",
};

function noun(format: EditableFormat): string {
  return format === "pptx" ? "slide" : "page";
}

/** The tool definitions of a format, built from its operation schemas. */
export function buildToolSet(
  format: EditableFormat,
  schemas: OperationSchemaSet,
): ToolSet {
  const pages = noun(format);
  const tracked =
    format === "docx"
      ? 'With changeMode "tracked" (set by the host) the batch is written as Word tracked changes for a person to accept or reject; replaceText, insertParagraph, deleteElement of a paragraph, setTableCell, setTextStyle and setParagraphStyle have a tracked form.'
      : `${format.toUpperCase()} has no tracked changes: changes apply directly; the host reviews them with document_checkpoint.`;
  const definitions: ToolDefinition[] = [
    {
      name: "document_describe",
      description: `The document as text, one line per element: [id] ${pages} N kind "label": text. Call it first; copy ids from the brackets. Budget in characters.`,
      inputSchema: object({
        pageRange: PAGE_RANGE,
        kinds: KINDS,
        maxTextChars: MAX_TEXT_CHARS,
        maxChars: {
          type: "integer",
          minimum: 1,
          description: "Characters the description may take; default 50 000.",
        },
      }),
    },
    {
      name: "document_outline",
      description: `The elements as JSON in reading order, nested where the format nests (table cells, group members), with ids, ${pages}s, text, labels and the operations each accepts.`,
      inputSchema: object({
        pageRange: PAGE_RANGE,
        kinds: KINDS,
        maxTextChars: MAX_TEXT_CHARS,
      }),
    },
    {
      name: "document_find",
      description:
        "Elements and text ranges a quoted text, a citation or a kind names, best first with a score (1 exact, 0.9 folded whitespace and case, below 0.85 fuzzy, 0.5 by kind only). Use the range of a candidate in replaceText and setTextStyle.",
      inputSchema: object({
        text: {
          type: "string",
          minLength: 1,
          maxLength: 2000,
          description: "Text to find; whitespace and case do not matter.",
        },
        citation: {
          type: "object",
          required: ["text"],
          additionalProperties: false,
          properties: {
            text: { type: "string", minLength: 1, maxLength: 2000 },
            pageNumber: {
              type: "integer",
              minimum: 1,
              description: `The 1-based ${pages} the passage is expected on; a hint.`,
            },
          },
        },
        pageIndex: {
          type: "integer",
          minimum: 0,
          description: `Only this 0-based ${pages}.`,
        },
        kinds: KINDS,
        within: {
          ...ID,
          description: "Only this element and its descendants.",
        },
        maxResults: { type: "integer", minimum: 1, maximum: 50 },
      }),
    },
    {
      name: "document_inspect",
      description:
        "One element with its text, style, table rows and the operations it accepts; geometry left out.",
      inputSchema: object({ id: ID }, ["id"]),
    },
    {
      name: "document_preview",
      description:
        "Validates a batch and reports what applying it would do (pages changed, elements created) without changing anything. Recommended before document_apply for batches above one operation.",
      inputSchema: batchSchema(schemas, false),
    },
    {
      name: "document_apply",
      description: `Applies a batch of operations as one undo step and returns the receipt. Address elements by the ids document_describe lists; the same batch is validated like document_preview. ${tracked}`,
      inputSchema: batchSchema(schemas, true),
    },
    {
      name: "document_undo",
      description: "Undoes the last change (a batch or a checkpoint restore).",
      inputSchema: object({}),
    },
    {
      name: "document_checkpoint",
      description:
        "Named checkpoints: create one before a series of changes, list them, or restore one to drop everything since it (one undo step).",
      inputSchema: object(
        {
          action: { enum: ["create", "restore", "list"] },
          id: {
            type: "string",
            minLength: 1,
            maxLength: 40,
            description: "The checkpoint to restore.",
          },
          label: { type: "string", maxLength: 200 },
        },
        ["action"],
      ),
    },
  ];
  return Object.freeze({
    version: TOOL_SET_VERSION,
    format,
    definitions: Object.freeze(
      definitions.map((definition) => Object.freeze(definition)),
    ),
  });
}

function object(
  properties: Readonly<Record<string, JsonSchema>>,
  required: readonly string[] = [],
): JsonSchema {
  return {
    type: "object",
    ...(required.length > 0 ? { required } : {}),
    additionalProperties: false,
    properties,
  };
}

/**
 * The batch argument: every operation of the format as one `oneOf`, with
 * the formats' shared definitions hoisted to the root and binary fields
 * narrowed to asset references, which a host registers with `addAsset()`.
 */
function batchSchema(
  schemas: OperationSchemaSet,
  withLabel: boolean,
): JsonSchema {
  const defs: Record<string, unknown> = {};
  const branches = Object.values(schemas.operations).map((schema) => {
    const { $defs, ...rest } = schema;
    if ($defs && typeof $defs === "object") Object.assign(defs, $defs);
    return assetsByReference(rest);
  });
  return {
    type: "object",
    required: ["operations"],
    additionalProperties: false,
    properties: {
      operations: {
        type: "array",
        minItems: 1,
        items: { oneOf: branches },
        description: `Operations of the ${schemas.format} session, applied in order; "$<n>" as a target names the element operation n created.`,
      },
      ...(withLabel
        ? {
            label: {
              type: "string",
              maxLength: 200,
              description: "What the change is, for the host's history.",
            },
          }
        : {}),
    },
    ...(Object.keys(defs).length > 0 ? { $defs: defs } : {}),
  };
}

function assetsByReference(schema: JsonSchema): JsonSchema {
  const out: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(schema)) {
    if (key === "x-binary") continue;
    out[key] = isRecord(value)
      ? value["x-binary"] === true
        ? {
            type: "string",
            pattern: ASSET_PATTERN,
            description:
              "An asset reference (asset:<sha-256>) the host registered with addAsset(); never raw bytes or base64.",
          }
        : assetsByReference(value)
      : Array.isArray(value)
        ? value.map((item) => (isRecord(item) ? assetsByReference(item) : item))
        : value;
  }
  return out;
}

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Runs one tool call on a session. */
export async function callTool(
  session: ToolSource,
  tools: ToolSet,
  call: ToolCall,
  options: ToolCallOptions = {},
): Promise<ToolResult> {
  const definition = tools.definitions.find((tool) => tool.name === call.name);
  if (!definition) {
    const name = String(call.name).slice(0, 80);
    return refused(
      [
        {
          operationIndex: -1,
          path: "/name",
          code: "unknown-tool",
          message: `Unknown tool ${name}`,
        },
      ],
      `No tool named ${name}; the tools are ${tools.definitions.map((tool) => tool.name).join(", ")}.`,
    );
  }
  const args = call.arguments ?? {};
  const schema =
    definition.name === "document_apply"
      ? APPLY_ARGUMENTS
      : definition.name === "document_preview"
        ? PREVIEW_ARGUMENTS
        : definition.inputSchema;
  const issues = validateSchema(args, schema, -1);
  if (issues.length > 0)
    return refused(
      issues,
      `Invalid arguments for ${definition.name}: ${issues
        .map((issue) => `${issue.path || "/"} ${issue.message}`)
        .join("; ")}.`,
    );
  const read = options.signal ? { signal: options.signal } : {};
  try {
    switch (definition.name) {
      case "document_describe":
        return await describeTool(session, args as DescribeOptions, read);
      case "document_outline":
        return await outlineTool(session, args as OutlineOptions, read);
      case "document_find":
        return await findTool(session, args as TargetQuery, read);
      case "document_inspect":
        return await inspectTool(session, (args as { id: string }).id, read);
      case "document_preview":
        return await batchTool(session, args as BatchArguments, options, true);
      case "document_apply":
        return await batchTool(session, args as BatchArguments, options, false);
      case "document_undo":
        return await undoTool(session, options);
      default:
        return await checkpointTool(
          session,
          args as CheckpointArguments,
          options,
        );
    }
  } catch (error) {
    return refusal(session, error);
  }
}

/** The loose shape of a preview call; the operations themselves are validated by `apply()`. */
const PREVIEW_ARGUMENTS: JsonSchema = {
  type: "object",
  required: ["operations"],
  additionalProperties: false,
  properties: {
    operations: {
      type: "array",
      minItems: 1,
      items: { type: "object" },
    },
  },
};

/** The loose shape of an apply call: a preview plus the history label. */
const APPLY_ARGUMENTS: JsonSchema = {
  ...PREVIEW_ARGUMENTS,
  properties: {
    ...(PREVIEW_ARGUMENTS.properties as Record<string, JsonSchema>),
    label: { type: "string", maxLength: 200 },
  },
};

interface BatchArguments {
  readonly operations: readonly EditOperation[];
  readonly label?: string;
}

interface CheckpointArguments {
  readonly action: "create" | "restore" | "list";
  readonly id?: string;
  readonly label?: string;
}

function refused(issues: readonly OperationIssue[], text: string): ToolResult {
  const frozen = Object.freeze(
    issues.map((issue) => Object.freeze({ ...issue })),
  );
  return Object.freeze({
    ok: false,
    content: { issues: frozen },
    text,
    issues: frozen,
  });
}

function ok(content: unknown, text: string): ToolResult {
  return Object.freeze({ ok: true, content, text });
}

/** A session error a model can act on becomes a refusal; the session's own errors propagate. */
function refusal(session: ToolSource, error: unknown): ToolResult {
  if (!(error instanceof ViewerError)) throw error;
  if (error.code === "lifecycle-error" || error.code === "aborted") throw error;
  // A session that could not recover answers nothing a model can act on.
  if (
    error.code === "worker-crashed" ||
    (error.code === "edit-failed" && error.details?.recovered === false)
  )
    throw error;
  if (error.code === "invalid-operation") {
    const reported = error.details?.issues as
      readonly OperationIssue[] | undefined;
    const issues =
      reported && reported.length > 0
        ? reported
        : [
            {
              operationIndex: -1,
              path: "",
              code: "invalid-operation",
              message: error.message,
            },
          ];
    const first = issues[0];
    return refused(
      issues,
      first
        ? `Refused: operation ${first.operationIndex < 0 ? "batch" : first.operationIndex} at ${first.path || "/"}: ${first.message}. Fix the operation and send the batch again.`
        : `Refused: ${error.message}.`,
    );
  }
  if (error.code === "edit-conflict")
    return refused(
      [
        {
          operationIndex: -1,
          path: "",
          code: "edit-conflict",
          message: error.message,
        },
      ],
      `The document changed since it was read (now revision ${session.state.revision}); call document_describe again and retry with the new revision.`,
    );
  return refused(
    [
      {
        operationIndex: -1,
        path: "",
        code: error.code,
        message: error.message,
      },
    ],
    `${error.message}. The document is unchanged.`,
  );
}

async function describeTool(
  session: ToolSource,
  args: DescribeOptions,
  read: ReadOptions,
): Promise<ToolResult> {
  const result = await session.describe({ ...args, ...read });
  const item = result.item!;
  return ok(
    { revision: result.revision, ...item },
    `Described ${count(item.elementCount, "element")} on ${count(item.pageCount, noun(session.format))} at revision ${result.revision}${item.truncated ? " (cut to the budget)" : ""}.`,
  );
}

async function outlineTool(
  session: ToolSource,
  args: OutlineOptions,
  read: ReadOptions,
): Promise<ToolResult> {
  const result = await session.getOutline({ ...args, ...read });
  return ok(
    {
      revision: result.revision,
      nodeCount: result.nodeCount,
      truncated: result.truncated,
      nodes: result.items,
    },
    `Outlined ${count(result.nodeCount, "element")} at revision ${result.revision}${result.truncated ? " (cut at the node limit)" : ""}.`,
  );
}

async function findTool(
  session: ToolSource,
  query: TargetQuery,
  read: ReadOptions,
): Promise<ToolResult> {
  const result = await session.resolveTargets(query, read);
  const best = result.items[0];
  return ok(
    { revision: result.revision, candidates: result.items },
    best
      ? `Found ${count(result.items.length, "candidate")}; best: ${best.reason} match in ${best.elementId}${best.pageIndex >= 0 ? ` on ${noun(session.format)} ${best.pageIndex + 1}` : ""} (score ${best.score}).`
      : "Nothing matched; try a shorter quote, a different page, or document_describe to see the text.",
  );
}

async function inspectTool(
  session: ToolSource,
  id: string,
  read: ReadOptions,
): Promise<ToolResult> {
  const result = await session.getElement(id, read);
  const element = result.item;
  if (!element)
    return refused(
      [
        {
          operationIndex: -1,
          path: "/id",
          code: "unknown-target",
          message: `No element ${id}`,
        },
      ],
      `No element ${id}; ids come from document_describe at the current revision.`,
    );
  const {
    bounds: _bounds,
    frame: _frame,
    fragments: _fragments,
    rotation: _rotation,
    ...rest
  } = element as EditElement & Record<string, unknown>;
  const text = element.text;
  return ok(
    { revision: result.revision, ...rest },
    `${element.kind} ${element.id}${element.pageIndex >= 0 ? ` on ${noun(session.format)} ${element.pageIndex + 1}` : ""}${
      text
        ? `: ${foldText(text.slice(0, 80))}${text.length > 80 ? "…" : ""}`
        : ""
    }. Operations: ${element.operations.join(", ") || "none"}.`,
  );
}

async function batchTool(
  session: ToolSource,
  args: BatchArguments,
  options: ToolCallOptions,
  dryRun: boolean,
): Promise<ToolResult> {
  const assetIssues = checkAssets(session.schemas, args.operations);
  if (assetIssues.length > 0)
    return refused(
      assetIssues,
      "Binary payloads travel as asset references: the host registers the bytes with addAsset() and passes the asset:<sha-256> reference; never send base64 in a tool call.",
    );
  const receipt = await session.applyJson(args.operations, {
    ...(options.expectedRevision === undefined
      ? {}
      : { expectedRevision: options.expectedRevision }),
    ...(options.changeMode === undefined
      ? {}
      : { changeMode: options.changeMode }),
    ...(options.author === undefined ? {} : { author: options.author }),
    ...(options.signal ? { signal: options.signal } : {}),
    ...(args.label === undefined ? {} : { label: args.label }),
    dryRun,
  });
  return ok(receipt, describeReceipt(session.format, args.operations, receipt));
}

/** Binary fields of a batch must be asset references before `apply()` sees them. */
function checkAssets(
  schemas: OperationSchemaSet,
  operations: readonly EditOperation[],
): OperationIssue[] {
  const issues: OperationIssue[] = [];
  operations.forEach((operation, operationIndex) => {
    if (!isRecord(operation) || typeof operation.op !== "string") return;
    for (const field of binaryFields(schemas.operations[operation.op])) {
      const value = operation[field];
      if (value === undefined || isAssetReference(value)) continue;
      issues.push({
        operationIndex,
        path: `/${field}`,
        code: "unknown-asset",
        message: "Expected an asset reference registered with addAsset()",
      });
    }
  });
  return issues;
}

async function undoTool(
  session: ToolSource,
  options: ToolCallOptions,
): Promise<ToolResult> {
  const receipt = await session.undo({
    ...(options.expectedRevision === undefined
      ? {}
      : { expectedRevision: options.expectedRevision }),
    ...(options.signal ? { signal: options.signal } : {}),
  });
  return ok(
    receipt,
    receipt.operationCount === 0 && receipt.changedPages.length === 0
      ? session.state.canUndo
        ? `Undid the last change; revision ${receipt.revision}.`
        : "Nothing to undo."
      : `Undid the last change${receipt.removedIds.length > 0 ? `, removing ${count(receipt.removedIds.length, "element")}` : ""}; revision ${receipt.revision}.`,
  );
}

async function checkpointTool(
  session: ToolSource,
  args: CheckpointArguments,
  options: ToolCallOptions,
): Promise<ToolResult> {
  switch (args.action) {
    case "create": {
      const checkpoint = await session.createCheckpoint(args.label);
      return ok(
        checkpoint,
        `Checkpoint ${checkpoint.id} created at revision ${checkpoint.revision}${checkpoint.label ? ` ("${checkpoint.label}")` : ""}.`,
      );
    }
    case "list": {
      const checkpoints = session.listCheckpoints();
      return ok(
        { checkpoints },
        checkpoints.length === 0
          ? "No checkpoints."
          : `${count(checkpoints.length, "checkpoint")}: ${checkpoints
              .map(
                (checkpoint) =>
                  `${checkpoint.id} (revision ${checkpoint.revision}${checkpoint.label ? `, "${checkpoint.label}"` : ""})`,
              )
              .join(", ")}.`,
      );
    }
    default: {
      if (args.id === undefined)
        return refused(
          [
            {
              operationIndex: -1,
              path: "/id",
              code: "required",
              message: "restore needs the checkpoint id",
            },
          ],
          'Pass the id of the checkpoint to restore; list them with action "list".',
        );
      if (!session.listCheckpoints().some((item) => item.id === args.id))
        return refused(
          [
            {
              operationIndex: -1,
              path: "/id",
              code: "unknown-checkpoint",
              message: `No checkpoint ${args.id}`,
            },
          ],
          `No checkpoint ${args.id}; list them with action "list".`,
        );
      const receipt = await session.restoreCheckpoint(args.id, {
        ...(options.expectedRevision === undefined
          ? {}
          : { expectedRevision: options.expectedRevision }),
        ...(options.signal ? { signal: options.signal } : {}),
      });
      return ok(
        receipt,
        receipt.changedPages.length === 0
          ? `Already at checkpoint ${args.id}; revision ${receipt.revision}.`
          : `Restored checkpoint ${args.id}${receipt.removedIds.length > 0 ? `, removing ${count(receipt.removedIds.length, "element")}` : ""}${receipt.createdIds.length > 0 ? `, bringing back ${count(receipt.createdIds.length, "element")}` : ""}; revision ${receipt.revision}.`,
      );
    }
  }
}

/**
 * What a batch did, in the format's words: one sentence per operation with
 * the id it created when the receipt names one per creating operation,
 * then what else the receipt reports.
 */
export function describeReceipt(
  format: EditableFormat,
  operations: readonly EditOperation[],
  receipt: EditReceipt,
): string {
  const pages = noun(format);
  const creating = operations.filter((operation) => CREATING.has(operation.op));
  const idPer =
    creating.length > 0 && receipt.createdIds.length === creating.length;
  let created = 0;
  const sentences = operations.map((operation) => {
    const id =
      idPer && CREATING.has(operation.op)
        ? receipt.createdIds[created++]
        : undefined;
    return describeOperation(format, pages, operation, id);
  });
  if (!idPer && receipt.createdIds.length > 0)
    sentences.push(`Created ${receipt.createdIds.join(", ")}`);
  if (receipt.removedIds.length > 0)
    sentences.push(
      receipt.removedIds.length <= 5
        ? `Removed ${receipt.removedIds.join(", ")}`
        : `Removed ${count(receipt.removedIds.length, "element")}`,
    );
  if (receipt.remappedIds && Object.keys(receipt.remappedIds).length > 0)
    sentences.push(
      `Renamed ${Object.entries(receipt.remappedIds)
        .map(([from, to]) => `${from} → ${to}`)
        .join(", ")}`,
    );
  for (const warning of receipt.warnings)
    sentences.push(`Warning (${warning.code}): ${warning.message}`);
  const body = sentences.map((sentence) => `${sentence}.`).join(" ");
  if (receipt.dryRun)
    return `Preview, nothing changed: ${body} Would repaint ${pageList(receipt.changedPages, pages)}; ${count(receipt.pageCount, pages)} after.`;
  return `${body} Revision ${receipt.revision}; ${receipt.changedPages.length > 0 ? `changed ${pageList(receipt.changedPages, pages)}` : `no ${pages} changed`}; ${count(receipt.pageCount, pages)}.`;
}

const CREATING = new Set([
  "insertTextBox",
  "insertParagraph",
  "insertImage",
  "insertShape",
  "insertTable",
  "insertPage",
  "insertSlide",
  "duplicateSlide",
]);

function describeOperation(
  format: EditableFormat,
  pages: string,
  operation: EditOperation,
  id: string | undefined,
): string {
  const fields = operation as EditOperation & Record<string, unknown>;
  const at = (): string =>
    typeof fields.pageIndex === "number"
      ? ` on ${pages} ${fields.pageIndex + 1}`
      : "";
  const next = (): string =>
    typeof fields.before === "string"
      ? ` before ${fields.before}`
      : typeof fields.after === "string"
        ? ` after ${fields.after}`
        : "";
  const tag = id ? ` (${id})` : "";
  const target = targetName(fields.target);
  const ranged = fields.range ? " a range of" : "";
  const rows = (): string => {
    const grid = fields.rows;
    if (!Array.isArray(grid) || grid.length === 0) return "";
    const columns = Array.isArray(grid[0]) ? (grid[0] as unknown[]).length : 0;
    return ` ${grid.length}×${columns}`;
  };
  switch (operation.op) {
    case "insertTextBox":
      return format === "docx"
        ? `Added a paragraph${next()}${tag}`
        : `Added a text box${at()}${tag}`;
    case "insertParagraph":
      return `Added a paragraph${next()}${tag}`;
    case "insertImage":
      return `Added a picture${at()}${next()}${tag}`;
    case "insertShape":
      return `Added a ${typeof fields.shape === "string" ? fields.shape : "shape"}${at()}${tag}`;
    case "insertTable":
      return `Added a${rows()} table${at()}${next()}${tag}`;
    case "replaceText":
      return `Replaced${ranged} the text of ${target}`;
    case "setTextStyle":
      return `Changed the text style of${ranged} ${target}`;
    case "setParagraphStyle":
      return `Changed the paragraph style of ${target}`;
    case "setShapeStyle":
      return `Changed the shape style of ${target}`;
    case "setTableCell":
      return `Set cell ${String(fields.row)},${String(fields.column)} of ${target}`;
    case "moveElement":
      return `Moved ${target}${next()}`;
    case "resizeElement":
      return `Resized ${target}`;
    case "deleteElement":
      return `Deleted ${target}`;
    case "insertPage":
    case "insertSlide":
      return `Added ${pages} ${Number(fields.index) + 1}${tag}`;
    case "duplicateSlide":
      return `Duplicated ${pages} ${Number(fields.pageIndex) + 1}${tag}`;
    case "deletePage":
    case "deleteSlide":
      return `Deleted ${pages} ${Number(fields.pageIndex) + 1}`;
    case "movePage":
    case "moveSlide":
      return `Moved ${pages} ${Number(fields.from) + 1} to ${Number(fields.to) + 1}`;
    case "rotatePage":
      return `Rotated ${pages} ${Number(fields.pageIndex) + 1} to ${String(fields.rotation)}°`;
    default:
      return `Applied ${operation.op}${target === "the element" ? "" : ` to ${target}`}${tag}`;
  }
}

function targetName(target: unknown): string {
  if (typeof target !== "string") return "the element";
  const reference = /^\$(\d+)$/.exec(target);
  return reference
    ? `the element created by operation ${Number(reference[1]) + 1}`
    : target;
}

function pageList(pages: readonly number[], pageNoun: string): string {
  if (pages.length === 0) return `no ${pageNoun}`;
  if (pages.length > 6)
    return `${count(pages.length, pageNoun)} (${pages[0]! + 1}–${pages.at(-1)! + 1})`;
  return `${pageNoun}${pages.length === 1 ? "" : "s"} ${pages.map((page) => page + 1).join(", ")}`;
}

function count(value: number, word: string): string {
  return `${value} ${word}${value === 1 ? "" : "s"}`;
}
