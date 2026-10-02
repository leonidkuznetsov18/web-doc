import type { ResourceLimits } from "../../contracts.js";
import type {
  EditableFormat,
  EditElement,
  EditState,
  ElementQuery,
  ReadItem,
  ReadOptions,
  ReadResult,
} from "../types.js";
import type {
  DescribeOptions,
  DocumentDescription,
  OutlineNode,
  OutlineOptions,
  OutlineResult,
} from "./types.js";

/*
 * The outline is `getElements()` reshaped for a prompt: nested where the
 * format nests (a table's cell paragraphs, a group's members, a paragraph's
 * inline pictures), numbered in reading order, with a short text and a
 * human label per node and nothing a model cannot use (geometry, styles).
 * The description is the outline as text in a fixed line grammar, so a host
 * pastes it into a prompt without a formatter of its own.
 */

const DEFAULT_TEXT_CHARS = 160;
const DEFAULT_DESCRIBE_CHARS = 50_000;

/** What the outline needs of a session: its format, its page count and its elements. */
export interface OutlineSource {
  readonly format: EditableFormat;
  readonly state: EditState;
  getElements(
    query?: ElementQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<EditElement>>;
}

/** The outline of a session's document; one `getElements()` read, reshaped. */
export async function readOutline(
  session: OutlineSource,
  limits: ResourceLimits,
  options: OutlineOptions = {},
): Promise<OutlineResult> {
  const { pageRange, signal } = options;
  // A single page is read as such: cheaper for formats that load pages on demand.
  const query: ElementQuery =
    pageRange && pageRange[0] === pageRange[1]
      ? { pageIndex: pageRange[0] }
      : {};
  const read = await session.getElements(query, signal ? { signal } : {});
  const built = buildOutline(
    session.format,
    read.items,
    options,
    limits.maxOutlineNodes,
  );
  return Object.freeze({
    sessionId: read.sessionId,
    revision: read.revision,
    items: built.nodes,
    nodeCount: built.nodeCount,
    truncated: built.truncated,
  });
}

/** The description of a session's document: the outline rendered within a character budget. */
export async function readDescription(
  session: OutlineSource,
  limits: ResourceLimits,
  options: DescribeOptions = {},
): Promise<ReadItem<DocumentDescription>> {
  const { maxChars, ...outlineOptions } = options;
  const outline = await readOutline(session, limits, outlineOptions);
  const budget = Math.max(
    1,
    Math.min(maxChars ?? DEFAULT_DESCRIBE_CHARS, limits.maxDescribeChars),
  );
  return Object.freeze({
    sessionId: outline.sessionId,
    revision: outline.revision,
    item: renderDescription(
      session.format,
      session.state.pageCount,
      outline,
      budget,
      limits.maxOutlineNodes,
    ),
  });
}

export interface BuiltOutline {
  readonly nodes: readonly OutlineNode[];
  readonly nodeCount: number;
  readonly truncated: boolean;
}

/**
 * Nests elements by `parentId`, numbers them in reading order, keeps the
 * nodes the options ask for (a container stays when a descendant is kept)
 * and cuts the result at `maxNodes`. Ordinals are assigned before any
 * filtering, so "3.2" names the same paragraph whatever the query.
 */
export function buildOutline(
  format: EditableFormat,
  elements: readonly EditElement[],
  options: Pick<OutlineOptions, "pageRange" | "kinds" | "maxTextChars">,
  maxNodes: number,
): BuiltOutline {
  const maxText = Math.max(0, options.maxTextChars ?? DEFAULT_TEXT_CHARS);
  const ids = new Set(elements.map((element) => element.id));
  const childrenOf = new Map<string, EditElement[]>();
  const roots: EditElement[] = [];
  for (const element of elements) {
    const parent = element.parentId;
    if (parent === undefined || parent === element.id || !ids.has(parent)) {
      roots.push(element);
      continue;
    }
    const siblings = childrenOf.get(parent);
    if (siblings) siblings.push(element);
    else childrenOf.set(parent, [element]);
  }

  const visited = new Set<string>();
  const toNode = (element: EditElement, ordinal: string): OutlineNode => {
    visited.add(element.id);
    const children = (childrenOf.get(element.id) ?? [])
      .filter((child) => !visited.has(child.id))
      .map((child, index) => toNode(child, `${ordinal}.${index + 1}`));
    return nodeOf(format, element, ordinal, maxText, children);
  };
  const tree = roots.map((element, index) => toNode(element, `${index + 1}`));
  // Elements whose parents form a cycle are never reached from a root;
  // they are listed after the roots rather than dropped.
  let extra = tree.length;
  for (const element of elements)
    if (!visited.has(element.id)) {
      extra += 1;
      tree.push(toNode(element, `${extra}`));
    }

  const kept = prune(tree, keepPredicate(options));
  const cut = cutNodes(kept, maxNodes);
  return {
    nodes: cut.nodes,
    nodeCount: cut.count,
    truncated: cut.truncated,
  };
}

function nodeOf(
  format: EditableFormat,
  element: EditElement,
  ordinal: string,
  maxText: number,
  children: readonly OutlineNode[],
): OutlineNode {
  const text = element.text;
  const label = labelOf(format, element);
  const extra = element as {
    readonly readOnlyReason?: string;
    readonly hidden?: boolean;
  };
  return Object.freeze({
    id: element.id,
    ordinal,
    kind: element.kind,
    pageIndex: element.pageIndex,
    ...(text === undefined ? {} : { text: cutText(text, maxText) }),
    textLength: text?.length ?? 0,
    truncated: text !== undefined && text.length > maxText,
    ...(label === undefined ? {} : { label }),
    ...(extra.readOnlyReason === undefined
      ? {}
      : { readOnlyReason: extra.readOnlyReason }),
    ...(extra.hidden ? { hidden: true } : {}),
    operations: Object.freeze([...element.operations]),
    ...(children.length > 0 ? { children: Object.freeze(children) } : {}),
  });
}

/** The first `max` code units of `text`, never ending inside a surrogate pair. */
function cutText(text: string, max: number): string {
  if (text.length <= max) return text;
  let end = max;
  const last = text.charCodeAt(end - 1);
  if (end > 0 && last >= 0xd800 && last <= 0xdbff) end -= 1;
  return text.slice(0, end);
}

type Keep = (node: OutlineNode) => boolean;

function keepPredicate(
  options: Pick<OutlineOptions, "pageRange" | "kinds">,
): Keep | undefined {
  const { pageRange, kinds } = options;
  if (!pageRange && !kinds) return undefined;
  const wanted = kinds ? new Set(kinds) : undefined;
  return (node) =>
    (!wanted || wanted.has(node.kind)) &&
    (!pageRange ||
      (node.pageIndex >= pageRange[0] && node.pageIndex <= pageRange[1]));
}

/** Keeps the nodes `keep` accepts and every ancestor of one, in order. */
function prune(
  nodes: readonly OutlineNode[],
  keep: Keep | undefined,
): readonly OutlineNode[] {
  if (!keep) return nodes;
  const out: OutlineNode[] = [];
  for (const node of nodes) {
    const children = node.children ? prune(node.children, keep) : [];
    if (!keep(node) && children.length === 0) continue;
    out.push(
      children.length === node.children?.length
        ? node
        : withChildren(node, children),
    );
  }
  return out;
}

/** The first `max` nodes in pre-order; the rest are cut, subtrees included. */
function cutNodes(
  nodes: readonly OutlineNode[],
  max: number,
): { nodes: readonly OutlineNode[]; count: number; truncated: boolean } {
  let count = 0;
  let truncated = false;
  const take = (list: readonly OutlineNode[]): readonly OutlineNode[] => {
    const out: OutlineNode[] = [];
    for (const node of list) {
      if (count >= max) {
        truncated = true;
        break;
      }
      count += 1;
      if (!node.children) {
        out.push(node);
        continue;
      }
      const children = take(node.children);
      out.push(
        children.length === node.children.length
          ? node
          : withChildren(node, children),
      );
      // A cut inside the children leaves nothing for the siblings.
      if (truncated) break;
    }
    return Object.freeze(out);
  };
  const kept = take(nodes);
  return { nodes: kept, count, truncated };
}

/** `node` with `children`, the key left out when there are none. */
function withChildren(
  node: OutlineNode,
  children: readonly OutlineNode[],
): OutlineNode {
  const { children: _dropped, ...rest } = node;
  return Object.freeze(
    children.length > 0 ? { ...rest, children: Object.freeze(children) } : rest,
  );
}

/**
 * What a person calls the element, when the format knows: a shape's name
 * or placeholder in a deck, a paragraph's style in a Word document, a
 * table's size everywhere.
 */
export function labelOf(
  format: EditableFormat,
  element: EditElement,
): string | undefined {
  const known = element as {
    readonly table?: { readonly rows: readonly (readonly string[])[] };
    readonly name?: string;
    readonly placeholder?: { readonly type: string };
    readonly paragraphStyle?: {
      readonly styleId?: string;
      readonly numbering?: { readonly level: number };
    };
  };
  const size = known.table ? tableSize(known.table.rows) : undefined;
  if (format === "pptx") {
    const name = known.name?.trim();
    if (name) return size ? `${name} (${size})` : name;
    if (known.placeholder) return placeholderLabel(known.placeholder.type);
  }
  if (format === "docx" && element.kind === "paragraph") {
    const style = known.paragraphStyle;
    const styleId = style?.styleId;
    const name =
      styleId && styleId !== "Normal" ? styleName(styleId) : undefined;
    if (name) return name;
    if (style?.numbering)
      return `List item (level ${style.numbering.level + 1})`;
    return undefined;
  }
  if (size) return `Table (${size})`;
  return undefined;
}

function tableSize(rows: readonly (readonly string[])[]): string {
  const columns = rows.reduce((max, row) => Math.max(max, row.length), 0);
  return `${rows.length}×${columns}`;
}

const PLACEHOLDER_LABELS: Readonly<Record<string, string>> = {
  title: "Title",
  ctrTitle: "Title",
  subTitle: "Subtitle",
  body: "Body",
  obj: "Content",
  pic: "Picture",
  tbl: "Table",
  chart: "Chart",
  dgm: "Diagram",
  media: "Media",
  clipArt: "Clip art",
  dt: "Date",
  ftr: "Footer",
  sldNum: "Slide number",
  hdr: "Header",
  sldImg: "Slide image",
};

function placeholderLabel(type: string): string {
  return PLACEHOLDER_LABELS[type] ?? `${type} placeholder`;
}

/** "Heading1" → "Heading 1", "ListParagraph" → "List Paragraph". */
function styleName(styleId: string): string {
  return styleId
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .replace(/([A-Za-z])(\d)/g, "$1 $2");
}

/**
 * The description's grammar, one line per node:
 *
 *     <format>: <n> pages|slides, <m> elements
 *     [<id>] page <p> <kind> "<label>" (<flags>): <text>…
 *       [<child id>] page <p> <kind>: <text>
 *
 * The page is left out while the node has none, the label and the flags
 * while the node has none, the text while it is empty. Tabs in the text
 * become ` | `, line breaks ` ⏎ `; a cut text ends in `…`. Children are
 * indented two spaces per level. The budget cuts whole lines and the last
 * line then says how many were left out.
 */
export function renderDescription(
  format: EditableFormat,
  pageCount: number,
  outline: BuiltOutline | OutlineResult,
  maxChars: number,
  maxNodes: number,
): DocumentDescription {
  const nodes = "items" in outline ? outline.items : outline.nodes;
  const noun = format === "pptx" ? "slide" : "page";
  const lines: string[] = [
    `${format}: ${count(pageCount, noun)}, ${count(outline.nodeCount, "element")}`,
  ];
  const push = (node: OutlineNode, depth: number): void => {
    lines.push(`${"  ".repeat(depth)}${describeNode(node, noun)}`);
    for (const child of node.children ?? []) push(child, depth + 1);
  };
  for (const node of nodes) push(node, 0);

  // Whole lines within the budget; the last line then says what was left
  // out, and makes room for itself by dropping lines from the end.
  let keep = lines.length;
  let text = "";
  let truncated = outline.truncated;
  for (;;) {
    const left = lines.length - keep;
    const tail =
      left > 0
        ? `… (${count(left, "more line")})`
        : outline.truncated
          ? `… (outline cut at ${count(maxNodes, "node")})`
          : undefined;
    const body = lines.slice(0, keep).join("\n");
    text = tail === undefined ? body : body ? `${body}\n${tail}` : tail;
    if (text.length <= maxChars) break;
    if (keep === 0) {
      // Not even the header fits: hand back what does.
      text = lines[0]!.slice(0, maxChars);
      truncated = true;
      break;
    }
    keep -= 1;
    truncated = true;
  }
  return Object.freeze({
    format,
    pageCount,
    elementCount: outline.nodeCount,
    text,
    truncated,
  });
}

function describeNode(node: OutlineNode, noun: string): string {
  const parts = [`[${node.id}]`];
  if (node.pageIndex >= 0) parts.push(`${noun} ${node.pageIndex + 1}`);
  parts.push(node.kind);
  if (node.label !== undefined)
    parts.push(`"${node.label.replaceAll('"', "'")}"`);
  const flags = [
    ...(node.readOnlyReason ? [`read-only: ${node.readOnlyReason}`] : []),
    ...(node.hidden ? ["hidden"] : []),
  ];
  if (flags.length > 0) parts.push(`(${flags.join(", ")})`);
  let line = parts.join(" ");
  if (node.text) line += `: ${foldText(node.text)}${node.truncated ? "…" : ""}`;
  else if (node.truncated) line += ": …";
  return line;
}

/** One line of text: tabs as cell separators, breaks marked, other controls blanked. */
export function foldText(text: string): string {
  return text
    .replaceAll("\r\n", "\n")
    .replaceAll("\t", " | ")
    .replaceAll(/[\n\r\u2028\u2029\v\f]/g, " ⏎ ")
    .replaceAll(/[\u0000-\u0008\u000e-\u001f\u007f]/g, " ");
}

function count(value: number, noun: string): string {
  return `${value} ${noun}${value === 1 ? "" : "s"}`;
}
