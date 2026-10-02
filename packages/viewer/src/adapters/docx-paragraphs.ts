/*
 * The paragraph bridge for DOCX text runs. The renderer lays out from its
 * own model and tags every run with a `source`: the story it came from
 * (`body`, a header or footer, a note, a text box), the story instance and
 * the path of block indices that leads to the paragraph (a table adds its
 * row and cell indices, the last entry is the run index). The display
 * pre-pass marks every `w:p` of the file with a hidden `_wd<id>` bookmark,
 * and the engine keeps bookmark names on its model paragraphs, so a run's
 * source leads to a model paragraph whose bookmark names its `w:p`.
 *
 * Instance names follow the engine: `body` for the body, `default`, `first`
 * or `even` for the document's headers and footers and `section:<i>:<kind>`
 * for those a section break carries (`i` being the break's index in the
 * body), the note id for footnotes and endnotes, and for a text box the
 * `story:instance:path` key of the shape run that holds it.
 */

export type DocxStory =
  "body" | "header" | "footer" | "footnote" | "endnote" | "textbox";

export interface DocxRunSource {
  readonly story: DocxStory | string;
  readonly storyInstance: string;
  readonly path: readonly number[];
}

export interface DocxModelParagraph {
  readonly type: "paragraph";
  readonly paragraphId?: string;
  readonly bookmarks?: readonly string[];
  readonly runs?: readonly DocxModelRun[];
}

export interface DocxModelRun {
  readonly type?: string;
  /** The blocks a text box shape holds. */
  readonly textBoxContent?: readonly DocxModelBlock[];
}

export interface DocxModelTable {
  readonly type: "table";
  readonly rows: readonly {
    readonly cells: readonly { readonly content: readonly DocxModelBlock[] }[];
  }[];
}

export interface DocxModelHeaderFooter {
  readonly body: readonly DocxModelBlock[];
}

export interface DocxModelHeadersFooters {
  readonly default?: DocxModelHeaderFooter | null;
  readonly first?: DocxModelHeaderFooter | null;
  readonly even?: DocxModelHeaderFooter | null;
}

export interface DocxModelSectionBreak {
  readonly type: "sectionBreak";
  readonly headers?: DocxModelHeadersFooters;
  readonly footers?: DocxModelHeadersFooters;
}

export interface DocxModelPageBreak {
  readonly type: "pageBreak";
}

export type DocxModelBlock =
  | DocxModelParagraph
  | DocxModelTable
  | DocxModelSectionBreak
  | DocxModelPageBreak
  | { readonly type: string };

export interface DocxModelNote {
  readonly id: string;
  readonly content: readonly DocxModelBlock[];
}

/** The part of the engine's document model the bridge walks. */
export interface DocxModelDocument {
  readonly body: readonly DocxModelBlock[];
  readonly headers?: DocxModelHeadersFooters;
  readonly footers?: DocxModelHeadersFooters;
  readonly footnotes?: readonly DocxModelNote[];
  readonly endnotes?: readonly DocxModelNote[];
}

/** Resolves a run's source to the id of its `w:p`, or `undefined`. */
export type DocxParagraphIdResolver = (
  source: DocxRunSource | undefined,
) => string | undefined;

/** The name prefix of the pre-pass bookmarks that carry paragraph ids. */
const BOOKMARK_PREFIX = "_wd";
const HEADER_KINDS = ["default", "first", "even"] as const;

/**
 * A resolver over one document model. Results are cached per source, so a
 * page's runs (many per paragraph) cost one walk per paragraph.
 */
export function createDocxParagraphIdResolver(
  model: DocxModelDocument | undefined,
): DocxParagraphIdResolver {
  if (!model || !Array.isArray(model.body)) return () => undefined;
  const cache = new Map<string, string | undefined>();
  return (source) => {
    if (!source || !Array.isArray(source.path)) return undefined;
    const key = `${source.story}\u0000${source.storyInstance}\u0000${source.path.join(",")}`;
    if (cache.has(key)) return cache.get(key);
    let id: string | undefined;
    try {
      const container = containerOf(model, source.story, source.storyInstance);
      const located = container && paragraphAt(container, source.path);
      id = located && paragraphIdAt(located.blocks, located.index);
    } catch {
      id = undefined;
    }
    cache.set(key, id);
    return id;
  };
}

/** The id a model paragraph carries: its own `paragraphId` or its bookmark. */
export function paragraphIdOf(
  paragraph: DocxModelParagraph,
): string | undefined {
  if (typeof paragraph.paragraphId === "string" && paragraph.paragraphId)
    return paragraph.paragraphId;
  for (const name of paragraph.bookmarks ?? [])
    if (typeof name === "string" && name.startsWith(BOOKMARK_PREFIX))
      return name.slice(BOOKMARK_PREFIX.length);
  return undefined;
}

interface LocatedParagraph {
  readonly blocks: readonly DocxModelBlock[];
  readonly index: number;
  readonly paragraph: DocxModelParagraph;
  /** Path entries after the paragraph: the run index, when present. */
  readonly rest: readonly number[];
}

function containerOf(
  model: DocxModelDocument,
  story: string,
  instance: string,
): readonly DocxModelBlock[] | undefined {
  switch (story) {
    case "body":
      return model.body;
    case "header":
    case "footer":
      return headerFooterBody(model, story, instance);
    case "footnote":
      return model.footnotes?.find((note) => note.id === instance)?.content;
    case "endnote":
      return model.endnotes?.find((note) => note.id === instance)?.content;
    case "textbox":
      return textBoxContent(model, instance);
    default:
      return undefined;
  }
}

function headerFooterBody(
  model: DocxModelDocument,
  story: "header" | "footer",
  instance: string,
): readonly DocxModelBlock[] | undefined {
  const kindOf = (
    value: string | undefined,
  ): (typeof HEADER_KINDS)[number] | undefined =>
    HEADER_KINDS.find((kind) => kind === value);
  const sectionMatch = /^section:(\d+):([a-z]+)$/.exec(instance);
  if (sectionMatch) {
    const block = model.body[Number(sectionMatch[1])];
    const kind = kindOf(sectionMatch[2]);
    if (!block || block.type !== "sectionBreak" || !kind) return undefined;
    const set = (block as DocxModelSectionBreak)[
      story === "header" ? "headers" : "footers"
    ];
    return set?.[kind]?.body;
  }
  const kind = kindOf(instance);
  if (!kind) return undefined;
  const set = story === "header" ? model.headers : model.footers;
  return set?.[kind]?.body;
}

/**
 * A text box's instance is the `story:instance:path` key of the shape run
 * that holds it; the path's last entry is that run's index in its paragraph.
 */
function textBoxContent(
  model: DocxModelDocument,
  instance: string,
): readonly DocxModelBlock[] | undefined {
  const first = instance.indexOf(":");
  const last = instance.lastIndexOf(":");
  if (first < 0 || last <= first) return undefined;
  const story = instance.slice(0, first);
  const hostInstance = instance.slice(first + 1, last);
  const path = instance
    .slice(last + 1)
    .split(".")
    .map((entry) => Number(entry));
  if (path.length === 0 || path.some((entry) => !Number.isInteger(entry)))
    return undefined;
  const container = containerOf(model, story, hostInstance);
  const located = container && paragraphAt(container, path);
  const runIndex = located?.rest[0];
  if (!located || runIndex === undefined) return undefined;
  const run = located.paragraph.runs?.[runIndex];
  return Array.isArray(run?.textBoxContent) ? run.textBoxContent : undefined;
}

/** Follows a source path through blocks and table cells to its paragraph. */
function paragraphAt(
  container: readonly DocxModelBlock[],
  path: readonly number[],
): LocatedParagraph | undefined {
  let blocks = container;
  let at = 0;
  while (at < path.length) {
    const index = path[at]!;
    const node = blocks[index];
    if (!node) return undefined;
    if (node.type === "paragraph")
      return {
        blocks,
        index,
        paragraph: node as DocxModelParagraph,
        rest: path.slice(at + 1),
      };
    if (node.type !== "table") return undefined;
    const cell = (node as DocxModelTable).rows[path[at + 1] ?? -1]?.cells[
      path[at + 2] ?? -1
    ];
    if (!cell || !Array.isArray(cell.content)) return undefined;
    blocks = cell.content;
    at += 3;
  }
  return undefined;
}

/**
 * The id of the paragraph at `index`. The engine splits a paragraph that
 * holds a page break into two model paragraphs around a hoisted page break,
 * and only the first keeps the bookmark, so an unmarked paragraph right
 * after a page break takes the id of the paragraph before that break.
 * Any other unmarked paragraph has no id: guessing would name the wrong
 * `w:p`.
 */
function paragraphIdAt(
  blocks: readonly DocxModelBlock[],
  index: number,
): string | undefined {
  let at = index;
  for (;;) {
    const node = blocks[at];
    if (!node || node.type !== "paragraph") return undefined;
    const id = paragraphIdOf(node as DocxModelParagraph);
    if (id !== undefined) return id;
    if (
      at >= 2 &&
      blocks[at - 1]?.type === "pageBreak" &&
      blocks[at - 2]?.type === "paragraph"
    )
      at -= 2;
    else return undefined;
  }
}
