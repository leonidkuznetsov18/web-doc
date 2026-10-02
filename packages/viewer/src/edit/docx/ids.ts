import type { XmlElement, XmlPart } from "../ooxml/xml.js";

/*
 * Paragraph ids shared by the display pre-pass and the DOCX edit engine.
 * A paragraph's id is its `w14:paraId` when the file has one; otherwise a
 * deterministic id from its position among the unmarked paragraphs of the
 * document, numbered part by part in one sequence (the main part first,
 * then headers, footers, footnotes and endnotes in relationship order) and
 * skipping every id the document already uses. Both sides walk the same
 * paragraphs in the same order, so the pre-pass's bookmarks and the
 * engine's element ids agree on the original bytes.
 */

export const W_NS =
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";
export const OFFICE_RELATIONSHIPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
/** Story parts whose paragraphs share the id sequence, in the order they are numbered. */
export const STORY_RELATIONSHIP_TYPES = [
  `${OFFICE_RELATIONSHIPS}header`,
  `${OFFICE_RELATIONSHIPS}footer`,
  `${OFFICE_RELATIONSHIPS}footnotes`,
  `${OFFICE_RELATIONSHIPS}endnotes`,
] as const;

/** Prefix of the hidden bookmarks that carry paragraph ids. */
export const PARAGRAPH_BOOKMARK_PREFIX = "_wd";
/** First bookmark id the pre-pass uses, above what Word writes. */
export const BOOKMARK_ID_BASE = 7_000_000;

/** First generated id; Word keeps `w14:paraId` below 0x80000000. */
const ID_BASE = 0x1a000000;
const ID_STEP = 0x9e37;

/** Ids and bookmark numbers shared by every story part of one document. */
export interface DocxIdState {
  readonly taken: Set<string>;
  generated: number;
  bookmarkId: number;
}

export function newIdState(): DocxIdState {
  return { taken: new Set(), generated: 0, bookmarkId: BOOKMARK_ID_BASE };
}

/**
 * The paragraph id generated for the `index`-th paragraph without one, in
 * document order, skipping ids the document already uses.
 */
export function generatedParagraphId(
  index: number,
  taken: ReadonlySet<string>,
): string {
  let candidate = (ID_BASE + index * ID_STEP) % 0x80000000;
  for (;;) {
    const value = candidate.toString(16).toUpperCase().padStart(8, "0");
    if (!taken.has(value)) return value;
    candidate = (candidate + 1) % 0x80000000;
  }
}

/** Every `w:p` of a part in document order, nested ones included. */
export function paragraphsOf(part: XmlPart): readonly XmlElement[] {
  return part.findAll("p").filter((node) => node.namespace === W_NS);
}

/** Collects the ids and bookmark numbers a part already uses. */
export function collectIds(part: XmlPart, state: DocxIdState): void {
  for (const node of part.findAll("bookmarkStart")) {
    if (node.namespace !== W_NS) continue;
    const id = Number(part.attribute(node, "w:id") ?? NaN);
    if (Number.isInteger(id) && id >= state.bookmarkId)
      state.bookmarkId = id + 1;
    const name = part.attribute(node, "w:name") ?? "";
    if (name.startsWith(PARAGRAPH_BOOKMARK_PREFIX))
      state.taken.add(
        name.slice(PARAGRAPH_BOOKMARK_PREFIX.length).toUpperCase(),
      );
  }
  for (const paragraph of paragraphsOf(part)) {
    const existing = part.attribute(paragraph, "w14:paraId");
    if (existing) state.taken.add(existing.toUpperCase());
  }
}

export interface ParagraphId {
  readonly paragraph: XmlElement;
  /** Eight upper-case hex digits. */
  readonly id: string;
  /** True when the file carries the id as `w14:paraId`. */
  readonly authored: boolean;
}

/**
 * The id of every paragraph of a part, numbering the unmarked ones from
 * where `state` stands. Call it for the parts in numbering order after
 * `collectIds` has seen all of them.
 */
export function assignParagraphIds(
  part: XmlPart,
  state: DocxIdState,
): readonly ParagraphId[] {
  const out: ParagraphId[] = [];
  for (const paragraph of paragraphsOf(part)) {
    const authored = part.attribute(paragraph, "w14:paraId");
    if (authored) {
      out.push({ paragraph, id: authored.toUpperCase(), authored: true });
      continue;
    }
    const id = generatedParagraphId(state.generated, state.taken);
    state.taken.add(id);
    state.generated += 1;
    out.push({ paragraph, id, authored: false });
  }
  return out;
}

/** The paragraph id an element id is built on: `p:X`, `tbl:X`, `img:X.n`, `other:X.n`. */
export function paragraphIdOfElement(elementId: string): string | undefined {
  const match = /^(?:p|tbl|img|other):([0-9A-Fa-f]{8})(?:\.\d+)?$/.exec(
    elementId,
  );
  return match ? match[1]!.toUpperCase() : undefined;
}

/**
 * A fresh id for the `index`-th paragraph an operation creates: derived
 * from the batch's state id and the operation's position, so a replay of
 * the same batches issues the same ids, and skipping any id in use.
 */
export function freshParagraphId(
  stateId: number,
  operationIndex: number,
  index: number,
  taken: ReadonlySet<string>,
): string {
  const base =
    (0x2a000000 + stateId * 0x10000 + operationIndex * 0x100 + index) %
    0x80000000;
  let candidate = base;
  for (;;) {
    const value = candidate.toString(16).toUpperCase().padStart(8, "0");
    if (!taken.has(value)) return value;
    candidate = (candidate + 1) % 0x80000000;
  }
}
