import { ViewerError } from "../../errors.js";
import type { OoxmlPackage } from "../ooxml/package.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import {
  assignParagraphIds,
  collectIds,
  newIdState,
  OFFICE_RELATIONSHIPS,
  STORY_RELATIONSHIP_TYPES,
  W_NS,
} from "./ids.js";
import { DocxStyles } from "./style.js";
import { readParagraphText, type ParagraphText } from "./text.js";
import type { DocxReadOnlyReason } from "./types.js";

/*
 * The block index of the body story: every paragraph and table of
 * `w:body` in document order (block-level `w:sdt` and `w:customXml`
 * unwrapped, the paragraphs of table cells included), each with its id,
 * its text model and the XML it came from. Built from the package at a
 * revision and rebuilt when the revision changes. Headers, footers and
 * notes are not listed in this module.
 */

export interface ParagraphRecord {
  readonly kind: "paragraph";
  /** Eight upper-case hex digits: the `w14:paraId` or the generated id. */
  readonly id: string;
  /** "p:<id>". */
  readonly elementId: string;
  readonly node: XmlElement;
  readonly pPr?: XmlElement;
  /** `w:pPr/w:sectPr`: the paragraph ends a section. */
  readonly sectPr?: XmlElement;
  readonly text: ParagraphText;
  /** The table whose cell holds the paragraph, when any. */
  readonly table?: TableRecord;
  /** The block-level wrapper (`w:sdt`, `w:customXml`) the paragraph sits in, when any. */
  readonly wrapper?: XmlElement;
  readonly readOnlyReason?: DocxReadOnlyReason;
  /** Inline pictures and other objects of the paragraph, in order. */
  readonly inlines: readonly InlineRecord[];
}

export interface CellRecord {
  readonly node: XmlElement;
  readonly paragraphs: readonly ParagraphRecord[];
}

export interface TableRecord {
  readonly kind: "table";
  /** The id of the table's first paragraph. */
  readonly id: string;
  /** "tbl:<id>". */
  readonly elementId: string;
  readonly node: XmlElement;
  readonly rows: readonly (readonly CellRecord[])[];
  readonly wrapper?: XmlElement;
}

export interface InlineRecord {
  readonly kind: "image" | "other";
  /** "img:<paragraph id>.<n>" or "other:<paragraph id>.<n>". */
  readonly elementId: string;
  readonly paragraph: ParagraphRecord;
  /** The `w:drawing`, `w:object`, `w:pict`, math or alternate-content node. */
  readonly node: XmlElement;
  /** The `w:r` holding the node, when it sits in a run. */
  readonly run?: XmlElement;
  /** For an inline picture: `wp:extent` in EMU. */
  readonly extent?: { readonly cx: number; readonly cy: number };
}

export type BlockRecord = ParagraphRecord | TableRecord;
export type AnyRecord = ParagraphRecord | TableRecord | InlineRecord;

export class DocxModel {
  private constructor(
    readonly revision: number,
    readonly mainPart: string,
    readonly document: XmlPart,
    readonly body: XmlElement,
    /** The body's own `w:sectPr`, when present. */
    readonly bodySectPr: XmlElement | undefined,
    /** Top-level blocks of the body in order; cell paragraphs hang off their table. */
    readonly blocks: readonly BlockRecord[],
    /** Every listed record in document order: a table before its cell paragraphs, a paragraph before its inlines. */
    readonly records: readonly AnyRecord[],
    readonly byId: ReadonlyMap<string, AnyRecord>,
    /** Id of every `w:p` of the main part, listed or not. */
    readonly paragraphIds: ReadonlyMap<XmlElement, string>,
    readonly styles: DocxStyles,
  ) {}

  static async load(
    pkg: OoxmlPackage,
    signal?: AbortSignal,
  ): Promise<DocxModel> {
    const root = await pkg.relationships("/", signal);
    const main = root.byType(`${OFFICE_RELATIONSHIPS}officeDocument`)[0]
      ?.targetPart;
    if (!main || !pkg.has(main))
      throw new ViewerError("invalid-file", "The package has no document part");
    const document = await pkg.xml(main, signal);
    const body = document.root.children.find(
      (child) => child.local === "body" && child.namespace === W_NS,
    );
    if (!body)
      throw new ViewerError("invalid-file", "The document part has no body");
    // Ids are numbered across every story part, the main part first.
    const relationships = await pkg.relationships(main, signal);
    const state = newIdState();
    collectIds(document, state);
    for (const type of STORY_RELATIONSHIP_TYPES)
      for (const item of relationships.byType(type))
        if (item.targetPart && pkg.has(item.targetPart))
          collectIds(await pkg.xml(item.targetPart, signal), state);
    const paragraphIds = new Map<XmlElement, string>();
    for (const entry of assignParagraphIds(document, state))
      paragraphIds.set(entry.paragraph, entry.id);
    const styles = await DocxStyles.load(pkg, main, signal);
    const builder = new Builder(document, paragraphIds);
    let bodySectPr: XmlElement | undefined;
    for (const child of body.children) {
      if (child.namespace === W_NS && child.local === "sectPr") {
        bodySectPr = child;
        continue;
      }
      builder.block(child, undefined, undefined);
    }
    return new DocxModel(
      pkg.revision,
      main,
      document,
      body,
      bodySectPr,
      builder.blocks,
      builder.records,
      builder.byId,
      paragraphIds,
      styles,
    );
  }

  get paragraphs(): readonly ParagraphRecord[] {
    return this.records.filter(
      (record): record is ParagraphRecord => record.kind === "paragraph",
    );
  }
}

class Builder {
  readonly blocks: BlockRecord[] = [];
  readonly records: AnyRecord[] = [];
  readonly byId = new Map<string, AnyRecord>();

  constructor(
    private readonly part: XmlPart,
    private readonly ids: ReadonlyMap<XmlElement, string>,
  ) {}

  /** A child of the body or of a cell; wrappers are unwrapped. */
  block(
    node: XmlElement,
    table: TableRecord | undefined,
    wrapper: XmlElement | undefined,
    into?: ParagraphRecord[],
  ): void {
    if (node.namespace !== W_NS) return;
    switch (node.local) {
      case "p": {
        const record = this.paragraph(node, table, wrapper);
        if (into) into.push(record);
        else this.blocks.push(record);
        return;
      }
      case "tbl":
        // A table inside a cell is left out: its structure is not edited here.
        if (!table) {
          const record = this.table(node, wrapper);
          if (record) this.blocks.push(record);
        }
        return;
      case "sdt": {
        const content = node.children.find(
          (child) => child.local === "sdtContent" && child.namespace === W_NS,
        );
        if (content)
          for (const child of content.children)
            this.block(child, table, node, into);
        return;
      }
      case "customXml":
        for (const child of node.children) this.block(child, table, node, into);
        return;
      default:
        return;
    }
  }

  paragraph(
    node: XmlElement,
    table: TableRecord | undefined,
    wrapper: XmlElement | undefined,
  ): ParagraphRecord {
    const id = this.ids.get(node) ?? "00000000";
    const pPr = node.children.find(
      (child) => child.local === "pPr" && child.namespace === W_NS,
    );
    const sectPr = pPr?.children.find(
      (child) => child.local === "sectPr" && child.namespace === W_NS,
    );
    const text = readParagraphText(this.part, node);
    const fieldOnly =
      text.items.some((item) => item.kind === "field") &&
      !text.items.some((item) => item.kind === "text" && item.text.length > 0);
    const readOnlyReason: DocxReadOnlyReason | undefined = text.tracked
      ? "tracked-changes"
      : sectPr
        ? "section-break"
        : fieldOnly
          ? "unsupported-content"
          : undefined;
    const inlines: InlineRecord[] = [];
    const record: ParagraphRecord = {
      kind: "paragraph",
      id,
      elementId: `p:${id}`,
      node,
      ...(pPr ? { pPr } : {}),
      ...(sectPr ? { sectPr } : {}),
      text,
      ...(table ? { table } : {}),
      ...(wrapper ? { wrapper } : {}),
      ...(readOnlyReason ? { readOnlyReason } : {}),
      inlines,
    };
    this.register(record);
    let pictures = 0;
    let others = 0;
    for (const item of text.items) {
      if (item.kind === "picture") {
        const inline = item.child.children.find(
          (child) => child.local === "inline",
        );
        const extentNode = inline?.children.find(
          (child) => child.local === "extent",
        );
        const cx = Number(
          extentNode ? this.part.attribute(extentNode, "cx") : NaN,
        );
        const cy = Number(
          extentNode ? this.part.attribute(extentNode, "cy") : NaN,
        );
        const picture: InlineRecord = {
          kind: "image",
          elementId: `img:${id}.${pictures}`,
          paragraph: record,
          node: item.child,
          run: item.run,
          ...(Number.isFinite(cx) && Number.isFinite(cy)
            ? { extent: { cx, cy } }
            : {}),
        };
        pictures += 1;
        inlines.push(picture);
        this.register(picture);
      } else if (item.kind === "anchor" || item.kind === "object") {
        const other: InlineRecord = {
          kind: "other",
          elementId: `other:${id}.${others}`,
          paragraph: record,
          node: item.child,
          ...(item.run === item.child ? {} : { run: item.run }),
        };
        others += 1;
        inlines.push(other);
        this.register(other);
      }
    }
    return record;
  }

  table(
    node: XmlElement,
    wrapper: XmlElement | undefined,
  ): TableRecord | undefined {
    const first = this.part
      .findAll("p", node)
      .find((candidate) => candidate.namespace === W_NS);
    const id = first ? this.ids.get(first) : undefined;
    if (!id) return undefined;
    const rows: CellRecord[][] = [];
    const record: TableRecord = {
      kind: "table",
      id,
      elementId: `tbl:${id}`,
      node,
      rows,
      ...(wrapper ? { wrapper } : {}),
    };
    this.register(record);
    for (const row of this.rowsOf(node))
      rows.push(this.cellsOf(row).map((cell) => this.cell(cell, record)));
    return record;
  }

  cell(node: XmlElement, table: TableRecord): CellRecord {
    const paragraphs: ParagraphRecord[] = [];
    for (const child of node.children)
      this.block(child, table, undefined, paragraphs);
    return { node, paragraphs };
  }

  /** `w:tr` children of a table, through row-level `w:sdt` wrappers. */
  rowsOf(table: XmlElement): XmlElement[] {
    const out: XmlElement[] = [];
    const walk = (node: XmlElement): void => {
      for (const child of node.children) {
        if (child.namespace !== W_NS) continue;
        if (child.local === "tr") out.push(child);
        else if (
          child.local === "sdt" ||
          child.local === "sdtContent" ||
          child.local === "customXml"
        )
          walk(child);
      }
    };
    walk(table);
    return out;
  }

  cellsOf(row: XmlElement): XmlElement[] {
    const out: XmlElement[] = [];
    const walk = (node: XmlElement): void => {
      for (const child of node.children) {
        if (child.namespace !== W_NS) continue;
        if (child.local === "tc") out.push(child);
        else if (
          child.local === "sdt" ||
          child.local === "sdtContent" ||
          child.local === "customXml"
        )
          walk(child);
      }
    };
    walk(row);
    return out;
  }

  private register(record: AnyRecord): void {
    this.records.push(record);
    this.byId.set(record.elementId, record);
  }
}

/** The text of a table: cells joined by tabs, rows by newlines. */
export function tableText(table: TableRecord): string {
  return table.rows
    .map((row) => row.map((cell) => cellText(cell)).join("\t"))
    .join("\n");
}

export function cellText(cell: CellRecord): string {
  return cell.paragraphs.map((paragraph) => paragraph.text.text).join("\n");
}

/** The paragraph a cell paragraph's table starts with, for `tbl:` ids. */
export function tableRows(table: TableRecord): readonly (readonly string[])[] {
  return table.rows.map((row) => row.map((cell) => cellText(cell)));
}
