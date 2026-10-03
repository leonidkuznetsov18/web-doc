import type { OoxmlPackage } from "../ooxml/package.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { PageRect } from "../types.js";
import {
  groupMatrix,
  IDENTITY,
  multiply,
  placeFrame,
  readXfrm,
  type Matrix,
  type PlacedFrame,
  type Xfrm,
} from "./geometry.js";
import type { DeckModel, SlideRecord } from "./model.js";
import { resolveTextStyle, readShapeStyle, type StyleSource } from "./style.js";
import {
  firstTextRun,
  readTableCells,
  readTextModel,
  type ParagraphInfo,
  type TextModel,
} from "./text.js";
import { runsCovering, sharedStyle, type TextSpan } from "../range-style.js";
import type {
  PptxElement,
  PptxElementKind,
  PptxPlaceholder,
  PptxTextStyle,
} from "./types.js";

/*
 * The elements of a slide, read from its XML: every child of the shape tree
 * with its frame (own, or inherited from the layout and master for a
 * placeholder, mapped through enclosing groups), its text and its styles.
 */

const TABLE_URI = "http://schemas.openxmlformats.org/drawingml/2006/table";

/** What the engine knows about one element beyond what it reports. */
export interface ShapeRecord {
  readonly element: PptxElement;
  readonly slide: SlideRecord;
  readonly part: XmlPart;
  /** The element node in the slide part (`p:sp`, `p:pic`, …). */
  readonly node: XmlElement;
  readonly cNvPr: XmlElement;
  readonly cNvPrId: number;
  /** The element's own `a:xfrm` or `p:xfrm` node, when it has one. */
  readonly xfrmNode?: XmlElement;
  /** The frame in the element's own space (EMU), inherited when `inherited` is true. */
  readonly xfrm?: Xfrm;
  readonly inherited: boolean;
  /** Transform from the element's space to slide space. */
  readonly parents: Matrix;
  readonly placed?: PlacedFrame;
  readonly txBody?: XmlElement;
  readonly spPr?: XmlElement;
  readonly text?: TextModel;
  /** Inside the fallback branch of `mc:AlternateContent`. */
  readonly readOnly: boolean;
  readonly parent?: ShapeRecord;
  /** The style a span of the shape's text shows: what every run it covers shares. */
  readonly styleOf?: (span: TextSpan) => Partial<PptxTextStyle> | undefined;
}

export interface SlideElements {
  readonly slide: SlideRecord;
  readonly pageIndex: number;
  readonly part: XmlPart;
  /** Document order, groups before their children. */
  readonly records: readonly ShapeRecord[];
  byId(id: string): ShapeRecord | undefined;
}

/** A placeholder of a layout or master, for inheritance. */
export interface PlaceholderRecord {
  readonly part: XmlPart;
  readonly node: XmlElement;
  readonly type: string;
  readonly idx?: number;
  readonly xfrm?: Xfrm;
  readonly lstStyle?: XmlElement;
}

export interface PlaceholderTable {
  readonly part: XmlPart;
  readonly placeholders: readonly PlaceholderRecord[];
  /** `p:txStyles` of a master. */
  readonly txStyles?: XmlElement;
}

/** The placeholders of a layout or master part. */
export async function readPlaceholders(
  pkg: OoxmlPackage,
  partName: string,
  signal?: AbortSignal,
): Promise<PlaceholderTable> {
  const part = await pkg.xml(partName, signal);
  const placeholders: PlaceholderRecord[] = [];
  const spTree = part.find("spTree");
  if (spTree)
    for (const node of spTree.children) {
      if (node.local !== "sp") continue;
      const ph = placeholderOf(part, node);
      if (!ph) continue;
      const spPr = node.children.find((child) => child.local === "spPr");
      const xfrm = readXfrm(
        part,
        spPr?.children.find((child) => child.local === "xfrm"),
      );
      const txBody = node.children.find((child) => child.local === "txBody");
      const lstStyle = txBody?.children.find(
        (child) => child.local === "lstStyle",
      );
      placeholders.push({
        part,
        node,
        type: ph.type,
        ...(ph.idx !== undefined ? { idx: ph.idx } : {}),
        ...(xfrm ? { xfrm } : {}),
        ...(lstStyle ? { lstStyle } : {}),
      });
    }
  const txStyles = part.find("txStyles");
  return { part, placeholders, ...(txStyles ? { txStyles } : {}) };
}

export function placeholderOf(
  part: XmlPart,
  node: XmlElement,
): PptxPlaceholder | undefined {
  const nv = node.children.find(
    (child) => child.local.startsWith("nv") && child.local.endsWith("Pr"),
  );
  const nvPr = nv?.children.find((child) => child.local === "nvPr");
  const ph = nvPr?.children.find((child) => child.local === "ph");
  if (!ph) return undefined;
  const idxValue = part.attribute(ph, "idx");
  const idx = idxValue === undefined ? undefined : Number(idxValue);
  return {
    type: part.attribute(ph, "type") ?? "body",
    ...(idx !== undefined && Number.isInteger(idx) ? { idx } : {}),
  };
}

/** Placeholder types that share position and style: titles, and body-like content. */
export function placeholderFamily(type: string): string {
  switch (type) {
    case "title":
    case "ctrTitle":
      return "title";
    case "body":
    case "subTitle":
    case "obj":
      return "body";
    default:
      return type;
  }
}

/** The placeholder a slide placeholder inherits from: same idx first, then the same family. */
export function matchPlaceholder(
  table: PlaceholderTable | undefined,
  ph: PptxPlaceholder,
): PlaceholderRecord | undefined {
  if (!table) return undefined;
  const family = placeholderFamily(ph.type);
  if (ph.idx !== undefined) {
    const byIdx = table.placeholders.filter((item) => item.idx === ph.idx);
    const same = byIdx.find((item) => placeholderFamily(item.type) === family);
    if (same) return same;
    if (byIdx[0] && family === "body") return byIdx[0];
  }
  return table.placeholders.find(
    (item) => placeholderFamily(item.type) === family,
  );
}

export interface InheritanceSources {
  layout?: PlaceholderTable;
  master?: PlaceholderTable;
  /** `p:defaultTextStyle` of the presentation. */
  defaults?: StyleSource;
}

interface Walk {
  readonly records: ShapeRecord[];
  readonly counts: Map<number, number>;
}

/** Reads every element of a slide. */
export async function readSlideElements(
  pkg: OoxmlPackage,
  model: DeckModel,
  slide: SlideRecord,
  pageIndex: number,
  sources: InheritanceSources,
  signal?: AbortSignal,
): Promise<SlideElements> {
  const part = await pkg.xml(slide.part, signal);
  const fonts = await model.themeFonts(slide.layout?.master);
  const walk: Walk = { records: [], counts: new Map() };
  const spTree = part.find("spTree");
  const context: ElementContext = { slide, pageIndex, part, sources, fonts };
  if (spTree) walkTree(context, spTree, IDENTITY, undefined, false, walk);
  const byId = new Map(
    walk.records.map((record) => [record.element.id, record]),
  );
  return {
    slide,
    pageIndex,
    part,
    records: walk.records,
    byId: (id) => byId.get(id),
  };
}

interface ElementContext {
  readonly slide: SlideRecord;
  readonly pageIndex: number;
  readonly part: XmlPart;
  readonly sources: InheritanceSources;
  readonly fonts: { readonly major: string; readonly minor: string };
}

function walkTree(
  context: ElementContext,
  container: XmlElement,
  parents: Matrix,
  parent: ShapeRecord | undefined,
  readOnly: boolean,
  walk: Walk,
): void {
  for (const node of container.children) {
    switch (node.local) {
      case "sp":
      case "pic":
      case "graphicFrame":
      case "cxnSp":
      case "grpSp": {
        const record = readRecord(
          context,
          node,
          parents,
          parent,
          readOnly,
          walk,
        );
        if (!record) break;
        walk.records.push(record);
        if (node.local === "grpSp")
          walkTree(
            context,
            node,
            record.xfrm ? multiply(parents, groupMatrix(record.xfrm)) : parents,
            record,
            readOnly,
            walk,
          );
        break;
      }
      case "AlternateContent": {
        const fallback = node.children.find(
          (child) => child.local === "Fallback",
        );
        if (fallback) walkTree(context, fallback, parents, parent, true, walk);
        break;
      }
      default:
        break;
    }
  }
}

function readRecord(
  context: ElementContext,
  node: XmlElement,
  parents: Matrix,
  parent: ShapeRecord | undefined,
  readOnly: boolean,
  walk: Walk,
): ShapeRecord | undefined {
  const { part, slide } = context;
  const nv = node.children.find(
    (child) => child.local.startsWith("nv") && child.local.endsWith("Pr"),
  );
  const cNvPr = nv?.children.find((child) => child.local === "cNvPr");
  if (!cNvPr) return undefined;
  const cNvPrId = Number(part.attribute(cNvPr, "id") ?? NaN);
  if (!Number.isInteger(cNvPrId)) return undefined;
  const seen = (walk.counts.get(cNvPrId) ?? 0) + 1;
  walk.counts.set(cNvPrId, seen);
  const id = `${slide.key}:${cNvPrId}${seen > 1 ? `#${seen}` : ""}`;
  const name = part.attribute(cNvPr, "name") ?? "";
  const hiddenAttribute = part.attribute(cNvPr, "hidden");
  const hidden = hiddenAttribute === "1" || hiddenAttribute === "true";
  const placeholder = placeholderOf(part, node);

  let kind: PptxElementKind;
  let xfrmHolder: XmlElement | undefined;
  switch (node.local) {
    case "sp":
      kind = "shape";
      xfrmHolder = node.children.find((child) => child.local === "spPr");
      break;
    case "pic":
      kind = "image";
      xfrmHolder = node.children.find((child) => child.local === "spPr");
      break;
    case "cxnSp":
      kind = "connector";
      xfrmHolder = node.children.find((child) => child.local === "spPr");
      break;
    case "grpSp":
      kind = "group";
      xfrmHolder = node.children.find((child) => child.local === "grpSpPr");
      break;
    default: {
      const data = node.children
        .find((child) => child.local === "graphic")
        ?.children.find((child) => child.local === "graphicData");
      kind =
        data && part.attribute(data, "uri") === TABLE_URI ? "table" : "other";
      xfrmHolder = node;
      break;
    }
  }
  const xfrmNode =
    node.local === "graphicFrame"
      ? node.children.find((child) => child.local === "xfrm")
      : xfrmHolder?.children.find((child) => child.local === "xfrm");
  let xfrm = readXfrm(part, xfrmNode);
  let inherited = false;
  if (!xfrm && placeholder && !parent) {
    xfrm =
      matchPlaceholder(context.sources.layout, placeholder)?.xfrm ??
      matchPlaceholder(context.sources.master, placeholder)?.xfrm;
    inherited = xfrm !== undefined;
  }
  const placed = xfrm ? placeFrame(xfrm, parents) : undefined;
  const spPr = node.children.find((child) => child.local === "spPr");
  const txBody =
    kind === "shape"
      ? node.children.find((child) => child.local === "txBody")
      : undefined;
  const text = txBody ? readTextModel(part, txBody) : undefined;

  const zero: PageRect = { x: 0, y: 0, width: 0, height: 0 };
  const bounds = placed?.bounds ?? zero;
  const operations = readOnly ? [] : operationsFor(kind, txBody !== undefined);
  let elementText: string | undefined = text?.text;
  let table: { readonly rows: readonly (readonly string[])[] } | undefined;
  if (kind === "table") {
    const tbl = part.find("tbl", node);
    const rows = tbl
      ? readTableCells(part, tbl).map((row) =>
          row.map((cell) => cell.model.text),
        )
      : [];
    table = { rows };
    elementText = rows.map((row) => row.join("\t")).join("\n");
  }
  const element: PptxElement = {
    id,
    kind,
    pageIndex: context.pageIndex,
    bounds,
    ...(placed && placed.frame.rotation !== 0
      ? { rotation: placed.frame.rotation }
      : {}),
    ...(placed ? { frame: placed.frame } : {}),
    ...(elementText !== undefined ? { text: elementText } : {}),
    ...(parent ? { parentId: parent.element.id } : {}),
    operations,
    name,
    ...(placeholder ? { placeholder } : {}),
    ...(text && txBody
      ? { textStyle: textStyleOf(context, placeholder, txBody, text) }
      : {}),
    ...(kind === "shape" || kind === "connector"
      ? { shapeStyle: readShapeStyle(part, spPr) }
      : {}),
    ...(table ? { table } : {}),
    ...(hidden ? { hidden: true } : {}),
  };
  return {
    element,
    slide,
    part,
    node,
    cNvPr,
    cNvPrId,
    ...(xfrmNode ? { xfrmNode } : {}),
    ...(xfrm ? { xfrm } : {}),
    inherited,
    parents,
    ...(placed ? { placed } : {}),
    ...(txBody ? { txBody } : {}),
    ...(spPr ? { spPr } : {}),
    ...(text ? { text } : {}),
    readOnly,
    ...(parent ? { parent } : {}),
    ...(text && txBody
      ? {
          styleOf: (span: TextSpan) =>
            rangeStyleOf(context, placeholder, txBody, text, span),
        }
      : {}),
  };
}

function operationsFor(kind: PptxElementKind, hasText: boolean): string[] {
  const geometry = ["moveElement", "resizeElement", "deleteElement"];
  switch (kind) {
    case "shape":
      return [
        ...(hasText ? ["replaceText", "setTextStyle"] : []),
        "setShapeStyle",
        ...geometry,
      ];
    case "connector":
      return ["setShapeStyle", ...geometry];
    case "table":
      return ["setTableCell", ...geometry];
    default:
      return geometry;
  }
}

/** The list-style chain of a shape, in priority order. */
export function listStylesOf(
  context: {
    readonly part: XmlPart;
    readonly sources: InheritanceSources;
  },
  txBody: XmlElement,
  placeholder: PptxPlaceholder | undefined,
): StyleSource[] {
  const lists: StyleSource[] = [];
  const own = txBody.children.find((child) => child.local === "lstStyle");
  if (own) lists.push({ part: context.part, node: own });
  const { layout, master, defaults } = context.sources;
  if (placeholder) {
    const fromLayout = matchPlaceholder(layout, placeholder);
    if (fromLayout?.lstStyle)
      lists.push({ part: fromLayout.part, node: fromLayout.lstStyle });
    const fromMaster = matchPlaceholder(master, placeholder);
    if (fromMaster?.lstStyle)
      lists.push({ part: fromMaster.part, node: fromMaster.lstStyle });
  }
  if (master?.txStyles) {
    const family = placeholder ? placeholderFamily(placeholder.type) : "other";
    const styleName =
      family === "title"
        ? "titleStyle"
        : family === "body"
          ? "bodyStyle"
          : "otherStyle";
    const style = master.txStyles.children.find(
      (child) => child.local === styleName,
    );
    if (style) lists.push({ part: master.part, node: style });
  }
  if (defaults) lists.push(defaults);
  return lists;
}

/**
 * The style a span of a text body shows: each run it covers resolves through
 * its paragraph and the inherited lists, and the properties they all share
 * are kept. An empty paragraph reads its end-of-paragraph properties.
 */
function rangeStyleOf(
  context: ElementContext,
  placeholder: PptxPlaceholder | undefined,
  txBody: XmlElement,
  text: TextModel,
  span: TextSpan,
): Partial<PptxTextStyle> | undefined {
  const lists = listStylesOf(context, txBody, placeholder);
  const styleOf = (paragraph: ParagraphInfo, rPr: XmlElement | undefined) =>
    resolveTextStyle({
      ...(rPr ? { rPr: { part: context.part, node: rPr } } : {}),
      ...(paragraph.pPr
        ? { pPr: { part: context.part, node: paragraph.pPr } }
        : {}),
      level: paragraph.level,
      lists,
      fonts: context.fonts,
    });
  const runs = text.paragraphs.flatMap((paragraph) =>
    paragraph.runs.map((run) => ({
      start: run.start,
      end: run.end,
      style: () => styleOf(paragraph, run.rPr),
    })),
  );
  const covered = runsCovering(runs, span);
  if (covered.length > 0) return sharedStyle(covered.map((run) => run.style()));
  const paragraph =
    text.paragraphs.find(
      (candidate) =>
        candidate.start <= span.start && span.start <= candidate.end,
    ) ?? text.paragraphs[0];
  return paragraph
    ? sharedStyle([styleOf(paragraph, paragraph.endParaRPr)])
    : undefined;
}

function textStyleOf(
  context: ElementContext,
  placeholder: PptxPlaceholder | undefined,
  txBody: XmlElement,
  text: TextModel,
) {
  const first = firstTextRun(text);
  const paragraph = first?.paragraph ?? text.paragraphs[0];
  const rPr = first?.run.rPr ?? paragraph?.endParaRPr;
  return resolveTextStyle({
    ...(rPr ? { rPr: { part: context.part, node: rPr } } : {}),
    ...(paragraph?.pPr
      ? { pPr: { part: context.part, node: paragraph.pPr } }
      : {}),
    level: paragraph?.level ?? 0,
    lists: listStylesOf(context, txBody, placeholder),
    fonts: context.fonts,
  });
}
