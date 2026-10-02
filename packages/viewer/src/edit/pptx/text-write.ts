import type { EditColor } from "../types.js";
import { escapeText, type XmlElement, type XmlPart } from "../ooxml/xml.js";
import type { ParagraphInfo, RunInfo, TextModel } from "./text.js";
import { LINE_BREAK, PARAGRAPH_BREAK } from "./text.js";
import type { PptxTextAlign, PptxTextStyleChange } from "./types.js";

/*
 * Writing DrawingML text: paragraphs rebuilt from run items that are either
 * the original bytes of an untouched run or a run re-serialized with its
 * properties, run properties changed in schema order, colours in their
 * theme form, and the autofit scale dropped when a body's text changes.
 */

/** A run of a rebuilt paragraph: untouched bytes, or a run or break with properties. */
export type RunItem =
  | { readonly kind: "raw"; readonly xml: string }
  | { readonly kind: "run"; readonly rPr: string; readonly text: string }
  | { readonly kind: "br"; readonly rPr: string };

export interface ParagraphDraft {
  /** The `a:pPr` bytes, or "". */
  pPr: string;
  items: RunItem[];
  /** The `a:endParaRPr` bytes, or "". */
  endParaRPr: string;
}

const FILLS = new Set([
  "noFill",
  "solidFill",
  "gradFill",
  "blipFill",
  "pattFill",
  "grpFill",
]);
/** Children of a:rPr that follow the fill in the schema. */
const AFTER_FILL = new Set([
  "effectLst",
  "effectDag",
  "highlight",
  "uLnTx",
  "uLn",
  "uFillTx",
  "uFill",
  "latin",
  "ea",
  "cs",
  "sym",
  "hlinkClick",
  "hlinkMouseOver",
  "rtl",
  "extLst",
]);
/** Children of a:rPr that follow a:latin in the schema. */
const AFTER_LATIN = new Set([
  "ea",
  "cs",
  "sym",
  "hlinkClick",
  "hlinkMouseOver",
  "rtl",
  "extLst",
]);

export const THEME_COLORS = new Set([
  "bg1",
  "tx1",
  "bg2",
  "tx2",
  "dk1",
  "lt1",
  "dk2",
  "lt2",
  "accent1",
  "accent2",
  "accent3",
  "accent4",
  "accent5",
  "accent6",
  "hlink",
  "folHlink",
]);

export const COLOR_MODIFIERS = new Set([
  "tint",
  "shade",
  "alpha",
  "lumMod",
  "lumOff",
  "satMod",
  "satOff",
  "hueMod",
  "hueOff",
]);

/** Characters XML 1.0 cannot carry; a lone surrogate is caught separately. */
const FORBIDDEN = /[\u0000-\u0008\u000c\u000e-\u001f\ufffe\uffff]/;

const LONE_SURROGATE =
  /[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/;

/** Why a text cannot be written, or undefined when it can. */
export function textProblem(text: string): string | undefined {
  if (FORBIDDEN.test(text)) return "control characters XML cannot carry";
  if (LONE_SURROGATE.test(text)) return "a lone surrogate";
  return undefined;
}

export function colorProblem(color: unknown): string | undefined {
  if (typeof color === "string")
    return /^#[0-9A-Fa-f]{6}$/.test(color) ? undefined : "not #RRGGBB";
  if (!color || typeof color !== "object") return "not a colour";
  const { theme, mods } = color as {
    readonly theme?: unknown;
    readonly mods?: unknown;
  };
  if (typeof theme !== "string" || !THEME_COLORS.has(theme))
    return "not a theme colour slot";
  if (mods !== undefined) {
    if (!mods || typeof mods !== "object" || Array.isArray(mods))
      return "mods is not an object";
    for (const [name, value] of Object.entries(mods)) {
      if (!COLOR_MODIFIERS.has(name)) return `unknown modifier ${name}`;
      if (
        typeof value !== "number" ||
        !Number.isInteger(value) ||
        Math.abs(value) > 100_000
      )
        return `modifier ${name} is not an integer within ±100000`;
    }
  }
  return undefined;
}

/** `<a:solidFill>` for a colour in its theme or RGB form. */
export function solidFillXml(color: EditColor): string {
  if (typeof color === "string")
    return `<a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill>`;
  const mods = Object.entries(color.mods ?? {})
    .map(([name, value]) => `<a:${name} val="${value}"/>`)
    .join("");
  return mods
    ? `<a:solidFill><a:schemeClr val="${color.theme}">${mods}</a:schemeClr></a:solidFill>`
    : `<a:solidFill><a:schemeClr val="${color.theme}"/></a:solidFill>`;
}

export function alignValue(align: PptxTextAlign): string {
  switch (align) {
    case "left":
      return "l";
    case "center":
      return "ctr";
    case "right":
      return "r";
    default:
      return "just";
  }
}

/** The bytes of a node, or "" without one. */
export function sliceOf(part: XmlPart, node: XmlElement | undefined): string {
  return node ? part.text.slice(node.start, node.end) : "";
}

/**
 * Run properties with a change applied: attributes kept and set, children
 * kept in order with the fill and the Latin face replaced where the schema
 * puts them. `source` may be an `a:rPr`, an `a:endParaRPr` (renamed) or
 * nothing.
 */
export function changedRunProperties(
  part: XmlPart,
  source: XmlElement | undefined,
  change: PptxTextStyleChange,
  tag = "a:rPr",
): string {
  const attributes = new Map<string, string>();
  const children: { local: string; xml: string }[] = [];
  if (source) {
    for (const attribute of source.attributes)
      attributes.set(attribute.name, attribute.rawValue);
    for (const child of source.children)
      children.push({ local: child.local, xml: sliceOf(part, child) });
  }
  if (change.bold !== undefined) attributes.set("b", change.bold ? "1" : "0");
  if (change.italic !== undefined)
    attributes.set("i", change.italic ? "1" : "0");
  if (change.underline !== undefined)
    attributes.set("u", change.underline ? "sng" : "none");
  if (change.fontSize !== undefined)
    attributes.set("sz", String(Math.round(change.fontSize * 100)));
  if (change.color !== undefined)
    replaceChild(
      children,
      FILLS,
      AFTER_FILL,
      "solidFill",
      solidFillXml(change.color),
    );
  if (change.fontFamily !== undefined)
    replaceChild(
      children,
      new Set(["latin"]),
      AFTER_LATIN,
      "latin",
      `<a:latin typeface="${escapeAttributeValue(change.fontFamily)}"/>`,
    );
  const attributeText = [...attributes]
    .map(([name, raw]) => ` ${name}="${raw}"`)
    .join("");
  const content = children.map((child) => child.xml).join("");
  return content
    ? `<${tag}${attributeText}>${content}</${tag}>`
    : `<${tag}${attributeText}/>`;
}

function replaceChild(
  children: { local: string; xml: string }[],
  replaced: ReadonlySet<string>,
  after: ReadonlySet<string>,
  local: string,
  xml: string,
): void {
  const first = children.findIndex((child) => replaced.has(child.local));
  for (let index = children.length - 1; index >= 0; index -= 1)
    if (replaced.has(children[index]!.local)) children.splice(index, 1);
  let at =
    first >= 0 ? first : children.findIndex((child) => after.has(child.local));
  if (at < 0) at = children.length;
  children.splice(at, 0, { local, xml });
}

function escapeAttributeValue(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll('"', "&quot;")
    .replaceAll("\t", "&#9;")
    .replaceAll("\n", "&#10;")
    .replaceAll("\r", "&#13;");
}

/** Text as the engine stores it: Windows and classic Mac line ends become "\n". */
export function normalizeText(text: string): string {
  return text.replaceAll("\r\n", "\n").replaceAll("\r", "\n");
}

/** `a:pPr` bytes with `algn` set, created when the paragraph has none. */
export function alignedParagraphProperties(
  part: XmlPart,
  pPr: XmlElement | undefined,
  align: PptxTextAlign,
): string {
  const value = alignValue(align);
  if (!pPr) return `<a:pPr algn="${value}"/>`;
  const xml = sliceOf(part, pPr);
  const existing = pPr.attributes.find(
    (attribute) => attribute.name === "algn",
  );
  if (existing)
    return (
      xml.slice(0, existing.start - pPr.start) +
      `algn="${value}"` +
      xml.slice(existing.end - pPr.start)
    );
  const insertAt = pPr.start + 1 + pPr.name.length - pPr.start;
  return `${xml.slice(0, insertAt)} algn="${value}"${xml.slice(insertAt)}`;
}

/** A run or break as the file will carry it. */
export function runXml(item: RunItem): string {
  switch (item.kind) {
    case "raw":
      return item.xml;
    case "br":
      return item.rPr ? `<a:br>${item.rPr}</a:br>` : "<a:br/>";
    default:
      return item.text.length === 0
        ? ""
        : `<a:r>${item.rPr}<a:t>${escapeText(item.text)}</a:t></a:r>`;
  }
}

export function paragraphXml(draft: ParagraphDraft): string {
  return `<a:p>${draft.pPr}${draft.items.map(runXml).join("")}${draft.endParaRPr}</a:p>`;
}

/** Items of text with its paragraph breaks removed: runs and breaks in `rPr`. */
export function itemsOfSegment(segment: string, rPr: string): RunItem[] {
  const items: RunItem[] = [];
  const lines = segment.split(LINE_BREAK);
  lines.forEach((line, index) => {
    if (index > 0) items.push({ kind: "br", rPr });
    if (line.length > 0) items.push({ kind: "run", rPr, text: line });
  });
  return items;
}

/** The rPr a run of inserted text takes at `offset`: the run before it, else the run after, else the paragraph end. */
export function styleAt(
  part: XmlPart,
  paragraph: ParagraphInfo,
  offset: number,
): string {
  let before: RunInfo | undefined;
  for (const run of paragraph.runs)
    if (run.kind !== "br" && run.start < offset && offset <= run.end)
      before = run;
  const after = paragraph.runs.find(
    (run) => run.kind !== "br" && run.start <= offset && offset < run.end,
  );
  const inside = paragraph.runs.find(
    (run) => run.kind !== "br" && run.start < offset && offset < run.end,
  );
  const source = inside ?? before ?? after;
  if (source) return runProperties(part, source);
  return paragraph.endParaRPr
    ? renamedProperties(part, paragraph.endParaRPr, "a:rPr")
    : "";
}

/** The `a:rPr` bytes of a run, or "". */
export function runProperties(part: XmlPart, run: RunInfo): string {
  return sliceOf(part, run.rPr);
}

/** The bytes of a properties element under another tag name. */
export function renamedProperties(
  part: XmlPart,
  node: XmlElement,
  tag: string,
): string {
  const xml = sliceOf(part, node);
  if (node.selfClosing)
    return `<${tag}${xml.slice(1 + node.name.length, -2)}/>`;
  return `<${tag}${xml.slice(1 + node.name.length, node.contentStart - node.start)}${xml.slice(
    node.contentStart - node.start,
    node.contentEnd - node.start,
  )}</${tag}>`;
}

/**
 * The bytes of a text body's content before its first paragraph, with the
 * stale autofit scale dropped: `a:normAutofit`'s `fontScale` and
 * `lnSpcReduction` describe the old text.
 */
export function bodyPrefix(
  part: XmlPart,
  txBody: XmlElement,
  firstParagraph: XmlElement | undefined,
): string {
  const end = firstParagraph ? firstParagraph.start : txBody.contentEnd;
  let prefix = part.text.slice(txBody.contentStart, end);
  const bodyPr = txBody.children.find((child) => child.local === "bodyPr");
  const autofit = bodyPr?.children.find(
    (child) => child.local === "normAutofit",
  );
  if (!autofit) return prefix;
  const cuts = autofit.attributes
    .filter((attribute) =>
      ["fontScale", "lnSpcReduction"].includes(attribute.name),
    )
    .map((attribute) => {
      let start = attribute.start;
      while (start > autofit.start && /\s/.test(part.text[start - 1]!))
        start -= 1;
      return {
        start: start - txBody.contentStart,
        end: attribute.end - txBody.contentStart,
      };
    })
    .sort((a, b) => b.start - a.start);
  for (const cut of cuts)
    prefix = prefix.slice(0, cut.start) + prefix.slice(cut.end);
  return prefix;
}

/** The bytes after the last paragraph (an extension list, usually nothing). */
export function bodySuffix(
  part: XmlPart,
  txBody: XmlElement,
  lastParagraph: XmlElement | undefined,
): string {
  return lastParagraph
    ? part.text.slice(lastParagraph.end, txBody.contentEnd)
    : "";
}

/** The index of the paragraph that holds an offset, taking the one before a break first. */
export function paragraphAt(model: TextModel, offset: number): number {
  const index = model.paragraphs.findIndex(
    (paragraph) => paragraph.start <= offset && offset <= paragraph.end,
  );
  return index < 0 ? Math.max(0, model.paragraphs.length - 1) : index;
}

export { PARAGRAPH_BREAK };
