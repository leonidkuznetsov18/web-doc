import { patches, type XmlPatch } from "../ooxml/patch.js";
import { escapeText, type XmlElement, type XmlPart } from "../ooxml/xml.js";
import type { EditColor } from "../types.js";
import { W14_NS, W_NS } from "./ids.js";
import { isThemeColorName, type DocxStyles } from "./style.js";
import { LINE_BREAK, PAGE_BREAK, TAB } from "./text.js";
import type { DocxParagraphStyleChange, DocxTextStyleChange } from "./types.js";

/*
 * Writing WordprocessingML: runs re-serialized from text and their
 * properties, run and paragraph properties changed in schema order with
 * every other child kept as its bytes, paragraph start tags carrying their
 * `w14:paraId`, and the namespace declarations that attribute needs.
 */

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Children of `w:rPr` in schema order (CT_RPr). */
const RPR_ORDER = [
  "rStyle",
  "rFonts",
  "b",
  "bCs",
  "i",
  "iCs",
  "caps",
  "smallCaps",
  "strike",
  "dstrike",
  "outline",
  "shadow",
  "emboss",
  "imprint",
  "noProof",
  "snapToGrid",
  "vanish",
  "webHidden",
  "color",
  "spacing",
  "w",
  "kern",
  "position",
  "sz",
  "szCs",
  "highlight",
  "u",
  "effect",
  "bdr",
  "shd",
  "fitText",
  "vertAlign",
  "rtl",
  "cs",
  "em",
  "lang",
  "eastAsianLayout",
  "specVanish",
  "oMath",
  "rPrChange",
];

/** Children of `w:pPr` in schema order (CT_PPr). */
const PPR_ORDER = [
  "pStyle",
  "keepNext",
  "keepLines",
  "pageBreakBefore",
  "framePr",
  "widowControl",
  "numPr",
  "suppressLineNumbers",
  "pBdr",
  "shd",
  "tabs",
  "suppressAutoHyphens",
  "kinsoku",
  "wordWrap",
  "overflowPunct",
  "topLinePunct",
  "autoSpaceDE",
  "autoSpaceDN",
  "bidi",
  "adjustRightInd",
  "snapToGrid",
  "spacing",
  "ind",
  "contextualSpacing",
  "mirrorIndents",
  "suppressOverlap",
  "jc",
  "textDirection",
  "textAlignment",
  "textboxTightWrap",
  "outlineLvl",
  "divId",
  "cnfStyle",
  "rPr",
  "sectPr",
  "pPrChange",
];

const FORBIDDEN = /[\u0000-\u0008\u000e-\u001f￾￿]/;
const LONE_SURROGATE =
  /[\ud800-\udbff](?![\udc00-\udfff])|(?<![\ud800-\udbff])[\udc00-\udfff]/;

/** Why a text cannot be written, or undefined when it can. */
export function textProblem(text: string): string | undefined {
  if (FORBIDDEN.test(text)) return "control characters XML cannot carry";
  if (LONE_SURROGATE.test(text)) return "a lone surrogate";
  return undefined;
}

/** Why a value cannot be written as an attribute: text problems plus the breaks text may hold. */
export function attributeProblem(value: string): string | undefined {
  if (/[\u000b\u000c]/.test(value))
    return "line or page breaks an attribute cannot carry";
  return textProblem(value);
}

export function colorProblem(color: unknown): string | undefined {
  if (typeof color === "string")
    return color === "auto" || /^#[0-9A-Fa-f]{6}$/.test(color)
      ? undefined
      : "not #RRGGBB or auto";
  if (!color || typeof color !== "object") return "not a colour";
  const { theme } = color as { readonly theme?: unknown };
  return typeof theme === "string" && isThemeColorName(theme)
    ? undefined
    : "not a theme colour name";
}

/** Text as the engine stores it: Windows and classic Mac line ends become "\n". */
export function normalizeText(text: string): string {
  return text.replaceAll("\r\n", "\n").replaceAll("\r", "\n");
}

export function sliceOf(part: XmlPart, node: XmlElement | undefined): string {
  return node ? part.text.slice(node.start, node.end) : "";
}

export function escapeAttributeValue(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll('"', "&quot;")
    .replaceAll("\t", "&#9;")
    .replaceAll("\n", "&#10;")
    .replaceAll("\r", "&#13;");
}

/** The content elements of a run for `text`: `w:t`, `w:tab`, `w:br`. */
export function runContentXml(text: string): string {
  let out = "";
  let pending = "";
  const flush = (): void => {
    if (pending.length === 0) return;
    const preserve = /^\s|\s$/.test(pending) ? ' xml:space="preserve"' : "";
    out += `<w:t${preserve}>${escapeText(pending)}</w:t>`;
    pending = "";
  };
  for (const char of text) {
    if (char === TAB) {
      flush();
      out += "<w:tab/>";
    } else if (char === LINE_BREAK) {
      flush();
      out += "<w:br/>";
    } else if (char === PAGE_BREAK) {
      flush();
      out += '<w:br w:type="page"/>';
    } else pending += char;
  }
  flush();
  return out;
}

/** A run with `rPr` bytes (or "") and the given content; nothing for empty content. */
export function runXml(rPr: string, content: string): string {
  return content.length === 0 ? "" : `<w:r>${rPr}${content}</w:r>`;
}

/**
 * Properties with changes merged in schema order: every child of `source`
 * keeps its bytes unless `set` names its local name, in which case the
 * replacement (or nothing, for `null`) takes its place; new children go
 * where the schema puts them. Returns "" for no properties at all.
 */
export function mergedProperties(
  part: XmlPart,
  source: XmlElement | undefined,
  tag: "w:rPr" | "w:pPr",
  order: readonly string[],
  set: ReadonlyMap<string, string | null>,
): string {
  const children: { local: string; xml: string }[] = source
    ? source.children.map((child) => ({
        local: child.local,
        xml: sliceOf(part, child),
      }))
    : [];
  for (const [local, xml] of set) {
    const index = children.findIndex((child) => child.local === local);
    for (let at = children.length - 1; at >= 0; at -= 1)
      if (children[at]!.local === local) children.splice(at, 1);
    if (xml === null) continue;
    let at = index;
    if (at < 0) {
      // Before the first child the schema places later; a child the
      // schema list does not name (an extension) counts as later too.
      const rank = order.indexOf(local);
      at = children.findIndex((child) => {
        const other = order.indexOf(child.local);
        return other < 0 || other > rank;
      });
      if (at < 0) at = children.length;
    }
    children.splice(at, 0, { local, xml });
  }
  const attributes = source
    ? source.attributes.map((a) => ` ${a.name}="${a.rawValue}"`).join("")
    : "";
  const content = children.map((child) => child.xml).join("");
  if (!content) return source ? `<${tag}${attributes}/>` : "";
  return `<${tag}${attributes}>${content}</${tag}>`;
}

function toggleXml(local: string, on: boolean): string {
  return on ? `<w:${local}/>` : `<w:${local} w:val="0"/>`;
}

/** `w:rPr` bytes with a text style change applied. */
export function changedRunProperties(
  part: XmlPart,
  rPr: XmlElement | undefined,
  change: DocxTextStyleChange,
  styles: DocxStyles,
): string {
  const set = new Map<string, string | null>();
  if (change.bold !== undefined) {
    set.set("b", toggleXml("b", change.bold));
    set.set("bCs", toggleXml("bCs", change.bold));
  }
  if (change.italic !== undefined) {
    set.set("i", toggleXml("i", change.italic));
    set.set("iCs", toggleXml("iCs", change.italic));
  }
  if (change.underline !== undefined)
    set.set("u", `<w:u w:val="${change.underline ? "single" : "none"}"/>`);
  if (change.fontSize !== undefined) {
    const half = String(Math.round(change.fontSize * 2));
    set.set("sz", `<w:sz w:val="${half}"/>`);
    set.set("szCs", `<w:szCs w:val="${half}"/>`);
  }
  if (change.color !== undefined)
    set.set("color", colorXml(change.color, styles));
  if (change.highlight !== undefined)
    set.set(
      "highlight",
      change.highlight === "none"
        ? null
        : `<w:highlight w:val="${change.highlight}"/>`,
    );
  if (change.fontFamily !== undefined)
    set.set("rFonts", fontsXml(part, rPr, change.fontFamily));
  return mergedProperties(part, rPr, "w:rPr", RPR_ORDER, set);
}

/** `w:color` for a colour: RGB, automatic, or a theme name with the theme's value. */
export function colorXml(color: EditColor, styles: DocxStyles): string {
  if (typeof color === "string") {
    if (color === "auto") return '<w:color w:val="auto"/>';
    return `<w:color w:val="${color.slice(1).toUpperCase()}"/>`;
  }
  const value = styles.themeColor(color.theme) ?? "auto";
  return `<w:color w:val="${value}" w:themeColor="${escapeAttributeValue(color.theme)}"/>`;
}

/** `w:rFonts` with the ASCII and high-ANSI faces set and their theme overrides dropped. */
function fontsXml(
  part: XmlPart,
  rPr: XmlElement | undefined,
  family: string,
): string {
  const existing = rPr?.children.find(
    (child) => child.local === "rFonts" && child.namespace === W_NS,
  );
  const kept = (existing?.attributes ?? []).filter(
    (attribute) =>
      !["w:ascii", "w:hAnsi", "w:asciiTheme", "w:hAnsiTheme"].includes(
        attribute.name,
      ),
  );
  const face = escapeAttributeValue(family);
  const rest = kept.map((a) => ` ${a.name}="${a.rawValue}"`).join("");
  return `<w:rFonts w:ascii="${face}" w:hAnsi="${face}"${rest}/>`;
}

/** `w:pPr` bytes with alignment and spacing changed, other children kept. */
export function changedParagraphProperties(
  part: XmlPart,
  pPr: XmlElement | undefined,
  change: DocxParagraphStyleChange,
): string {
  const set = new Map<string, string | null>();
  if (change.align !== undefined)
    set.set("jc", `<w:jc w:val="${alignValue(change.align)}"/>`);
  if (change.spacing !== undefined) {
    const existing = pPr?.children.find(
      (child) => child.local === "spacing" && child.namespace === W_NS,
    );
    const attributes = new Map<string, string>();
    for (const attribute of existing?.attributes ?? [])
      attributes.set(attribute.name, attribute.rawValue);
    // Word prefers the line-based and automatic forms over w:before and
    // w:after, so a set value drops them.
    if (change.spacing.before !== undefined) {
      attributes.set(
        "w:before",
        String(Math.round(change.spacing.before * 20)),
      );
      attributes.delete("w:beforeLines");
      attributes.delete("w:beforeAutospacing");
    }
    if (change.spacing.after !== undefined) {
      attributes.set("w:after", String(Math.round(change.spacing.after * 20)));
      attributes.delete("w:afterLines");
      attributes.delete("w:afterAutospacing");
    }
    if (change.spacing.line !== undefined) {
      attributes.set("w:line", String(Math.round(change.spacing.line * 240)));
      attributes.set("w:lineRule", "auto");
    }
    set.set(
      "spacing",
      `<w:spacing${[...attributes].map(([name, raw]) => ` ${name}="${raw}"`).join("")}/>`,
    );
  }
  return mergedProperties(part, pPr, "w:pPr", PPR_ORDER, set);
}

export function alignValue(
  align: NonNullable<DocxParagraphStyleChange["align"]>,
): string {
  switch (align) {
    case "center":
      return "center";
    case "right":
      return "right";
    case "justify":
      return "both";
    default:
      return "left";
  }
}

/** `w:pPr` bytes with the paragraph mark's `w:rPr` changed (created when absent). */
export function paragraphMarkProperties(
  part: XmlPart,
  pPr: XmlElement | undefined,
  change: DocxTextStyleChange,
  styles: DocxStyles,
): string {
  const rPr = pPr?.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
  const changed = changedRunProperties(part, rPr, change, styles);
  return mergedProperties(
    part,
    pPr,
    "w:pPr",
    PPR_ORDER,
    new Map([["rPr", changed || null]]),
  );
}

/** `w:pPr` bytes without the section properties, for a paragraph copied next to one that ends a section. */
export function paragraphPropertiesWithoutSection(
  part: XmlPart,
  pPr: XmlElement | undefined,
): string {
  if (!pPr) return "";
  return mergedProperties(
    part,
    pPr,
    "w:pPr",
    PPR_ORDER,
    new Map([["sectPr", null]]),
  );
}

/**
 * The start tag of a `w:p` with its id written: the paragraph's own
 * attributes, `w14:paraId` and `w14:textId` replaced.
 */
export function paragraphStartTag(
  source: XmlElement | undefined,
  id: string,
): string {
  const kept = (source?.attributes ?? []).filter(
    (attribute) =>
      attribute.name !== "w14:paraId" && attribute.name !== "w14:textId",
  );
  const rest = kept.map((a) => ` ${a.name}="${a.rawValue}"`).join("");
  return `<w:p${rest} w14:paraId="${id}">`;
}

/** A whole paragraph from its start tag, property bytes and content. */
export function paragraphXml(
  source: XmlElement | undefined,
  id: string,
  pPr: string,
  content: string,
): string {
  return `${paragraphStartTag(source, id)}${pPr}${content}</w:p>`;
}

/**
 * Patches that declare the `w14` namespace on the document root and list
 * it as ignorable, when the root does not yet; nothing when it does.
 */
export function namespacePatches(part: XmlPart): XmlPatch[] {
  const root = part.root;
  const items: XmlPatch[] = [];
  const has = (name: string): string | undefined =>
    root.attributes.find((attribute) => attribute.name === name)?.value;
  if (has("xmlns:w14") === undefined)
    items.push(patches.setAttribute(part, root, "xmlns:w14", W14_NS));
  if (has("xmlns:mc") === undefined)
    items.push(patches.setAttribute(part, root, "xmlns:mc", MC_NS));
  const ignorable = has("mc:Ignorable");
  const listed = ignorable ? ignorable.split(/\s+/).filter(Boolean) : [];
  if (!listed.includes("w14"))
    items.push(
      patches.setAttribute(
        part,
        root,
        "mc:Ignorable",
        [...listed, "w14"].join(" "),
      ),
    );
  return items;
}

/** Whether the paragraph already carries a `w14:paraId`. */
export function isAuthored(part: XmlPart, paragraph: XmlElement): boolean {
  return part.attribute(paragraph, "w14:paraId") !== undefined;
}
