import type { OoxmlPackage } from "../ooxml/package.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { EditColor } from "../types.js";
import { OFFICE_RELATIONSHIPS, W_NS } from "./ids.js";
import type {
  DocxHighlight,
  DocxParagraphSpacing,
  DocxParagraphStyle,
  DocxTextAlign,
  DocxTextStyle,
} from "./types.js";

/*
 * Resolved styles of WordprocessingML. A run property comes from the first
 * of: the run's own `w:rPr`, its character style chain (`w:rStyle`, then
 * `w:basedOn` up), the paragraph style chain's run properties, the default
 * paragraph style, the document defaults, Word's built-in defaults. A
 * paragraph property comes from `w:pPr`, the paragraph style chain, the
 * default paragraph style, the document defaults. Theme fonts and colours
 * are reported in their theme form where the file uses them.
 */

const HIGHLIGHTS = new Set<string>([
  "yellow",
  "green",
  "cyan",
  "magenta",
  "blue",
  "red",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "black",
  "white",
]);

export interface StyleRecord {
  readonly id: string;
  readonly type: string;
  readonly basedOn?: string;
  readonly isDefault: boolean;
  readonly rPr?: XmlElement;
  readonly pPr?: XmlElement;
}

export interface ThemeFonts {
  readonly major: string;
  readonly minor: string;
}

/** `#RRGGBB` per colour scheme slot (`dk1`, `lt1`, `accent1`, …). */
export type ThemeColors = ReadonlyMap<string, string>;

const BUILTIN_FONTS: ThemeFonts = { major: "Calibri Light", minor: "Calibri" };

/** Word's `w:themeColor` names to the theme's colour scheme slots. */
const THEME_SLOTS: Readonly<Record<string, string>> = {
  dark1: "dk1",
  text1: "dk1",
  light1: "lt1",
  background1: "lt1",
  dark2: "dk2",
  text2: "dk2",
  light2: "lt2",
  background2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
};

/** Whether a name is one Word accepts in `w:themeColor`. */
export function isThemeColorName(name: string): boolean {
  return Object.hasOwn(THEME_SLOTS, name);
}

/** The styles and theme of a document, read once per revision. */
export class DocxStyles {
  private constructor(
    readonly part: XmlPart | undefined,
    readonly styles: ReadonlyMap<string, StyleRecord>,
    readonly defaultParagraph: StyleRecord | undefined,
    readonly defaultCharacter: StyleRecord | undefined,
    readonly docDefaultsRPr: XmlElement | undefined,
    readonly docDefaultsPPr: XmlElement | undefined,
    readonly fonts: ThemeFonts,
    readonly colors: ThemeColors,
  ) {}

  /** The `RRGGBB` (no `#`, as `w:val` takes it) a `w:themeColor` name resolves to in this theme, when it has one. */
  themeColor(name: string): string | undefined {
    const slot = THEME_SLOTS[name];
    return slot ? this.colors.get(slot) : undefined;
  }

  static async load(
    pkg: OoxmlPackage,
    documentPart: string,
    signal?: AbortSignal,
  ): Promise<DocxStyles> {
    const rels = await pkg.relationships(documentPart, signal);
    const stylesPart = rels.byType(`${OFFICE_RELATIONSHIPS}styles`)[0]
      ?.targetPart;
    const themePart = rels.byType(`${OFFICE_RELATIONSHIPS}theme`)[0]
      ?.targetPart;
    const { fonts, colors } = await readTheme(pkg, themePart, signal);
    if (!stylesPart || !pkg.has(stylesPart))
      return new DocxStyles(
        undefined,
        new Map(),
        undefined,
        undefined,
        undefined,
        undefined,
        fonts,
        colors,
      );
    const part = await pkg.xml(stylesPart, signal);
    const styles = new Map<string, StyleRecord>();
    let defaultParagraph: StyleRecord | undefined;
    let defaultCharacter: StyleRecord | undefined;
    for (const node of part.root.children) {
      if (node.namespace !== W_NS || node.local !== "style") continue;
      const id = part.attribute(node, "w:styleId");
      if (!id) continue;
      const basedOn = part.attribute(
        node.children.find((child) => child.local === "basedOn") ?? node,
        "w:val",
      );
      const defaultAttribute = part.attribute(node, "w:default");
      const record: StyleRecord = {
        id,
        type: part.attribute(node, "w:type") ?? "paragraph",
        ...(basedOn && basedOn !== id ? { basedOn } : {}),
        isDefault: defaultAttribute !== undefined && isOn(defaultAttribute),
        ...pick(node, "rPr", "pPr"),
      };
      styles.set(id, record);
      if (record.isDefault && record.type === "paragraph" && !defaultParagraph)
        defaultParagraph = record;
      if (record.isDefault && record.type === "character" && !defaultCharacter)
        defaultCharacter = record;
    }
    const docDefaults = part.root.children.find(
      (child) => child.local === "docDefaults",
    );
    const rPrDefault = docDefaults?.children.find(
      (child) => child.local === "rPrDefault",
    );
    const pPrDefault = docDefaults?.children.find(
      (child) => child.local === "pPrDefault",
    );
    return new DocxStyles(
      part,
      styles,
      defaultParagraph,
      defaultCharacter,
      rPrDefault?.children.find((child) => child.local === "rPr"),
      pPrDefault?.children.find((child) => child.local === "pPr"),
      fonts,
      colors,
    );
  }

  /** A style and the styles it is based on, nearest first; cycles stop. */
  chain(id: string | undefined): StyleRecord[] {
    const out: StyleRecord[] = [];
    const seen = new Set<string>();
    let current = id;
    while (current && !seen.has(current)) {
      seen.add(current);
      const record = this.styles.get(current);
      if (!record) break;
      out.push(record);
      current = record.basedOn;
    }
    return out;
  }
}

function pick(
  node: XmlElement,
  ...names: readonly ("rPr" | "pPr")[]
): { rPr?: XmlElement; pPr?: XmlElement } {
  const out: { rPr?: XmlElement; pPr?: XmlElement } = {};
  for (const name of names) {
    const child = node.children.find(
      (candidate) => candidate.local === name && candidate.namespace === W_NS,
    );
    if (child) out[name] = child;
  }
  return out;
}

async function readTheme(
  pkg: OoxmlPackage,
  themePart: string | undefined,
  signal?: AbortSignal,
): Promise<{ fonts: ThemeFonts; colors: ThemeColors }> {
  const none = { fonts: BUILTIN_FONTS, colors: new Map<string, string>() };
  if (!themePart || !pkg.has(themePart)) return none;
  try {
    const xml = await pkg.xml(themePart, signal);
    const face = (scheme: string): string | undefined => {
      const node = xml.find(scheme);
      const latin = node?.children.find((child) => child.local === "latin");
      return latin ? xml.attribute(latin, "typeface") : undefined;
    };
    const colors = new Map<string, string>();
    const scheme = xml.find("clrScheme");
    for (const slot of scheme?.children ?? []) {
      const value = slot.children[0];
      const rgb = value
        ? value.local === "srgbClr"
          ? xml.attribute(value, "val")
          : value.local === "sysClr"
            ? xml.attribute(value, "lastClr")
            : undefined
        : undefined;
      if (rgb && /^[0-9A-Fa-f]{6}$/.test(rgb))
        colors.set(slot.local, rgb.toUpperCase());
    }
    return {
      fonts: {
        major: face("majorFont") || BUILTIN_FONTS.major,
        minor: face("minorFont") || BUILTIN_FONTS.minor,
      },
      colors,
    };
  } catch {
    return none;
  }
}

/** `w:b`/`w:i` toggle values: a present element without `w:val` means on. */
export function isOn(value: string | undefined): boolean {
  return value === undefined || !/^(0|false|off)$/i.test(value);
}

/** The run property holders in resolution order for a run of a paragraph. */
export function runPropertySources(
  styles: DocxStyles,
  pPr: XmlElement | undefined,
  rPr: XmlElement | undefined,
): readonly XmlElement[] {
  const out: XmlElement[] = [];
  if (rPr) out.push(rPr);
  const rStyle = rPr
    ? attributeOf(
        rPr.children.find((child) => child.local === "rStyle"),
        "w:val",
      )
    : undefined;
  for (const record of styles.chain(rStyle))
    if (record.rPr) out.push(record.rPr);
  const pStyle = pPr
    ? attributeOf(
        pPr.children.find((child) => child.local === "pStyle"),
        "w:val",
      )
    : undefined;
  for (const record of styles.chain(pStyle ?? styles.defaultParagraph?.id))
    if (record.rPr) out.push(record.rPr);
  if (
    pStyle &&
    styles.defaultParagraph &&
    !styles.chain(pStyle).includes(styles.defaultParagraph) &&
    styles.defaultParagraph.rPr
  )
    out.push(styles.defaultParagraph.rPr);
  if (styles.docDefaultsRPr) out.push(styles.docDefaultsRPr);
  return out;
}

/** The paragraph property holders in resolution order. */
export function paragraphPropertySources(
  styles: DocxStyles,
  pPr: XmlElement | undefined,
): readonly XmlElement[] {
  const out: XmlElement[] = [];
  if (pPr) out.push(pPr);
  const pStyle = pPr
    ? attributeOf(
        pPr.children.find((child) => child.local === "pStyle"),
        "w:val",
      )
    : undefined;
  const chain = styles.chain(pStyle ?? styles.defaultParagraph?.id);
  for (const record of chain) if (record.pPr) out.push(record.pPr);
  if (
    pStyle &&
    styles.defaultParagraph &&
    !chain.includes(styles.defaultParagraph) &&
    styles.defaultParagraph.pPr
  )
    out.push(styles.defaultParagraph.pPr);
  if (styles.docDefaultsPPr) out.push(styles.docDefaultsPPr);
  return out;
}

function attributeOf(
  node: XmlElement | undefined,
  name: string,
): string | undefined {
  return node?.attributes.find((attribute) => attribute.name === name)?.value;
}

/** The first holder that has a `name` child, with that child. */
function firstChild(
  sources: readonly XmlElement[],
  name: string,
): XmlElement | undefined {
  for (const source of sources) {
    const child = source.children.find(
      (candidate) => candidate.local === name && candidate.namespace === W_NS,
    );
    if (child) return child;
  }
  return undefined;
}

export function resolveTextStyle(
  styles: DocxStyles,
  pPr: XmlElement | undefined,
  rPr: XmlElement | undefined,
): DocxTextStyle {
  const sources = runPropertySources(styles, pPr, rPr);
  const toggle = (name: string): boolean => {
    const node = firstChild(sources, name);
    return node ? isOn(attributeOf(node, "w:val")) : false;
  };
  const sz = firstChild(sources, "sz");
  const half = Number(sz ? attributeOf(sz, "w:val") : NaN);
  const underlineNode = firstChild(sources, "u");
  const underline = underlineNode
    ? (attributeOf(underlineNode, "w:val") ?? "single") !== "none"
    : false;
  const highlightNode = firstChild(sources, "highlight");
  const highlight = highlightNode
    ? attributeOf(highlightNode, "w:val")
    : undefined;
  return {
    fontFamily: resolveFont(sources, styles.fonts),
    fontSize: Number.isFinite(half) && half > 0 ? half / 2 : 10,
    bold: toggle("b"),
    italic: toggle("i"),
    underline,
    color: resolveColor(firstChild(sources, "color")),
    ...(highlight && HIGHLIGHTS.has(highlight) && highlight !== "none"
      ? { highlight: highlight as DocxHighlight }
      : {}),
  };
}

function resolveFont(
  sources: readonly XmlElement[],
  fonts: ThemeFonts,
): string {
  for (const source of sources) {
    const rFonts = source.children.find(
      (child) => child.local === "rFonts" && child.namespace === W_NS,
    );
    if (!rFonts) continue;
    const theme =
      attributeOf(rFonts, "w:asciiTheme") ??
      attributeOf(rFonts, "w:hAnsiTheme");
    if (theme) return theme.startsWith("major") ? fonts.major : fonts.minor;
    const face =
      attributeOf(rFonts, "w:ascii") ?? attributeOf(rFonts, "w:hAnsi");
    if (face) return face;
  }
  return fonts.minor;
}

function resolveColor(node: XmlElement | undefined): EditColor {
  if (!node) return "auto";
  const theme = attributeOf(node, "w:themeColor");
  if (theme) return { theme };
  const value = attributeOf(node, "w:val");
  return value && /^[0-9A-Fa-f]{6}$/.test(value)
    ? `#${value.toUpperCase()}`
    : "auto";
}

export function alignOf(value: string | undefined): DocxTextAlign | undefined {
  switch (value) {
    case "left":
    case "start":
      return "left";
    case "center":
      return "center";
    case "right":
    case "end":
      return "right";
    case "both":
    case "distribute":
    case "thaiDistribute":
    case "lowKashida":
    case "mediumKashida":
    case "highKashida":
      return "justify";
    default:
      return undefined;
  }
}

export function resolveParagraphStyle(
  styles: DocxStyles,
  pPr: XmlElement | undefined,
): DocxParagraphStyle {
  const sources = paragraphPropertySources(styles, pPr);
  const styleId = pPr
    ? attributeOf(
        pPr.children.find((child) => child.local === "pStyle"),
        "w:val",
      )
    : undefined;
  const jc = firstChild(sources, "jc");
  const numPr = firstChild(sources, "numPr");
  // Spacing attributes inherit one by one: a style that sets `w:before`
  // keeps the base style's `w:after` and `w:line`.
  const spacings = sources
    .map((source) =>
      source.children.find(
        (child) => child.local === "spacing" && child.namespace === W_NS,
      ),
    )
    .filter((node): node is XmlElement => node !== undefined);
  const spacingAttribute = (name: string): string | undefined => {
    for (const node of spacings) {
      const value = attributeOf(node, name);
      if (value !== undefined) return value;
    }
    return undefined;
  };
  const twips = (name: string): number | undefined => {
    const value = Number(spacingAttribute(name) ?? NaN);
    return Number.isFinite(value) ? value / 20 : undefined;
  };
  const before = twips("w:before");
  const after = twips("w:after");
  const lineHolder = spacings.find(
    (node) => attributeOf(node, "w:line") !== undefined,
  );
  const lineRaw = Number(lineHolder ? attributeOf(lineHolder, "w:line") : NaN);
  const lineRule = lineHolder
    ? (attributeOf(lineHolder, "w:lineRule") ?? "auto")
    : undefined;
  const line = Number.isFinite(lineRaw)
    ? lineRule === "auto"
      ? lineRaw / 240
      : lineRaw / 20
    : undefined;
  const out: DocxParagraphSpacing = {
    ...(before === undefined ? {} : { before }),
    ...(after === undefined ? {} : { after }),
    ...(line === undefined ? {} : { line }),
    ...(line === undefined ||
    (lineRule !== "auto" && lineRule !== "exact" && lineRule !== "atLeast")
      ? {}
      : { lineRule }),
  };
  const numId = Number(
    numPr
      ? attributeOf(
          numPr.children.find((c) => c.local === "numId"),
          "w:val",
        )
      : NaN,
  );
  const level = Number(
    numPr
      ? attributeOf(
          numPr.children.find((c) => c.local === "ilvl"),
          "w:val",
        )
      : 0,
  );
  return {
    ...(styleId ? { styleId } : {}),
    align: alignOf(jc ? attributeOf(jc, "w:val") : undefined) ?? "left",
    spacing: out,
    ...(Number.isInteger(numId) && numId > 0
      ? { numbering: { numId, level: Number.isInteger(level) ? level : 0 } }
      : {}),
  };
}
