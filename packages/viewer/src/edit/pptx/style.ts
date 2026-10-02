import type { EditColor } from "../types.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import { EMU_PER_PT } from "./geometry.js";
import type { ThemeFonts } from "./model.js";
import type {
  PptxLineStyle,
  PptxShapeStyle,
  PptxTextAlign,
  PptxTextStyle,
} from "./types.js";

/*
 * Reading DrawingML styles: colours in their theme form, run properties
 * resolved through the list-style chain (run, shape, layout placeholder,
 * master placeholder, master text styles, presentation defaults), fills and
 * lines of a shape.
 */

/** A container of `lvl<N>pPr` children, with the part it lives in. */
export interface StyleSource {
  readonly part: XmlPart;
  readonly node: XmlElement;
}

export interface TextStyleInputs {
  /** `a:rPr` of the run. */
  readonly rPr?: StyleSource;
  /** `a:pPr` of the paragraph. */
  readonly pPr?: StyleSource;
  /** 0-based outline level of the paragraph. */
  readonly level: number;
  /** List styles in priority order. */
  readonly lists: readonly StyleSource[];
  readonly fonts: ThemeFonts;
}

const DEFAULT_SIZE = 18;
const DEFAULT_LINE_WIDTH_PT = 0.75;

const COLOR_MODIFIERS = new Set([
  "tint",
  "shade",
  "alpha",
  "lumMod",
  "lumOff",
  "satMod",
  "satOff",
  "hueMod",
  "hueOff",
  "comp",
  "inv",
  "gray",
]);

/** The colour a fill container holds: `a:solidFill` → a colour, `a:noFill` → "none". */
export function readFill(
  part: XmlPart,
  container: XmlElement | undefined,
): EditColor | "none" | undefined {
  if (!container) return undefined;
  for (const child of container.children) {
    if (child.local === "noFill") return "none";
    if (child.local === "solidFill") return readColor(part, child) ?? "auto";
    if (
      child.local === "gradFill" ||
      child.local === "blipFill" ||
      child.local === "pattFill" ||
      child.local === "grpFill"
    )
      return "auto";
  }
  return undefined;
}

/** The colour inside a `a:solidFill`-like node. */
export function readColor(
  part: XmlPart,
  holder: XmlElement,
): EditColor | undefined {
  const node = holder.children[0];
  if (!node) return undefined;
  switch (node.local) {
    case "srgbClr": {
      const value = part.attribute(node, "val");
      return value && /^[0-9A-Fa-f]{6}$/.test(value)
        ? `#${value.toUpperCase()}`
        : "auto";
    }
    case "schemeClr": {
      const theme = part.attribute(node, "val");
      if (!theme) return "auto";
      const mods: Record<string, number> = {};
      for (const child of node.children) {
        if (!COLOR_MODIFIERS.has(child.local)) continue;
        const value = Number(part.attribute(child, "val") ?? "0");
        mods[child.local] = Number.isFinite(value) ? value : 0;
      }
      return Object.keys(mods).length > 0 ? { theme, mods } : { theme };
    }
    case "sysClr": {
      const value = part.attribute(node, "lastClr");
      return value && /^[0-9A-Fa-f]{6}$/.test(value)
        ? `#${value.toUpperCase()}`
        : "auto";
    }
    default:
      return "auto";
  }
}

/** Fill and line of a `p:spPr`; absent fields are inherited from the style or the theme. */
export function readShapeStyle(
  part: XmlPart,
  spPr: XmlElement | undefined,
): PptxShapeStyle {
  if (!spPr) return {};
  const fill = readFill(part, spPr);
  const ln = spPr.children.find((child) => child.local === "ln");
  let line: PptxLineStyle | undefined;
  if (ln) {
    const color = readFill(part, ln);
    if (color !== undefined) {
      const w = Number(part.attribute(ln, "w") ?? NaN);
      line = {
        color,
        width: Number.isFinite(w) ? w / EMU_PER_PT : DEFAULT_LINE_WIDTH_PT,
      };
    }
  }
  return { ...(fill !== undefined ? { fill } : {}), ...(line ? { line } : {}) };
}

function levelProperties(
  source: StyleSource,
  level: number,
): XmlElement | undefined {
  const name = `lvl${level + 1}pPr`;
  return source.node.children.find((child) => child.local === name);
}

function defaultRunProperties(
  source: StyleSource,
  level: number,
): XmlElement | undefined {
  return levelProperties(source, level)?.children.find(
    (child) => child.local === "defRPr",
  );
}

/** Theme font references resolved to the master's faces; other names pass through. */
export function resolveFontName(name: string, fonts: ThemeFonts): string {
  if (name.startsWith("+mj")) return fonts.major;
  if (name.startsWith("+mn")) return fonts.minor;
  return name;
}

export function alignOf(value: string | undefined): PptxTextAlign | undefined {
  switch (value) {
    case "l":
      return "left";
    case "ctr":
      return "center";
    case "r":
      return "right";
    case "just":
    case "justLow":
    case "dist":
    case "thaiDist":
      return "justify";
    default:
      return undefined;
  }
}

/** Resolves the style of one run through the chain; defaults follow PowerPoint. */
export function resolveTextStyle(inputs: TextStyleInputs): PptxTextStyle {
  const runSources: StyleSource[] = [];
  if (inputs.rPr) runSources.push(inputs.rPr);
  for (const list of inputs.lists) {
    const node = defaultRunProperties(list, inputs.level);
    if (node) runSources.push({ part: list.part, node });
  }
  const attribute = (name: string): string | undefined => {
    for (const source of runSources) {
      const value = source.part.attribute(source.node, name);
      if (value !== undefined) return value;
    }
    return undefined;
  };
  const flag = (name: string): boolean => {
    const value = attribute(name);
    return value === "1" || value === "true";
  };
  const size = Number(attribute("sz") ?? NaN);
  let fontFamily: string | undefined;
  let color: EditColor | undefined;
  for (const source of runSources) {
    if (fontFamily === undefined) {
      const latin = source.node.children.find(
        (child) => child.local === "latin",
      );
      const face = latin ? source.part.attribute(latin, "typeface") : undefined;
      if (face) fontFamily = face;
    }
    if (color === undefined) {
      const fill = readFill(source.part, source.node);
      if (fill !== undefined && fill !== "none") color = fill;
    }
    if (fontFamily !== undefined && color !== undefined) break;
  }
  const underline = attribute("u");
  let align: PptxTextAlign | undefined = inputs.pPr
    ? alignOf(inputs.pPr.part.attribute(inputs.pPr.node, "algn"))
    : undefined;
  if (!align)
    for (const list of inputs.lists) {
      const node = levelProperties(list, inputs.level);
      align = node ? alignOf(list.part.attribute(node, "algn")) : undefined;
      if (align) break;
    }
  return {
    fontFamily: resolveFontName(fontFamily ?? "+mn-lt", inputs.fonts),
    fontSize: Number.isFinite(size) ? size / 100 : DEFAULT_SIZE,
    bold: flag("b"),
    italic: flag("i"),
    underline: underline !== undefined && underline !== "none",
    color: color ?? { theme: "tx1" },
    align: align ?? "left",
  };
}
