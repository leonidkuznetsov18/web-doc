import { ViewerError } from "../../errors.js";
import type { OoxmlPackage } from "../ooxml/package.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";

/*
 * The deck index: which parts are the presentation, its slides in order,
 * their layouts, the masters and their themes. Built from the package at a
 * revision and rebuilt when the revision changes; everything else about a
 * slide is read on demand.
 */

export const PRESENTATION_NS =
  "http://schemas.openxmlformats.org/presentationml/2006/main";

const OFFICE_RELATIONSHIPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

export const RELATIONSHIP_TYPES = {
  officeDocument: `${OFFICE_RELATIONSHIPS}officeDocument`,
  slide: `${OFFICE_RELATIONSHIPS}slide`,
  slideLayout: `${OFFICE_RELATIONSHIPS}slideLayout`,
  slideMaster: `${OFFICE_RELATIONSHIPS}slideMaster`,
  notesSlide: `${OFFICE_RELATIONSHIPS}notesSlide`,
  theme: `${OFFICE_RELATIONSHIPS}theme`,
  image: `${OFFICE_RELATIONSHIPS}image`,
} as const;

export interface MasterRecord {
  /** "master1", from the part's number. */
  readonly id: string;
  readonly part: string;
  readonly theme?: string;
  readonly layouts: readonly string[];
}

export interface LayoutRecord {
  /** "layout2", from the part's number. */
  readonly id: string;
  readonly part: string;
  readonly name: string;
  readonly type?: string;
  readonly master: MasterRecord;
}

export interface SlideRecord {
  /** "sld3", from the part's number; the slide key of element ids. */
  readonly key: string;
  readonly number: number;
  readonly part: string;
  /** `p:sldId/@id`. */
  readonly id: string;
  /** The presentation relationship that names the part. */
  readonly rId: string;
  readonly node: XmlElement;
  readonly layout?: LayoutRecord;
}

export interface ThemeFonts {
  readonly major: string;
  readonly minor: string;
}

/** The number in a part name like `/ppt/slides/slide12.xml`. */
export function partNumber(part: string): number {
  const match = /(\d+)\.xml$/i.exec(part);
  return match ? Number(match[1]) : 0;
}

export class DeckModel {
  readonly #pkg: OoxmlPackage;
  readonly #themeFonts = new Map<string, Promise<ThemeFonts>>();

  private constructor(
    pkg: OoxmlPackage,
    readonly revision: number,
    readonly presentationPart: string,
    readonly presentation: XmlPart,
    /** EMU. */
    readonly slideSize: { readonly cx: number; readonly cy: number },
    readonly slides: readonly SlideRecord[],
    readonly layouts: readonly LayoutRecord[],
    readonly masters: readonly MasterRecord[],
  ) {
    this.#pkg = pkg;
  }

  static async load(
    pkg: OoxmlPackage,
    signal?: AbortSignal,
  ): Promise<DeckModel> {
    const root = await pkg.relationships("/", signal);
    const main = root.byType(RELATIONSHIP_TYPES.officeDocument)[0]?.targetPart;
    if (!main)
      throw new ViewerError(
        "invalid-file",
        "The package has no presentation part",
      );
    const presentation = await pkg.xml(main, signal);
    const size = presentation.find("sldSz");
    const slideSize = {
      cx: Number(size ? presentation.attribute(size, "cx") : NaN) || 9144000,
      cy: Number(size ? presentation.attribute(size, "cy") : NaN) || 6858000,
    };
    const rels = await pkg.relationships(main, signal);

    const masters: MasterRecord[] = [];
    const layouts: LayoutRecord[] = [];
    const layoutByPart = new Map<string, LayoutRecord>();
    for (const masterId of presentation.findAll("sldMasterId")) {
      const rId = presentation.attribute(masterId, "r:id");
      const part = rId ? rels.byId(rId)?.targetPart : undefined;
      if (!part || !pkg.has(part)) continue;
      const masterRels = await pkg.relationships(part, signal);
      const theme = masterRels.byType(RELATIONSHIP_TYPES.theme)[0]?.targetPart;
      const layoutParts = masterRels
        .byType(RELATIONSHIP_TYPES.slideLayout)
        .map((relationship) => relationship.targetPart)
        .filter((name): name is string => !!name && pkg.has(name))
        .sort((a, b) => partNumber(a) - partNumber(b));
      const master: MasterRecord = {
        id: `master${partNumber(part)}`,
        part,
        ...(theme ? { theme } : {}),
        layouts: layoutParts,
      };
      masters.push(master);
      for (const layoutPart of layoutParts) {
        if (layoutByPart.has(layoutPart)) continue;
        const xml = await pkg.xml(layoutPart, signal);
        const cSld = xml.find("cSld");
        const type = xml.attribute(xml.root, "type");
        const layout: LayoutRecord = {
          id: `layout${partNumber(layoutPart)}`,
          part: layoutPart,
          name: (cSld ? xml.attribute(cSld, "name") : undefined) ?? "",
          ...(type ? { type } : {}),
          master,
        };
        layouts.push(layout);
        layoutByPart.set(layoutPart, layout);
      }
    }

    const slides: SlideRecord[] = [];
    for (const node of presentation.findAll("sldId")) {
      // Section lists (p14:sldId) name slides too; only p:sldIdLst lists them.
      if (node.namespace !== PRESENTATION_NS) continue;
      const rId = presentation.attribute(node, "r:id");
      const id = presentation.attribute(node, "id");
      const part = rId ? rels.byId(rId)?.targetPart : undefined;
      if (!rId || !id || !part || !pkg.has(part))
        throw new ViewerError(
          "invalid-file",
          `Slide ${id ?? "?"} of the presentation has no part`,
        );
      const slideRels = await pkg.relationships(part, signal);
      const layoutPart = slideRels.byType(RELATIONSHIP_TYPES.slideLayout)[0]
        ?.targetPart;
      const layout = layoutPart ? layoutByPart.get(layoutPart) : undefined;
      const number = partNumber(part);
      if (slides.some((slide) => slide.number === number))
        throw new ViewerError(
          "invalid-file",
          `Two slide parts share the number ${number}`,
          { details: { part } },
        );
      slides.push({
        key: `sld${number}`,
        number,
        part,
        id,
        rId,
        node,
        ...(layout ? { layout } : {}),
      });
    }
    return new DeckModel(
      pkg,
      pkg.revision,
      main,
      presentation,
      slideSize,
      slides,
      layouts,
      masters,
    );
  }

  get pageCount(): number {
    return this.slides.length;
  }

  slideAt(pageIndex: number): SlideRecord {
    const slide = this.slides[pageIndex];
    if (!Number.isInteger(pageIndex) || !slide)
      throw new ViewerError(
        "invalid-operation",
        `Slide index ${pageIndex} is out of range`,
        { details: { pageIndex, pageCount: this.slides.length } },
      );
    return slide;
  }

  slideByKey(key: string): SlideRecord | undefined {
    return this.slides.find((slide) => slide.key === key);
  }

  layoutById(id: string): LayoutRecord | undefined {
    return this.layouts.find((layout) => layout.id === id);
  }

  /** Whether a slide is hidden (`p:sld/@show="0"`). */
  async hidden(slide: SlideRecord, signal?: AbortSignal): Promise<boolean> {
    const xml = await this.#pkg.xml(slide.part, signal);
    return xml.attribute(xml.root, "show") === "0";
  }

  /** The major and minor Latin faces of a master's theme. */
  themeFonts(master: MasterRecord | undefined): Promise<ThemeFonts> {
    const theme = master?.theme;
    if (!theme || !this.#pkg.has(theme))
      return Promise.resolve({ major: "Calibri Light", minor: "Calibri" });
    let fonts = this.#themeFonts.get(theme);
    if (!fonts) {
      fonts = this.#pkg.xml(theme).then((xml) => {
        const face = (scheme: string): string | undefined => {
          const node = xml.find(scheme);
          const latin = node
            ? node.children.find((child) => child.local === "latin")
            : undefined;
          return latin ? xml.attribute(latin, "typeface") : undefined;
        };
        return {
          major: face("majorFont") || "Calibri Light",
          minor: face("minorFont") || "Calibri",
        };
      });
      this.#themeFonts.set(theme, fonts);
    }
    return fonts;
  }
}
