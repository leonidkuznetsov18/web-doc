import type { ResourceLimits } from "../contracts.js";
import { OoxmlPackage } from "../edit/ooxml/package.js";
import { patches, type XmlPatch } from "../edit/ooxml/patch.js";
import type { XmlElement, XmlPart } from "../edit/ooxml/xml.js";

/*
 * The DOCX display pre-pass: what the renderer should see instead of the
 * file as stored, written as patches of the XML parts through the package
 * layer and never into what an editor saves.
 *
 * 1. Oversized inline pictures are scaled down to their section's content
 *    box, aspect ratio kept, as Word's own layout would not: Word draws a
 *    21-inch picture on a 6.5-inch column clipped, a viewer has no margin
 *    to spill into. Anchored pictures keep their geometry.
 * 2. Every paragraph is marked with a hidden bookmark that carries its id —
 *    the file's own `w14:paraId` when it has one, else a deterministic id
 *    from its position — because the renderer keeps bookmark names on its
 *    model paragraphs but does not read `w14:paraId`. An editor that walks
 *    the same paragraphs in the same order computes the same ids from the
 *    original bytes.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WP_NS =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const RELATIONSHIPS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
const STORY_RELATIONSHIPS = [
  `${RELATIONSHIPS}header`,
  `${RELATIONSHIPS}footer`,
  `${RELATIONSHIPS}footnotes`,
  `${RELATIONSHIPS}endnotes`,
];

/** EMU per twentieth of a point, the unit of page sizes and margins. */
const EMU_PER_TWIP = 635;
/** US Letter with one-inch margins, Word's default when a section says nothing. */
const DEFAULT_SECTION = {
  pageWidth: 12240,
  pageHeight: 15840,
  marginLeft: 1440,
  marginRight: 1440,
  marginTop: 1440,
  marginBottom: 1440,
};

export interface DocxPrepassResult {
  /** The bytes to render; the input when nothing had to change. */
  readonly bytes: Uint8Array;
  /** Inline pictures scaled down. */
  readonly scaledImages: number;
  /** Paragraphs marked with an id bookmark. */
  readonly markedParagraphs: number;
  /** Of those, paragraphs whose id was generated (no `w14:paraId` in the file). */
  readonly generatedIds: number;
}

/** Prefix of the hidden bookmarks that carry paragraph ids. */
export const PARAGRAPH_BOOKMARK_PREFIX = "_wd";
/** First bookmark id the pre-pass uses, above what Word writes. */
const BOOKMARK_ID_BASE = 7_000_000;

/** First generated id; Word keeps `w14:paraId` below 0x80000000. */
const ID_BASE = 0x1a000000;
const ID_STEP = 0x9e37;

/**
 * The paragraph id generated for the `index`-th paragraph without one, in
 * document order of a part, skipping ids the part already uses.
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

/**
 * Prepares DOCX bytes for display. A document the package layer cannot
 * open or scan is returned as it is: the renderer reports its own errors.
 */
export async function prepareDocxForDisplay(
  bytes: Uint8Array,
  limits: ResourceLimits,
  signal?: AbortSignal,
): Promise<DocxPrepassResult> {
  const unchanged = {
    bytes,
    scaledImages: 0,
    markedParagraphs: 0,
    generatedIds: 0,
  };
  let pkg: OoxmlPackage;
  try {
    pkg = await OoxmlPackage.open(bytes, {
      limits,
      ...(signal ? { signal } : {}),
    });
  } catch {
    return unchanged;
  }
  try {
    const root = await pkg.relationships("/", signal);
    const main = root.byType(`${RELATIONSHIPS}officeDocument`)[0]?.targetPart;
    if (!main || !pkg.has(main)) return unchanged;
    const transaction = pkg.transaction();
    let scaledImages = 0;
    let markedParagraphs = 0;
    let generatedIds = 0;
    const counted = (marked: number, generated: number): void => {
      markedParagraphs += marked;
      generatedIds += generated;
    };
    const document = await pkg.xml(main, signal);
    const documentPatches: XmlPatch[] = [
      ...fitInlinePictures(document, (count) => (scaledImages += count)),
    ];
    documentPatches.push(...paragraphIdPatches(document, counted));
    if (documentPatches.length > 0)
      transaction.patch(document, documentPatches);
    const relationships = await pkg.relationships(main, signal);
    for (const type of STORY_RELATIONSHIPS)
      for (const item of relationships.byType(type)) {
        if (!item.targetPart || !pkg.has(item.targetPart)) continue;
        const part = await pkg.xml(item.targetPart, signal);
        const items = paragraphIdPatches(part, counted);
        if (items.length > 0) transaction.patch(part, items);
      }
    if (scaledImages === 0 && markedParagraphs === 0) return unchanged;
    await transaction.commit(signal);
    return {
      bytes: await pkg.save({}, signal),
      scaledImages,
      markedParagraphs,
      generatedIds,
    };
  } catch {
    return unchanged;
  }
}

interface SectionGeometry {
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
}

function sectionGeometry(
  part: XmlPart,
  sectPr: XmlElement | undefined,
): SectionGeometry {
  const read = (
    node: XmlElement | undefined,
    name: string,
    fallback: number,
  ): number => {
    const value = node ? Number(part.attribute(node, name) ?? NaN) : NaN;
    return Number.isFinite(value) && value >= 0 ? value : fallback;
  };
  const pgSz = sectPr?.children.find(
    (child) => child.local === "pgSz" && child.namespace === W_NS,
  );
  const pgMar = sectPr?.children.find(
    (child) => child.local === "pgMar" && child.namespace === W_NS,
  );
  return {
    pageWidth: read(pgSz, "w:w", DEFAULT_SECTION.pageWidth),
    pageHeight: read(pgSz, "w:h", DEFAULT_SECTION.pageHeight),
    marginLeft: read(pgMar, "w:left", DEFAULT_SECTION.marginLeft),
    marginRight: read(pgMar, "w:right", DEFAULT_SECTION.marginRight),
    marginTop: read(pgMar, "w:top", DEFAULT_SECTION.marginTop),
    marginBottom: read(pgMar, "w:bottom", DEFAULT_SECTION.marginBottom),
  };
}

/** Patches that scale the oversized inline pictures of the body, section by section. */
function fitInlinePictures(
  part: XmlPart,
  count: (scaled: number) => void,
): XmlPatch[] {
  const body = part.root.children.find(
    (child) => child.local === "body" && child.namespace === W_NS,
  );
  if (!body) return [];
  const items: XmlPatch[] = [];
  // A w:sectPr closes the section that ends at it: the geometry of a run of
  // body elements is known once the paragraph carrying the break, or the
  // body's own w:sectPr, is reached.
  let pending: XmlElement[] = [];
  const flush = (sectPr: XmlElement | undefined): void => {
    const geometry = sectionGeometry(part, sectPr);
    for (const element of pending)
      items.push(...fitPicturesUnder(part, element, geometry, count));
    pending = [];
  };
  let bodySectPr: XmlElement | undefined;
  for (const element of body.children) {
    if (element.namespace === W_NS && element.local === "sectPr") {
      bodySectPr = element;
      continue;
    }
    pending.push(element);
    const sectPr =
      element.namespace === W_NS && element.local === "p"
        ? element.children
            .find((child) => child.local === "pPr" && child.namespace === W_NS)
            ?.children.find(
              (child) => child.local === "sectPr" && child.namespace === W_NS,
            )
        : undefined;
    if (sectPr) flush(sectPr);
  }
  flush(bodySectPr);
  return items;
}

function fitPicturesUnder(
  part: XmlPart,
  element: XmlElement,
  geometry: SectionGeometry,
  count: (scaled: number) => void,
): XmlPatch[] {
  const contentWidth =
    (geometry.pageWidth - geometry.marginLeft - geometry.marginRight) *
    EMU_PER_TWIP;
  const contentHeight =
    (geometry.pageHeight - geometry.marginTop - geometry.marginBottom) *
    EMU_PER_TWIP;
  if (!(contentWidth > 0 && contentHeight > 0)) return [];
  const items: XmlPatch[] = [];
  for (const inline of part.findAll("inline", element)) {
    if (inline.namespace !== WP_NS) continue;
    const extent = inline.children.find(
      (child) => child.local === "extent" && child.namespace === WP_NS,
    );
    if (!extent) continue;
    const cx = Number(part.attribute(extent, "cx") ?? NaN);
    const cy = Number(part.attribute(extent, "cy") ?? NaN);
    if (!(cx > 0 && cy > 0)) continue;
    const scale = Math.min(1, contentWidth / cx, contentHeight / cy);
    if (scale >= 1) continue;
    const width = String(Math.round(cx * scale));
    const height = String(Math.round(cy * scale));
    items.push(
      patches.setAttribute(part, extent, "cx", width),
      patches.setAttribute(part, extent, "cy", height),
    );
    // The picture's own transform repeats the extent; keep both in step.
    const ext = part
      .findAll("ext", inline)
      .find(
        (candidate) =>
          candidate.namespace === A_NS &&
          candidate.parent?.local === "xfrm" &&
          part.attribute(candidate, "cx") === String(cx) &&
          part.attribute(candidate, "cy") === String(cy),
      );
    if (ext)
      items.push(
        patches.setAttribute(part, ext, "cx", width),
        patches.setAttribute(part, ext, "cy", height),
      );
    count(1);
  }
  return items;
}

/**
 * Patches that mark every `w:p` of a part with a hidden bookmark carrying
 * its id: `_wd` plus the file's `w14:paraId`, or a generated id for a
 * paragraph without one. The bookmark pair goes right after `w:pPr`, where
 * the schema allows it, with ids above any the part already uses.
 */
function paragraphIdPatches(
  part: XmlPart,
  count: (marked: number, generated: number) => void,
): XmlPatch[] {
  const paragraphs = part
    .findAll("p")
    .filter((node) => node.namespace === W_NS);
  if (paragraphs.length === 0) return [];
  const taken = new Set<string>();
  let bookmarkId = BOOKMARK_ID_BASE;
  for (const node of part.findAll("bookmarkStart")) {
    if (node.namespace !== W_NS) continue;
    const id = Number(part.attribute(node, "w:id") ?? NaN);
    if (Number.isInteger(id) && id >= bookmarkId) bookmarkId = id + 1;
    const name = part.attribute(node, "w:name") ?? "";
    if (name.startsWith(PARAGRAPH_BOOKMARK_PREFIX))
      taken.add(name.slice(PARAGRAPH_BOOKMARK_PREFIX.length).toUpperCase());
  }
  for (const paragraph of paragraphs) {
    const existing = part.attribute(paragraph, "w14:paraId");
    if (existing) taken.add(existing.toUpperCase());
  }
  const items: XmlPatch[] = [];
  let generated = 0;
  let marked = 0;
  for (const paragraph of paragraphs) {
    const authored = part.attribute(paragraph, "w14:paraId");
    const id = authored
      ? authored.toUpperCase()
      : generatedParagraphId(generated, taken);
    if (!authored) {
      taken.add(id);
      generated += 1;
    }
    const name = `${PARAGRAPH_BOOKMARK_PREFIX}${id}`;
    const start = `<w:bookmarkStart w:id="${bookmarkId}" w:name="${name}"/>`;
    const end = `<w:bookmarkEnd w:id="${bookmarkId}"/>`;
    bookmarkId += 1;
    const pPr = paragraph.children.find(
      (child) => child.local === "pPr" && child.namespace === W_NS,
    );
    if (pPr)
      items.push(
        patches.insertAfter(part, pPr, start),
        patches.insertAfter(part, pPr, end),
      );
    else if (paragraph.selfClosing || paragraph.children.length === 0)
      items.push(patches.replaceContent(part, paragraph, start + end));
    else
      items.push(
        patches.insertBefore(part, paragraph.children[0]!, start),
        patches.insertBefore(part, paragraph.children[0]!, end),
      );
    marked += 1;
  }
  count(marked, generated);
  return items;
}
