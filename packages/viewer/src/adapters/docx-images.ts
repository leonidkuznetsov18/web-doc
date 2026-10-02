/**
 * Word draws an inline picture at its declared extent even when that is wider
 * than the text area, so a generated document that embeds a 21-inch chart on
 * a 6.5-inch column shows a clipped picture. A viewer has no margin to spill
 * into, so the oversized inline pictures are scaled down to fit the section's
 * content box, aspect ratio preserved. Anchored (floating) pictures keep
 * their geometry: their position is part of the author's layout.
 *
 * @deprecated The viewer no longer calls this helper: since `@silurus/ooxml`
 * 0.88 lays a document out inside `load()`, adjusting the parsed model
 * changes nothing on the page. The fitting now happens in the XML before the
 * engine reads it (`prepareDocxForDisplay`). The helper stays exported for
 * hosts that applied it to their own `main`-mode models and is removed in
 * the next major version.
 */
export interface DocxSectionGeometry {
  readonly pageWidth: number;
  readonly pageHeight: number;
  readonly marginLeft: number;
  readonly marginRight: number;
  readonly marginTop: number;
  readonly marginBottom: number;
}

interface DocxImageRunLike {
  readonly type: string;
  readonly anchor?: boolean;
  widthPt: number;
  heightPt: number;
}

interface DocxParagraphLike {
  readonly type: "paragraph";
  readonly runs?: readonly unknown[];
}

interface DocxTableLike {
  readonly type: "table";
  readonly rows?: readonly {
    readonly cells?: readonly { readonly content?: readonly unknown[] }[];
  }[];
}

interface DocxSectionBreakLike {
  readonly type: "sectionBreak";
  readonly geom?: DocxSectionGeometry;
}

export interface DocxModelLike {
  readonly section: DocxSectionGeometry;
  readonly body: readonly unknown[];
}

/**
 * Shrink every inline picture that would not fit its section's content box.
 * Mutates the model in place and returns how many pictures were scaled.
 *
 * @deprecated See the module note: the viewer fits pictures in the XML
 * pre-pass instead; this helper is removed in the next major version.
 */
export function fitInlineImagesToPage(model: DocxModelLike): number {
  let scaled = 0;
  // A `<w:sectPr>` closes the section that ENDS at it, so the geometry for a
  // run of body elements is only known once the break after them is reached.
  let pending: unknown[] = [];
  const flush = (geometry: DocxSectionGeometry): void => {
    for (const element of pending) scaled += fitElement(element, geometry);
    pending = [];
  };
  for (const element of model.body) {
    if (isSectionBreak(element)) {
      flush(element.geom ?? model.section);
      continue;
    }
    pending.push(element);
  }
  flush(model.section);
  return scaled;
}

function fitElement(element: unknown, geometry: DocxSectionGeometry): number {
  if (isParagraph(element)) {
    let scaled = 0;
    for (const run of element.runs ?? [])
      if (isInlineImage(run) && fitImage(run, geometry)) scaled += 1;
    return scaled;
  }
  if (isTable(element)) {
    let scaled = 0;
    for (const row of element.rows ?? [])
      for (const cell of row.cells ?? [])
        for (const child of cell.content ?? [])
          scaled += fitElement(child, geometry);
    return scaled;
  }
  return 0;
}

function fitImage(
  image: DocxImageRunLike,
  geometry: DocxSectionGeometry,
): boolean {
  const contentWidth =
    geometry.pageWidth - geometry.marginLeft - geometry.marginRight;
  const contentHeight =
    geometry.pageHeight - geometry.marginTop - geometry.marginBottom;
  if (
    !(contentWidth > 0 && contentHeight > 0) ||
    !(image.widthPt > 0 && image.heightPt > 0)
  )
    return false;
  const scale = Math.min(
    1,
    contentWidth / image.widthPt,
    contentHeight / image.heightPt,
  );
  if (scale >= 1) return false;
  image.widthPt *= scale;
  image.heightPt *= scale;
  return true;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isParagraph(value: unknown): value is DocxParagraphLike {
  return isRecord(value) && value.type === "paragraph";
}

function isTable(value: unknown): value is DocxTableLike {
  return isRecord(value) && value.type === "table";
}

function isSectionBreak(value: unknown): value is DocxSectionBreakLike {
  return isRecord(value) && value.type === "sectionBreak";
}

function isInlineImage(value: unknown): value is DocxImageRunLike {
  return (
    isRecord(value) &&
    value.type === "image" &&
    value.anchor !== true &&
    typeof value.widthPt === "number" &&
    typeof value.heightPt === "number"
  );
}
