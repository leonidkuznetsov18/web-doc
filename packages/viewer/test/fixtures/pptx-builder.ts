import { buildZip, type ZipFileSpec } from "./zip-builder.js";

/*
 * A PPTX builder for tests: a valid minimal deck (presentation, one master,
 * two layouts, a theme) whose slides hold the shape XML the caller gives,
 * so every element kind, inheritance case and group transform the engine
 * must handle can be produced on purpose. The decks open in the renderer.
 */

const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_TYPE = `${R}/`;
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export const SLIDE_WIDTH = 9144000;
export const SLIDE_HEIGHT = 6858000;

export interface DeckRelationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
}

export interface DeckSlide {
  /** Shape XML, in order, inside `p:spTree`. */
  readonly shapes: readonly string[];
  /** 1-based layout number; layout 1 is "Title and Content", layout 2 "Title Slide". */
  readonly layout?: number;
  readonly hidden?: boolean;
  /** Relationships beyond the layout (media, charts); ids start at rId2. */
  readonly relationships?: readonly DeckRelationship[];
  /** Part name without the leading slash, for decks with gaps or odd order. */
  readonly partName?: string;
}

export interface DeckOptions {
  readonly slides: readonly DeckSlide[];
  readonly size?: { readonly cx: number; readonly cy: number };
  /** Adds `ppt/vbaProject.bin` and marks the main part as a macro-enabled presentation. */
  readonly macro?: boolean;
  /** Binary parts, for example `ppt/media/image1.png`. */
  readonly parts?: readonly {
    readonly name: string;
    readonly data: Uint8Array | string;
    readonly contentType?: string;
  }[];
  /** Whether the presentation carries a `p:defaultTextStyle`; default true. */
  readonly defaultTextStyle?: boolean;
}

export interface ShapeSpec {
  readonly id: number;
  readonly name?: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** Degrees clockwise. */
  readonly rotation?: number;
  readonly flipH?: boolean;
  readonly flipV?: boolean;
  readonly hidden?: boolean;
}

export interface TextShapeSpec extends ShapeSpec {
  /** Paragraphs; each a list of runs, a run being text or `{ text, rPr }`. */
  readonly paragraphs?: readonly (readonly (
    | string
    | { readonly text: string; readonly rPr?: string }
    | { readonly br: true }
  )[])[];
  /** Attributes of the first `a:pPr` of every paragraph, for example 'algn="ctr"'. */
  readonly pPr?: string;
  readonly fill?: string;
  readonly line?: string;
  /** Extra `a:bodyPr` content, for example "<a:normAutofit fontScale=\"90000\"/>". */
  readonly bodyPr?: string;
  readonly lstStyle?: string;
  /** Omits `a:xfrm`, so the frame is inherited through the placeholder. */
  readonly inherit?: boolean;
  readonly placeholder?: { readonly type?: string; readonly idx?: number };
  readonly txBox?: boolean;
}

function xfrmXml(
  spec: ShapeSpec,
  child?: { x: number; y: number; cx: number; cy: number },
): string {
  const attributes = [
    spec.rotation ? ` rot="${Math.round(spec.rotation * 60000)}"` : "",
    spec.flipH ? ' flipH="1"' : "",
    spec.flipV ? ' flipV="1"' : "",
  ].join("");
  return `<a:xfrm${attributes}><a:off x="${spec.x}" y="${spec.y}"/><a:ext cx="${spec.cx}" cy="${spec.cy}"/>${
    child
      ? `<a:chOff x="${child.x}" y="${child.y}"/><a:chExt cx="${child.cx}" cy="${child.cy}"/>`
      : ""
  }</a:xfrm>`;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function cNvPr(spec: ShapeSpec, fallback: string): string {
  return `<p:cNvPr id="${spec.id}" name="${escapeXml(spec.name ?? `${fallback} ${spec.id}`)}"${
    spec.hidden ? ' hidden="1"' : ""
  }/>`;
}

function paragraphsXml(spec: TextShapeSpec): string {
  const paragraphs = spec.paragraphs ?? [["Hello"]];
  return paragraphs
    .map((runs) => {
      const body = runs
        .map((run) => {
          if (typeof run === "string")
            return `<a:r><a:rPr lang="en-US"/><a:t>${escapeXml(run)}</a:t></a:r>`;
          if ("br" in run) return "<a:br/>";
          return `<a:r><a:rPr lang="en-US"${run.rPr ? ` ${run.rPr}` : ""}/><a:t>${escapeXml(run.text)}</a:t></a:r>`;
        })
        .join("");
      return `<a:p>${spec.pPr ? `<a:pPr ${spec.pPr}/>` : ""}${body}<a:endParaRPr lang="en-US"/></a:p>`;
    })
    .join("");
}

/** A `p:sp` with a text body. */
export function textShape(spec: TextShapeSpec): string {
  const ph = spec.placeholder
    ? `<p:ph${spec.placeholder.type ? ` type="${spec.placeholder.type}"` : ""}${
        spec.placeholder.idx !== undefined
          ? ` idx="${spec.placeholder.idx}"`
          : ""
      }/>`
    : "";
  const spPr = [
    spec.inherit ? "" : xfrmXml(spec),
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>',
    spec.fill ?? "",
    spec.line ?? "",
  ].join("");
  return `<p:sp><p:nvSpPr>${cNvPr(spec, spec.txBox ? "TextBox" : "Shape")}<p:cNvSpPr${
    spec.txBox ? ' txBox="1"' : ""
  }/><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${spPr}</p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0">${
    spec.bodyPr ?? ""
  }</a:bodyPr>${spec.lstStyle ?? "<a:lstStyle/>"}${paragraphsXml(spec)}</p:txBody></p:sp>`;
}

/** A `p:pic` whose blip embeds `rId`. */
export function picture(spec: ShapeSpec & { readonly rId: string }): string {
  return `<p:pic><p:nvPicPr>${cNvPr(spec, "Picture")}<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${spec.rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrmXml(
    spec,
  )}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
}

/** A `p:graphicFrame` holding an `a:tbl`. */
export function table(
  spec: ShapeSpec & { readonly rows: readonly (readonly string[])[] },
): string {
  const columns = spec.rows[0]?.length ?? 1;
  const columnWidth = Math.floor(spec.cx / columns);
  const rowHeight = Math.floor(spec.cy / spec.rows.length);
  const grid = Array.from(
    { length: columns },
    () => `<a:gridCol w="${columnWidth}"/>`,
  ).join("");
  const rows = spec.rows
    .map(
      (row) =>
        `<a:tr h="${rowHeight}">${row
          .map(
            (cell) =>
              `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>${escapeXml(
                cell,
              )}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`,
          )
          .join("")}</a:tr>`,
    )
    .join("");
  return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(spec, "Table")}<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${spec.x}" y="${spec.y}"/><a:ext cx="${spec.cx}" cy="${spec.cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

/** A `p:graphicFrame` whose graphic the engine does not edit (a chart reference). */
export function graphicFrame(
  spec: ShapeSpec & { readonly uri: string; readonly inner: string },
): string {
  return `<p:graphicFrame><p:nvGraphicFramePr>${cNvPr(spec, "Chart")}<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${spec.x}" y="${spec.y}"/><a:ext cx="${spec.cx}" cy="${spec.cy}"/></p:xfrm><a:graphic><a:graphicData uri="${spec.uri}">${spec.inner}</a:graphicData></a:graphic></p:graphicFrame>`;
}

/** A `p:grpSp` with a child space and children. */
export function group(
  spec: ShapeSpec & {
    readonly child: { x: number; y: number; cx: number; cy: number };
    readonly children: readonly string[];
  },
): string {
  return `<p:grpSp><p:nvGrpSpPr>${cNvPr(spec, "Group")}<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>${xfrmXml(
    spec,
    spec.child,
  )}</p:grpSpPr>${spec.children.join("")}</p:grpSp>`;
}

/** A `p:cxnSp` line. */
export function connector(
  spec: ShapeSpec & { readonly line?: string },
): string {
  return `<p:cxnSp><p:nvCxnSpPr>${cNvPr(spec, "Connector")}<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrmXml(
    spec,
  )}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${spec.line ?? '<a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>'}</p:spPr></p:cxnSp>`;
}

/** `mc:AlternateContent` with a choice the renderer does not know and a fallback. */
export function alternateContent(choice: string, fallback: string): string {
  return `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`;
}

const THEME = `${XML}<a:theme xmlns:a="${A}" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;

function placeholderShape(
  id: number,
  name: string,
  ph: string,
  xfrm: string,
  lstStyle = "<a:lstStyle/>",
  text = "Click to edit",
): string {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${ph}</p:nvPr></p:nvSpPr><p:spPr>${xfrm}</p:spPr><p:txBody><a:bodyPr/>${lstStyle}<a:p><a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
}

const TREE_HEAD =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';

/** Frames of the master's placeholders, in EMU, for tests that check inheritance. */
export const MASTER_TITLE = { x: 457200, y: 274638, cx: 8229600, cy: 1143000 };
export const MASTER_BODY = { x: 457200, y: 1600200, cx: 8229600, cy: 4525963 };
/** Frames of layout 2 ("Title Slide"), which override the master's. */
export const LAYOUT2_TITLE = {
  x: 685800,
  y: 2130425,
  cx: 7772400,
  cy: 1470025,
};
export const LAYOUT2_SUBTITLE = {
  x: 1371600,
  y: 3886200,
  cx: 6400800,
  cy: 1752600,
};

function frame(box: { x: number; y: number; cx: number; cy: number }): string {
  return `<a:xfrm><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>`;
}

const MASTER = `${XML}<p:sldMaster xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${TREE_HEAD}${placeholderShape(
  2,
  "Title Placeholder 1",
  '<p:ph type="title"/>',
  frame(MASTER_TITLE),
)}${placeholderShape(
  3,
  "Text Placeholder 2",
  '<p:ph type="body" idx="1"/>',
  frame(MASTER_BODY),
)}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/><p:sldLayoutId id="2147483650" r:id="rId2"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr marL="342900" indent="-342900" algn="l"><a:buChar char="•"/><a:defRPr sz="3200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr><a:lvl2pPr marL="742950" indent="-285750" algn="l"><a:defRPr sz="2800"/></a:lvl2pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`;

const LAYOUT1 = `${XML}<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="obj" preserve="1"><p:cSld name="Title and Content"><p:spTree>${TREE_HEAD}${placeholderShape(
  2,
  "Title 1",
  '<p:ph type="title"/>',
  "",
)}${placeholderShape(
  3,
  "Content Placeholder 2",
  '<p:ph idx="1"/>',
  "",
)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

const LAYOUT2 = `${XML}<p:sldLayout xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" type="title" preserve="1"><p:cSld name="Title Slide"><p:spTree>${TREE_HEAD}${placeholderShape(
  2,
  "Title 1",
  '<p:ph type="ctrTitle"/>',
  frame(LAYOUT2_TITLE),
)}${placeholderShape(
  3,
  "Subtitle 2",
  '<p:ph type="subTitle" idx="1"/>',
  frame(LAYOUT2_SUBTITLE),
  '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>',
)}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;

function relationships(items: readonly DeckRelationship[]): string {
  return `${XML}<Relationships xmlns="${REL}">${items
    .map(
      (item) =>
        `<Relationship Id="${item.id}" Type="${item.type}" Target="${item.target}"/>`,
    )
    .join("")}</Relationships>`;
}

/** Builds a deck; every XML entry is deflated like PowerPoint writes it. */
export function buildDeck(options: DeckOptions): Uint8Array {
  const size = options.size ?? { cx: SLIDE_WIDTH, cy: SLIDE_HEIGHT };
  const files: ZipFileSpec[] = [];
  const overrides: string[] = [];
  const defaults = new Map<string, string>([
    ["rels", "application/vnd.openxmlformats-package.relationships+xml"],
    ["xml", "application/xml"],
    ["png", "image/png"],
    ["jpeg", "image/jpeg"],
  ]);
  const override = (part: string, type: string): void => {
    overrides.push(`<Override PartName="/${part}" ContentType="${type}"/>`);
  };
  const xml = (name: string, data: string): void => {
    files.push({ name, data, method: 8 });
  };

  const mainType = options.macro
    ? "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    : "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml";
  override("ppt/presentation.xml", mainType);
  override(
    "ppt/slideMasters/slideMaster1.xml",
    "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
  );
  for (const n of [1, 2])
    override(
      `ppt/slideLayouts/slideLayout${n}.xml`,
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    );
  override(
    "ppt/theme/theme1.xml",
    "application/vnd.openxmlformats-officedocument.theme+xml",
  );

  const slideParts = options.slides.map(
    (slide, index) => slide.partName ?? `ppt/slides/slide${index + 1}.xml`,
  );
  const presentationRels: DeckRelationship[] = [
    {
      id: "rId1",
      type: `${REL_TYPE}slideMaster`,
      target: "slideMasters/slideMaster1.xml",
    },
    { id: "rId2", type: `${REL_TYPE}theme`, target: "theme/theme1.xml" },
  ];
  const sldIds: string[] = [];
  options.slides.forEach((slide, index) => {
    const part = slideParts[index]!;
    const rId = `rId${index + 3}`;
    presentationRels.push({
      id: rId,
      type: `${REL_TYPE}slide`,
      target: part.replace(/^ppt\//, ""),
    });
    sldIds.push(`<p:sldId id="${256 + index}" r:id="${rId}"/>`);
    override(
      part,
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    xml(
      part,
      `${XML}<p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"${
        slide.hidden ? ' show="0"' : ""
      }><p:cSld><p:spTree>${TREE_HEAD}${slide.shapes.join(
        "",
      )}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    );
    const folder = part.slice(0, part.lastIndexOf("/"));
    const file = part.slice(part.lastIndexOf("/") + 1);
    xml(
      `${folder}/_rels/${file}.rels`,
      relationships([
        {
          id: "rId1",
          type: `${REL_TYPE}slideLayout`,
          target: `../slideLayouts/slideLayout${slide.layout ?? 1}.xml`,
        },
        ...(slide.relationships ?? []),
      ]),
    );
  });

  const defaultTextStyle =
    options.defaultTextStyle === false
      ? ""
      : '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l"><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/></a:defRPr></a:lvl1pPr></p:defaultTextStyle>';
  xml(
    "ppt/presentation.xml",
    `${XML}<p:presentation xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>${sldIds.join(
      "",
    )}</p:sldIdLst><p:sldSz cx="${size.cx}" cy="${size.cy}"/><p:notesSz cx="6858000" cy="9144000"/>${defaultTextStyle}</p:presentation>`,
  );
  xml("ppt/_rels/presentation.xml.rels", relationships(presentationRels));
  xml("ppt/slideMasters/slideMaster1.xml", MASTER);
  xml(
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    relationships([
      {
        id: "rId1",
        type: `${REL_TYPE}slideLayout`,
        target: "../slideLayouts/slideLayout1.xml",
      },
      {
        id: "rId2",
        type: `${REL_TYPE}slideLayout`,
        target: "../slideLayouts/slideLayout2.xml",
      },
      { id: "rId3", type: `${REL_TYPE}theme`, target: "../theme/theme1.xml" },
    ]),
  );
  xml("ppt/slideLayouts/slideLayout1.xml", LAYOUT1);
  xml("ppt/slideLayouts/slideLayout2.xml", LAYOUT2);
  for (const n of [1, 2])
    xml(
      `ppt/slideLayouts/_rels/slideLayout${n}.xml.rels`,
      relationships([
        {
          id: "rId1",
          type: `${REL_TYPE}slideMaster`,
          target: "../slideMasters/slideMaster1.xml",
        },
      ]),
    );
  xml("ppt/theme/theme1.xml", THEME);

  for (const part of options.parts ?? []) {
    const extension = part.name.slice(part.name.lastIndexOf(".") + 1);
    if (part.contentType) override(part.name, part.contentType);
    else if (!defaults.has(extension))
      defaults.set(extension, "application/octet-stream");
    files.push({ name: part.name, data: part.data, method: 0 });
  }
  if (options.macro) {
    override("ppt/vbaProject.bin", "application/vnd.ms-office.vbaProject");
    files.push({
      name: "ppt/vbaProject.bin",
      data: new Uint8Array([
        0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 7, 7, 7,
      ]),
      method: 0,
    });
  }

  const contentTypes = `${XML}<Types xmlns="${CT}">${[...defaults]
    .map(
      ([extension, type]) =>
        `<Default Extension="${extension}" ContentType="${type}"/>`,
    )
    .join("")}${overrides.join("")}</Types>`;
  return buildZip([
    { name: "[Content_Types].xml", data: contentTypes, method: 8 },
    {
      name: "_rels/.rels",
      data: relationships([
        {
          id: "rId1",
          type: `${REL_TYPE}officeDocument`,
          target: "ppt/presentation.xml",
        },
      ]),
      method: 8,
    },
    ...files,
  ]);
}

/** A deck of `count` slides, each with a title placeholder and a text box, that the renderer opens. */
export function syntheticDeck(count: number): Uint8Array {
  return buildDeck({
    slides: Array.from({ length: count }, (_, index) => ({
      shapes: [
        textShape({
          id: 2,
          name: `Title ${index + 1}`,
          inherit: true,
          placeholder: { type: "title" },
          x: 0,
          y: 0,
          cx: 0,
          cy: 0,
          paragraphs: [[`Slide ${index + 1}`]],
        }),
        textShape({
          id: 3,
          name: `Body ${index + 1}`,
          txBox: true,
          x: 914400,
          y: 1828800,
          cx: 6400800,
          cy: 914400,
          paragraphs: [[`Body text of slide ${index + 1}`]],
        }),
      ],
    })),
  });
}
