import { buildZip, type ZipFileSpec } from "./zip-builder.js";

/*
 * A DOCX builder for tests: a valid minimal package whose main part holds
 * the body XML the caller gives, plus optional header, footer and footnote
 * parts, so the display pre-pass can be exercised on every construct it
 * must handle.
 */

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const WP =
  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

export interface DocxOptions {
  /** Children of `w:body`, the closing `w:sectPr` included when wanted. */
  readonly body: string;
  /** Extra attributes on `w:document`, for example a `w14` declaration. */
  readonly rootAttributes?: string;
  readonly header?: string;
  readonly footer?: string;
  readonly footnotes?: string;
  /** Children of `w:styles`; a minimal docDefaults block when absent. */
  readonly styles?: string;
  /** Major and minor Latin faces of a theme part, when the document should have one. */
  readonly theme?: { readonly major: string; readonly minor: string };
  /** Media parts referenced from the body. */
  readonly media?: readonly {
    readonly name: string;
    readonly data: Uint8Array;
  }[];
}

export function sectPr(
  options: {
    readonly width?: number;
    readonly height?: number;
    readonly margin?: number;
  } = {},
): string {
  const margin = options.margin ?? 1440;
  return `<w:sectPr><w:pgSz w:w="${options.width ?? 12240}" w:h="${options.height ?? 15840}"/><w:pgMar w:top="${margin}" w:right="${margin}" w:bottom="${margin}" w:left="${margin}" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`;
}

export function paragraph(text: string, attributes = ""): string {
  return `<w:p${attributes ? ` ${attributes}` : ""}><w:r><w:t>${text}</w:t></w:r></w:p>`;
}

/** An inline picture of `cx` × `cy` EMU embedding `rId`. */
export function inlinePicture(
  cx: number,
  cy: number,
  rId = "rId9",
  anchored = false,
): string {
  const wrapper = anchored ? "wp:anchor" : "wp:inline";
  const position = anchored
    ? '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
    : "";
  return `<w:r><w:drawing><${wrapper} xmlns:a="${A}" xmlns:pic="${PIC}"${
    anchored
      ? ' distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"'
      : ""
  }>${position}<wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="${PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="image.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></${wrapper}></w:drawing></w:r>`;
}

function relationships(
  items: readonly { id: string; type: string; target: string }[],
): string {
  return `${XML}<Relationships xmlns="${REL}">${items
    .map(
      (item) =>
        `<Relationship Id="${item.id}" Type="${item.type}" Target="${item.target}"/>`,
    )
    .join("")}</Relationships>`;
}

export function buildDocx(options: DocxOptions): Uint8Array {
  const files: ZipFileSpec[] = [];
  const overrides: string[] = [
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
  ];
  const rels: { id: string; type: string; target: string }[] = [
    { id: "rId1", type: `${R}/styles`, target: "styles.xml" },
  ];
  files.push({
    name: "word/styles.xml",
    data: `${XML}<w:styles xmlns:w="${W}">${
      options.styles ??
      '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults>'
    }</w:styles>`,
    method: 8,
  });
  overrides.push(
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
  );
  if (options.theme) {
    files.push({
      name: "word/theme/theme1.xml",
      data: `${XML}<a:theme xmlns:a="${A}" name="Office"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:accent1><a:srgbClr val="4472C4"/></a:accent1></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="${options.theme.major}"/></a:majorFont><a:minorFont><a:latin typeface="${options.theme.minor}"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
      method: 8,
    });
    overrides.push(
      '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    );
    rels.push({
      id: `rId${rels.length + 1}`,
      type: `${R}/theme`,
      target: "theme/theme1.xml",
    });
  }
  const story = (
    name: string,
    kind: "header" | "footer" | "footnotes",
    content: string | undefined,
  ): void => {
    if (content === undefined) return;
    const tag =
      kind === "header" ? "w:hdr" : kind === "footer" ? "w:ftr" : "w:footnotes";
    files.push({
      name: `word/${name}`,
      data: `${XML}<${tag} xmlns:w="${W}" xmlns:r="${R}">${content}</${tag}>`,
      method: 8,
    });
    overrides.push(
      `<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`,
    );
    rels.push({
      id: `rId${rels.length + 1}`,
      type: `${R}/${kind}`,
      target: name,
    });
  };
  story("header1.xml", "header", options.header);
  story("footer1.xml", "footer", options.footer);
  story("footnotes.xml", "footnotes", options.footnotes);
  const defaults = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
  ];
  for (const media of options.media ?? []) {
    files.push({ name: media.name, data: media.data, method: 0 });
    const extension = media.name.slice(media.name.lastIndexOf(".") + 1);
    const type = extension === "png" ? "image/png" : "image/jpeg";
    if (!defaults.some((item) => item.includes(`Extension="${extension}"`)))
      defaults.push(
        `<Default Extension="${extension}" ContentType="${type}"/>`,
      );
    rels.push({
      id: "rId9",
      type: `${R}/image`,
      target: media.name.replace(/^word\//, ""),
    });
  }
  files.push({
    name: "word/document.xml",
    data: `${XML}<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="${WP}"${
      options.rootAttributes ? ` ${options.rootAttributes}` : ""
    }><w:body>${options.body}</w:body></w:document>`,
    method: 8,
  });
  files.push({
    name: "word/_rels/document.xml.rels",
    data: relationships(rels),
    method: 8,
  });
  return buildZip([
    {
      name: "[Content_Types].xml",
      data: `${XML}<Types xmlns="${CT}">${defaults.join("")}${overrides.join("")}</Types>`,
      method: 8,
    },
    {
      name: "_rels/.rels",
      data: relationships([
        {
          id: "rId1",
          type: `${R}/officeDocument`,
          target: "word/document.xml",
        },
      ]),
      method: 8,
    },
    ...files,
  ]);
}
