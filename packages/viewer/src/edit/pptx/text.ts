import type { XmlElement, XmlPart } from "../ooxml/xml.js";

/*
 * The text model of a text body: paragraphs joined by "\n", line breaks
 * (a:br) as "\v" — the character PowerPoint itself uses — and fields as
 * their cached text. Offsets are UTF-16 code units of the joined text.
 */

export const PARAGRAPH_BREAK = "\n";
export const LINE_BREAK = "\u000b";

export type RunKind = "r" | "br" | "fld";

export interface RunInfo {
  readonly node: XmlElement;
  readonly kind: RunKind;
  /** Half-open offsets in the body's text. */
  readonly start: number;
  readonly end: number;
  readonly text: string;
  /** `a:rPr` of the run or field, when present. */
  readonly rPr?: XmlElement;
}

export interface ParagraphInfo {
  readonly node: XmlElement;
  /** Half-open offsets in the body's text, the trailing "\n" excluded. */
  readonly start: number;
  readonly end: number;
  readonly runs: readonly RunInfo[];
  readonly pPr?: XmlElement;
  /** 0-based outline level (`a:pPr/@lvl`). */
  readonly level: number;
  readonly endParaRPr?: XmlElement;
}

export interface TextModel {
  readonly text: string;
  readonly paragraphs: readonly ParagraphInfo[];
}

/** Reads the text of a `p:txBody` or `a:txBody`. */
export function readTextModel(part: XmlPart, txBody: XmlElement): TextModel {
  const paragraphs: ParagraphInfo[] = [];
  let text = "";
  const paragraphNodes = txBody.children.filter((child) => child.local === "p");
  paragraphNodes.forEach((paragraph, index) => {
    if (index > 0) text += PARAGRAPH_BREAK;
    const start = text.length;
    const runs: RunInfo[] = [];
    let pPr: XmlElement | undefined;
    let endParaRPr: XmlElement | undefined;
    for (const child of paragraph.children) {
      switch (child.local) {
        case "pPr":
          pPr = child;
          break;
        case "endParaRPr":
          endParaRPr = child;
          break;
        case "r":
        case "fld": {
          const t = child.children.find((node) => node.local === "t");
          const value = t ? part.textOf(t) : "";
          const rPr = child.children.find((node) => node.local === "rPr");
          runs.push({
            node: child,
            kind: child.local,
            start: text.length,
            end: text.length + value.length,
            text: value,
            ...(rPr ? { rPr } : {}),
          });
          text += value;
          break;
        }
        case "br": {
          const rPr = child.children.find((node) => node.local === "rPr");
          runs.push({
            node: child,
            kind: "br",
            start: text.length,
            end: text.length + 1,
            text: LINE_BREAK,
            ...(rPr ? { rPr } : {}),
          });
          text += LINE_BREAK;
          break;
        }
        default:
          break;
      }
    }
    const level = pPr ? Number(part.attribute(pPr, "lvl") ?? "0") || 0 : 0;
    paragraphs.push({
      node: paragraph,
      start,
      end: text.length,
      runs,
      ...(pPr ? { pPr } : {}),
      level,
      ...(endParaRPr ? { endParaRPr } : {}),
    });
  });
  return { text, paragraphs };
}

/** The first run that draws text, with its paragraph; undefined for an empty body. */
export function firstTextRun(
  model: TextModel,
): { readonly run: RunInfo; readonly paragraph: ParagraphInfo } | undefined {
  for (const paragraph of model.paragraphs)
    for (const run of paragraph.runs)
      if (run.kind !== "br" && run.text.length > 0) return { run, paragraph };
  return undefined;
}

/** The cells of an `a:tbl`, row by row, as text models. */
export function readTableCells(
  part: XmlPart,
  tbl: XmlElement,
): readonly (readonly { node: XmlElement; model: TextModel }[])[] {
  const rows: { node: XmlElement; model: TextModel }[][] = [];
  for (const tr of tbl.children.filter((child) => child.local === "tr")) {
    const cells: { node: XmlElement; model: TextModel }[] = [];
    for (const tc of tr.children.filter((child) => child.local === "tc")) {
      const txBody = tc.children.find((child) => child.local === "txBody");
      cells.push({
        node: tc,
        model: txBody
          ? readTextModel(part, txBody)
          : { text: "", paragraphs: [] },
      });
    }
    rows.push(cells);
  }
  return rows;
}
