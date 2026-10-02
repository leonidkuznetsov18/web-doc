import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import { W_NS } from "./ids.js";

/*
 * The text model of a paragraph: `w:t` text, a tab, a line break (`\v`), a
 * page or column break (`\f`), an inline picture or an embedded object as
 * one object character, a field (`w:fldSimple`, or a complex field from
 * `fldChar begin` to `end`) as its cached result text, hyperlinks and
 * inline `w:sdt` as their runs, hidden runs as text, deleted text (`w:del`)
 * left out. Offsets are UTF-16 code units of the paragraph's text. Every
 * item remembers the XML it came from, so a later patch can keep the bytes
 * of what an edit does not touch.
 */

export const TAB = "\t";
export const LINE_BREAK = "\u000b";
export const PAGE_BREAK = "\f";
/** The object replacement character: an inline picture or an embedded object. */
export const OBJECT = "￼";

export type RunItemKind =
  | "text"
  | "tab"
  | "break"
  | "pageBreak"
  | "picture" // an inline w:drawing
  | "anchor" // an anchored w:drawing (no text)
  | "object" // w:object, w:pict, equations, alternate content
  | "field" // the cached result of a field
  | "note"; // a footnote or endnote reference (no text)

export interface RunItem {
  readonly kind: RunItemKind;
  /** Half-open offsets in the paragraph's text. */
  readonly start: number;
  readonly end: number;
  readonly text: string;
  /** The `w:r` holding the item, or the `w:fldSimple` of a simple field. */
  readonly run: XmlElement;
  /** The child of the run the item comes from (`w:t`, `w:tab`, `w:drawing`, …). */
  readonly child: XmlElement;
  /** `w:rPr` of the run, when present. */
  readonly rPr?: XmlElement;
  /** The hyperlink, inline sdt, smart tag or tracked-change wrapper the run sits in, when any. */
  readonly wrapper?: XmlElement;
  /** For a complex field: every run from `begin` to `end`, in order. */
  readonly fieldRuns?: readonly XmlElement[];
}

export interface ParagraphText {
  readonly text: string;
  readonly items: readonly RunItem[];
  /** The paragraph holds tracked insertions, deletions or moves. */
  readonly tracked: boolean;
}

/** Wrappers whose children are read as the paragraph's own runs. */
const TRANSPARENT = new Set([
  "hyperlink",
  "smartTag",
  "customXml",
  "dir",
  "bdo",
  "ins",
  "moveTo",
  "sdtContent",
  "fldSimple",
]);
/** Wrappers whose content is deleted text: left out, the paragraph is tracked. */
const DELETED = new Set(["del", "moveFrom"]);

interface Field {
  readonly begin: XmlElement;
  readonly runs: XmlElement[];
  /** Text of the result runs (after `separate`). */
  text: string;
  /** Per nesting level, whether the field is past its `separate`; the outermost first. */
  readonly inResult: boolean[];
}

class Reader {
  readonly items: RunItem[] = [];
  text = "";
  tracked = false;
  #field: Field | undefined;

  constructor(private readonly part: XmlPart) {}

  paragraph(paragraph: XmlElement): void {
    for (const child of paragraph.children) this.#child(child, undefined);
    if (this.#field) this.#closeField();
    // A paragraph mark inserted, deleted or moved under revision marks the
    // paragraph as well.
    const pPr = paragraph.children.find(
      (child) => child.local === "pPr" && child.namespace === W_NS,
    );
    if (pPr && paragraphMarkTracked(pPr)) this.#tracked();
  }

  #child(node: XmlElement, wrapper: XmlElement | undefined): void {
    if (node.namespace !== W_NS) {
      // DrawingML, math and markup compatibility live outside w:; a math
      // zone or an alternate-content block counts as one object.
      if (
        node.local === "oMath" ||
        node.local === "oMathPara" ||
        node.local === "AlternateContent"
      )
        this.#push("object", OBJECT, node, node, wrapper);
      return;
    }
    switch (node.local) {
      case "r":
        this.#run(node, wrapper);
        return;
      case "sdt": {
        const content = node.children.find(
          (child) => child.local === "sdtContent" && child.namespace === W_NS,
        );
        if (content)
          for (const child of content.children) this.#child(child, node);
        return;
      }
      case "ins":
      case "moveTo":
        this.#tracked();
        for (const child of node.children) this.#child(child, node);
        return;
      default:
        break;
    }
    if (DELETED.has(node.local)) {
      this.#tracked();
      return;
    }
    if (TRANSPARENT.has(node.local)) {
      if (node.local === "fldSimple") {
        this.#simpleField(node, wrapper);
        return;
      }
      for (const child of node.children) this.#child(child, node);
    }
  }

  #tracked(): void {
    this.tracked = true;
  }

  /** A simple field: one item carrying the text of its result runs. */
  #simpleField(node: XmlElement, wrapper: XmlElement | undefined): void {
    const inner = new Reader(this.part);
    for (const child of node.children) inner.#child(child, node);
    this.tracked ||= inner.tracked;
    this.#push("field", inner.text, node, node, wrapper);
  }

  #run(run: XmlElement, wrapper: XmlElement | undefined): void {
    const rPr = run.children.find(
      (child) => child.local === "rPr" && child.namespace === W_NS,
    );
    if (rPr) {
      const marked = rPr.children.some(
        (child) =>
          child.namespace === W_NS &&
          (child.local === "ins" || child.local === "del"),
      );
      if (marked) this.#tracked();
    }
    if (this.#field) this.#field.runs.push(run);
    for (const child of run.children) {
      if (child.namespace !== W_NS) {
        if (child.local === "AlternateContent")
          this.#emit("object", OBJECT, run, child, rPr, wrapper);
        continue;
      }
      switch (child.local) {
        case "rPr":
        case "lastRenderedPageBreak":
        case "instrText":
        case "delInstrText":
        case "delText":
          break;
        case "fldChar":
          this.#fieldChar(child, run);
          break;
        case "t":
          this.#emit("text", this.part.textOf(child), run, child, rPr, wrapper);
          break;
        case "tab":
        case "ptab":
          this.#emit("tab", TAB, run, child, rPr, wrapper);
          break;
        case "br": {
          const type = this.part.attribute(child, "w:type");
          if (type === "page" || type === "column")
            this.#emit("pageBreak", PAGE_BREAK, run, child, rPr, wrapper);
          else this.#emit("break", LINE_BREAK, run, child, rPr, wrapper);
          break;
        }
        case "cr":
          this.#emit("break", LINE_BREAK, run, child, rPr, wrapper);
          break;
        case "sym": {
          const code = Number.parseInt(
            this.part.attribute(child, "w:char") ?? "",
            16,
          );
          this.#emit(
            "text",
            Number.isFinite(code) && code > 0 && code <= 0x10ffff
              ? String.fromCodePoint(code)
              : "",
            run,
            child,
            rPr,
            wrapper,
          );
          break;
        }
        case "noBreakHyphen":
          this.#emit("text", "‑", run, child, rPr, wrapper);
          break;
        case "softHyphen":
          this.#emit("text", "­", run, child, rPr, wrapper);
          break;
        case "drawing": {
          const inline = child.children.some((node) => node.local === "inline");
          if (inline) this.#emit("picture", OBJECT, run, child, rPr, wrapper);
          else this.#emit("anchor", "", run, child, rPr, wrapper);
          break;
        }
        case "object":
        case "pict":
          this.#emit("object", OBJECT, run, child, rPr, wrapper);
          break;
        case "footnoteReference":
        case "endnoteReference":
          this.#emit("note", "", run, child, rPr, wrapper);
          break;
        case "ruby": {
          const base = child.children.find((node) => node.local === "rubyBase");
          if (base)
            for (const node of base.children) this.#child(node, wrapper);
          break;
        }
        default:
          break;
      }
    }
  }

  #fieldChar(node: XmlElement, run: XmlElement): void {
    const type = this.part.attribute(node, "w:fldCharType");
    if (type === "begin") {
      if (this.#field) this.#field.inResult.push(false);
      else
        this.#field = { begin: run, runs: [run], text: "", inResult: [false] };
      return;
    }
    if (!this.#field) return;
    const levels = this.#field.inResult;
    if (type === "separate") {
      levels[levels.length - 1] = true;
      return;
    }
    if (type === "end") {
      if (levels.length > 1) levels.pop();
      else this.#closeField();
    }
  }

  #closeField(): void {
    const field = this.#field!;
    this.#field = undefined;
    const rPr = field.begin.children.find(
      (child) => child.local === "rPr" && child.namespace === W_NS,
    );
    this.#append({
      kind: "field",
      start: this.text.length,
      end: this.text.length + field.text.length,
      text: field.text,
      run: field.begin,
      child: field.begin,
      ...(rPr ? { rPr } : {}),
      fieldRuns: field.runs,
    });
  }

  /** Text inside an open field goes to its result; outside it is an item. */
  #emit(
    kind: RunItemKind,
    text: string,
    run: XmlElement,
    child: XmlElement,
    rPr: XmlElement | undefined,
    wrapper: XmlElement | undefined,
  ): void {
    const field = this.#field;
    if (field) {
      // What Word shows: the result of the outer field, which holds the
      // results of the fields nested in it and none of their instructions.
      if (field.inResult.every(Boolean)) field.text += text;
      return;
    }
    this.#push(kind, text, run, child, wrapper, rPr);
  }

  #push(
    kind: RunItemKind,
    text: string,
    run: XmlElement,
    child: XmlElement,
    wrapper: XmlElement | undefined,
    rPr?: XmlElement,
  ): void {
    this.#append({
      kind,
      start: this.text.length,
      end: this.text.length + text.length,
      text,
      run,
      child,
      ...(rPr ? { rPr } : {}),
      ...(wrapper ? { wrapper } : {}),
    });
  }

  #append(item: RunItem): void {
    this.items.push(item);
    this.text += item.text;
  }
}

const MARK_REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo"]);

/**
 * Whether the paragraph mark's `w:rPr` carries an insertion, deletion or
 * move. Changed properties alone (`w:rPrChange`, `w:pPrChange`) leave a
 * paragraph editable, as module 06 decided.
 */
function paragraphMarkTracked(pPr: XmlElement): boolean {
  for (const child of pPr.children) {
    if (child.namespace !== W_NS) continue;
    if (
      child.local === "rPr" &&
      child.children.some(
        (mark) => mark.namespace === W_NS && MARK_REVISIONS.has(mark.local),
      )
    )
      return true;
  }
  return false;
}

/** Reads the text of a `w:p`. */
export function readParagraphText(
  part: XmlPart,
  paragraph: XmlElement,
): ParagraphText {
  const reader = new Reader(part);
  reader.paragraph(paragraph);
  return { text: reader.text, items: reader.items, tracked: reader.tracked };
}

/** The first item that draws text, for the resolved style of a paragraph. */
export function firstTextItem(model: ParagraphText): RunItem | undefined {
  return model.items.find(
    (item) =>
      (item.kind === "text" || item.kind === "field") && item.text.length > 0,
  );
}
