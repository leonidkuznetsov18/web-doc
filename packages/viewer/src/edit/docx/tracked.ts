import { scanXml, type XmlElement, type XmlPart } from "../ooxml/xml.js";
import { W_NS } from "./ids.js";
import type { ParagraphRecord } from "./model.js";
import type { DocxOperationContext, Issue } from "./operations.js";
import { LINE_BREAK, TAB, type RunItem } from "./text.js";
import { contentOf, rPrOf, unitsOf, type Unit } from "./text-ops.js";
import type { DocxRevision, DocxTextStyleChange } from "./types.js";
import {
  changedRunProperties,
  escapeAttributeValue,
  innerPropertiesXml,
  mergedProperties,
  paragraphXml,
  PPR_ORDER,
  runContentXml,
  runXml,
  sliceOf,
} from "./write.js";

/*
 * Tracked changes: the same paragraph rebuild as a direct edit, but what a
 * direct edit drops is kept inside `w:del` (text as `w:delText`) and what
 * it adds goes inside `w:ins`, each with an id, the author and the date of
 * the batch, so Word and Pages show the edit as a suggestion to accept or
 * reject. A paragraph rebuilt this way holds revisions and is read-only
 * for the next edit (decision 6 of the ai-edit module). Hyperlinks,
 * content controls and fields have no tracked form here: a change that
 * touches one is refused.
 */

/** The attributes of one revision element, with a fresh `w:id`. */
export function revisionAttributes(context: DocxOperationContext): string {
  const tracked = context.tracked!;
  return ` w:id="${context.nextRevisionId()}" w:author="${escapeAttributeValue(tracked.author)}"${
    tracked.date === undefined
      ? ""
      : ` w:date="${escapeAttributeValue(tracked.date)}"`
  }`;
}

/** A refusal for an operation, or a part of one, that has no tracked form. */
export function unsupportedTracked(
  issue: Issue,
  path: string,
  what: string,
): void {
  issue(
    path,
    "unsupported-change-mode",
    `${what} has no tracked form; apply it directly and review with checkpoints`,
  );
}

/** The child of `container` an item sits under. */
function topOf(node: XmlElement, container: XmlElement): XmlElement {
  let current = node;
  while (current.parent && current.parent !== container)
    current = current.parent;
  return current;
}

/**
 * Whether a tracked change of `[start, end)` would touch a hyperlink, a
 * content control, a field or a run already under revision, which the
 * tracked rebuild cannot express; reports it as an issue at `path`.
 */
export function trackedRangeProblem(
  record: ParagraphRecord,
  start: number,
  end: number,
  issue: Issue,
  path: string,
): boolean {
  for (const item of record.text.items) {
    const overlaps =
      item.start === item.end
        ? start < item.start && item.start < end
        : item.start < end && start < item.end;
    if (!overlaps) continue;
    const top = topOf(item.run, record.node);
    // An equation or an alternate-content block sits in the paragraph
    // without a run; a revision cannot wrap it either.
    const plainRun = item.run.local === "r" && item.run.namespace === W_NS;
    if (
      !plainRun ||
      item.fieldRuns ||
      item.kind === "field" ||
      top !== item.run ||
      item.wrapper !== undefined
    ) {
      unsupportedTracked(
        issue,
        path,
        "A change touching a hyperlink, content control, field or equation",
      );
      return true;
    }
  }
  return false;
}

/** Run content with `w:t` as `w:delText`, for a deleted run. */
function deletedContent(content: string): string {
  return content
    .replaceAll("<w:t>", "<w:delText>")
    .replaceAll("<w:t ", "<w:delText ")
    .replaceAll("</w:t>", "</w:delText>")
    .replaceAll("<w:instrText", "<w:delInstrText")
    .replaceAll("</w:instrText>", "</w:delInstrText>");
}

/** A deleted run: its properties and its content, text as `w:delText`. */
export function deletedRunXml(rPr: string, content: string): string {
  return content.length === 0
    ? ""
    : `<w:r>${rPr}${deletedContent(content)}</w:r>`;
}

/** The whole content of a run (every child but `w:rPr`) as bytes. */
function runContentOf(part: XmlPart, run: XmlElement): string {
  return run.children
    .filter((child) => !(child.local === "rPr" && child.namespace === W_NS))
    .map((child) => sliceOf(part, child))
    .join("");
}

/** `w:rPr` bytes of a run element. */
function runPropertiesOf(part: XmlPart, run: XmlElement): string {
  const rPr = run.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
  return sliceOf(part, rPr);
}

/**
 * The deleted part of a cut run: like `contentOf`, but a zero-width item
 * (a note reference, an anchored drawing) belongs to the deletion only
 * strictly inside `(from, to)`, so the kept head or tail keeps it instead
 * and nothing is written twice.
 */
function deletedContentOf(
  part: XmlPart,
  items: readonly RunItem[],
  from: number,
  to: number,
): string {
  let out = "";
  for (const item of items) {
    if (item.start === item.end) {
      if (from < item.start && item.start < to)
        out += sliceOf(part, item.child);
      continue;
    }
    const begin = Math.max(item.start, from);
    const stop = Math.min(item.end, to);
    if (stop <= begin) continue;
    if (item.kind === "text" && item.child.local === "t")
      out += runContentXml(
        item.text.slice(begin - item.start, stop - item.start),
      );
    else if (begin === item.start && stop === item.end)
      out += sliceOf(part, item.child);
  }
  return out;
}

export interface TrackedSplit {
  /** Everything before the change, bytes kept. */
  readonly before: string;
  /** The removed runs inside one `w:del`, or "". */
  readonly deleted: string;
  /** The `w:rPr` bytes new text takes. */
  readonly rPr: string;
  /** Everything after the change, bytes kept. */
  readonly after: string;
}

/**
 * The paragraph's content around `[start, end)`, unit by unit as the direct
 * edit walks it: the runs the range covers re-serialized as deleted, a cut
 * run split into a kept head, a deleted middle and a kept tail, everything
 * else as its bytes. New text takes the style of the first replaced run,
 * else of the run before a caret, else of the run after it. The caller
 * places the insertion between `deleted` and `after`. A wrapper or a field
 * inside the range never reaches here: validation refuses it.
 */
export function trackedSplit(
  context: DocxOperationContext,
  record: ParagraphRecord,
  start: number,
  end: number,
  fallbackRPr: string,
): TrackedSplit {
  const part = context.model.document;
  let before = "";
  let after = "";
  const deletedRuns: string[] = [];
  let rPr: string | undefined;
  let lastRPr: string | undefined;
  /** The change has been passed: zero-width children now follow it. */
  let passed = false;
  const raw = (unit: Unit): string =>
    unit.nodes.map((node) => sliceOf(part, node)).join("");
  for (const unit of unitsOf(record.node, record.text.items)) {
    if (unit.kind === "zero") {
      if (passed) after += raw(unit);
      else before += raw(unit);
      continue;
    }
    if (unit.end <= start) {
      before += raw(unit);
      lastRPr = rPrOf(part, unit.items.at(-1));
      // A caret right after this unit: the new text follows its style.
      if (start === end && unit.end === start) {
        rPr ??= lastRPr;
        passed = true;
      }
      continue;
    }
    if (unit.start >= end) {
      rPr ??= lastRPr ?? rPrOf(part, unit.items[0]);
      passed = true;
      after += raw(unit);
      continue;
    }
    if (unit.kind !== "run")
      throw new Error("A tracked change cannot cut a field or a wrapper");
    const run = unit.nodes[0]!;
    const style = runPropertiesOf(part, run);
    if (unit.start >= start && unit.end <= end) {
      // Covered whole: the run goes into the deletion as it is.
      deletedRuns.push(deletedRunXml(style, runContentOf(part, run)));
    } else {
      const head = contentOf(part, unit.items, unit.start, start, true);
      const middle = deletedContentOf(
        part,
        unit.items,
        Math.max(unit.start, start),
        Math.min(unit.end, end),
      );
      const tail = contentOf(part, unit.items, end, unit.end, start < end);
      if (head) before += runXml(style, head);
      if (middle) deletedRuns.push(deletedRunXml(style, middle));
      if (tail) after += runXml(style, tail);
    }
    rPr ??= style;
    if (unit.end >= end) passed = true;
  }
  const deleted =
    deletedRuns.length > 0
      ? `<w:del${revisionAttributes(context)}>${deletedRuns.join("")}</w:del>`
      : "";
  return { before, deleted, rPr: rPr ?? lastRPr ?? fallbackRPr, after };
}

/** `text` as an inserted run under revision; nothing for empty text. */
export function insertedRunXml(
  context: DocxOperationContext,
  rPr: string,
  text: string,
): string {
  const content = runContentXml(text);
  return content.length === 0
    ? ""
    : `<w:ins${revisionAttributes(context)}>${runXml(rPr, content)}</w:ins>`;
}

/** `w:pPr` bytes with the paragraph mark under revision (`w:ins` or `w:del` first in its `w:rPr`), section properties left out. */
export function markedParagraphProperties(
  context: DocxOperationContext,
  pPr: XmlElement | undefined,
  mark: "ins" | "del",
  insertion?: {
    readonly style: DocxTextStyleChange;
    readonly source: XmlElement | undefined;
  },
): string {
  const part = context.model.document;
  const rPr = pPr?.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
  const changed = insertion
    ? scanXml(
        "inserted-paragraph-mark",
        `<w:root xmlns:w="${W_NS}">${changedRunProperties(part, insertion.source, insertion.style, context.model.styles)}</w:root>`,
      )
    : undefined;
  const inner = changed
    ? innerPropertiesXml(changed, changed.root.children[0], ["ins", "del"])
    : innerPropertiesXml(part, rPr, ["ins", "del"]);
  const marked = `<w:rPr><w:${mark}${revisionAttributes(context)}/>${inner}</w:rPr>`;
  return mergedProperties(
    part,
    pPr,
    "w:pPr",
    PPR_ORDER,
    new Map([
      ["rPr", marked],
      ["sectPr", null],
    ]),
  );
}

/**
 * A whole paragraph marked deleted: its mark under `w:del`, every run
 * inside `w:del` with `w:delText`, other children kept. Word removes the
 * paragraph when the deletion is accepted.
 */
export function deletedParagraphXml(
  context: DocxOperationContext,
  record: ParagraphRecord,
): string {
  const part = context.model.document;
  const pPr = markedParagraphProperties(context, record.pPr, "del");
  let content = "";
  let open: string[] = [];
  const flush = (): void => {
    if (open.length > 0)
      content += `<w:del${revisionAttributes(context)}>${open.join("")}</w:del>`;
    open = [];
  };
  for (const child of record.node.children) {
    if (child.local === "pPr" && child.namespace === W_NS) continue;
    if (child.local === "r" && child.namespace === W_NS) {
      const run = deletedRunXml(
        runPropertiesOf(part, child),
        runContentOf(part, child),
      );
      if (run) open.push(run);
      continue;
    }
    flush();
    content += sliceOf(part, child);
  }
  flush();
  return paragraphXml(record.node, record.id, pPr, content);
}

const RUN_REVISIONS = new Set(["ins", "del", "moveFrom", "moveTo"]);

/**
 * The revisions a paragraph holds, in document order: run-level
 * insertions, deletions and moves with their text, the paragraph mark's
 * own revision, and changed run or paragraph properties.
 */
export function revisionsOf(
  part: XmlPart,
  paragraph: XmlElement,
): DocxRevision[] {
  const out: { start: number; revision: DocxRevision }[] = [];
  const attributes = (
    node: XmlElement,
  ): Pick<DocxRevision, "id" | "author" | "date"> => {
    const author = part.attribute(node, "w:author");
    const date = part.attribute(node, "w:date");
    const raw = part.attribute(node, "w:id");
    return {
      id: raw !== undefined && /^\d+$/.test(raw) ? Number(raw) : -1,
      ...(author === undefined ? {} : { author }),
      ...(date === undefined ? {} : { date }),
    };
  };
  const visit = (node: XmlElement, inRPr: boolean, inPPr: boolean): void => {
    for (const child of node.children) {
      if (child.namespace !== W_NS) {
        visit(child, inRPr, inPPr);
        continue;
      }
      if (child.local === "pPr") {
        visit(child, false, true);
        continue;
      }
      if (child.local === "rPr") {
        visit(child, true, inPPr);
        continue;
      }
      if (inRPr && RUN_REVISIONS.has(child.local)) {
        out.push({
          start: child.start,
          revision: {
            kind: child.local as DocxRevision["kind"],
            scope: inPPr ? "mark" : "runs",
            ...attributes(child),
            ...(inPPr ? { text: "\n" } : {}),
          },
        });
        continue;
      }
      if (inRPr && child.local === "rPrChange") {
        out.push({
          start: child.start,
          revision: {
            kind: "rPrChange",
            scope: inPPr ? "mark" : "runs",
            ...attributes(child),
            ...(inPPr || !node.parent
              ? {}
              : { text: textOf(part, node.parent) }),
          },
        });
        continue;
      }
      if (inPPr && child.local === "pPrChange") {
        out.push({
          start: child.start,
          revision: {
            kind: "pPrChange",
            scope: "paragraph",
            ...attributes(child),
          },
        });
        continue;
      }
      if (!inRPr && !inPPr && RUN_REVISIONS.has(child.local)) {
        out.push({
          start: child.start,
          revision: {
            kind: child.local as DocxRevision["kind"],
            scope: "runs",
            ...attributes(child),
            text: textOf(part, child),
          },
        });
        visit(child, false, false);
        continue;
      }
      visit(child, inRPr, inPPr);
    }
  };
  visit(paragraph, false, false);
  return out
    .sort((a, b) => a.start - b.start)
    .map((entry) => Object.freeze(entry.revision));
}

/** The text of every run under `node`: `w:t` and `w:delText`, tabs and breaks. */
function textOf(part: XmlPart, node: XmlElement): string {
  let text = "";
  const visit = (current: XmlElement): void => {
    for (const child of current.children) {
      if (child.namespace === W_NS) {
        if (child.local === "t" || child.local === "delText") {
          text += part.textOf(child);
          continue;
        }
        if (child.local === "tab") {
          text += TAB;
          continue;
        }
        if (child.local === "br" || child.local === "cr") {
          text += LINE_BREAK;
          continue;
        }
        if (child.local === "rPr" || child.local === "pPr") continue;
      }
      visit(child);
    }
  };
  visit(node);
  return text;
}
