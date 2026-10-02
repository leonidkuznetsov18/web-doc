import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import { W_NS } from "./ids.js";
import type { ParagraphRecord } from "./model.js";
import type { DocxOperationContext, Issue } from "./operations.js";
import { LINE_BREAK, TAB, type RunItem } from "./text.js";
import type { DocxRevision } from "./types.js";
import {
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
    if (
      item.fieldRuns ||
      item.kind === "field" ||
      top !== item.run ||
      item.wrapper !== undefined
    ) {
      unsupportedTracked(
        issue,
        path,
        "A change touching a hyperlink, content control or field",
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

function rPrOf(part: XmlPart, run: XmlElement): string {
  const rPr = run.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
  return sliceOf(part, rPr);
}

/** Runs' content in `[from, to)`, as `contentOf` of the direct edit does for one run. */
function contentOf(
  part: XmlPart,
  items: readonly RunItem[],
  from: number,
  to: number,
  includeStart: boolean,
): string {
  let out = "";
  for (const item of items) {
    if (item.start === item.end) {
      const after = includeStart ? item.start >= from : item.start > from;
      if (after && item.start <= to) out += sliceOf(part, item.child);
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
 * The paragraph's content around `[start, end)`: the runs the range covers
 * re-serialized as deleted, a cut run split into a kept head, a deleted
 * middle and a kept tail, everything else as its bytes. The caller places
 * the insertion between `deleted` and `after`.
 */
export function trackedSplit(
  context: DocxOperationContext,
  record: ParagraphRecord,
  start: number,
  end: number,
  fallbackRPr: string,
): TrackedSplit {
  const part = context.model.document;
  const container = record.node;
  const byRun = new Map<XmlElement, RunItem[]>();
  for (const item of record.text.items) {
    const list = byRun.get(item.run) ?? [];
    list.push(item);
    byRun.set(item.run, list);
  }
  let before = "";
  let after = "";
  const deletedRuns: string[] = [];
  let rPr: string | undefined;
  let lastRPr: string | undefined;
  let passed = false;
  for (const child of container.children) {
    if (child.local === "pPr" && child.namespace === W_NS) continue;
    const raw = sliceOf(part, child);
    const items = byRun.get(child);
    if (!items || items.length === 0) {
      // Bookmarks, proofing marks and other zero-width children stay where
      // they are: before the change until it, after it from then on.
      if (passed) after += raw;
      else before += raw;
      continue;
    }
    const unitStart = items[0]!.start;
    const unitEnd = items.at(-1)!.end;
    if (unitEnd <= start && !(unitStart === unitEnd && unitStart === start)) {
      before += raw;
      lastRPr = rPrOf(part, child);
      continue;
    }
    if (unitStart >= end && (unitStart > start || start < end)) {
      passed = true;
      rPr ??= lastRPr ?? rPrOf(part, child);
      after += raw;
      continue;
    }
    // The run overlaps the range: validation guaranteed it is a plain run.
    const style = rPrOf(part, child);
    const head = contentOf(part, items, unitStart, start, true);
    const middle = contentOf(
      part,
      items,
      Math.max(unitStart, start),
      Math.min(unitEnd, end),
      false,
    );
    const tail = contentOf(part, items, end, unitEnd, start < end);
    if (head) before += runXml(style, head);
    if (middle) deletedRuns.push(deletedRunXml(style, middle));
    rPr ??= style;
    if (tail) after += runXml(style, tail);
    if (unitEnd >= end) passed = true;
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
): string {
  const part = context.model.document;
  const rPr = pPr?.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
  const inner = innerPropertiesXml(part, rPr, ["ins", "del"]);
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
      const run = deletedRunXml(rPrOf(part, child), runContentOf(part, child));
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
    return {
      id: Number(part.attribute(node, "w:id") ?? -1),
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
