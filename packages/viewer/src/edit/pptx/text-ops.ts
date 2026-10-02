import { patches } from "../ooxml/patch.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { TextRange } from "../types.js";
import type { ShapeRecord } from "./elements.js";
import {
  committedParts,
  type Issue,
  type PptxOperationContext,
  type PptxOperationHandler,
  type PptxOperationResult,
} from "./operations.js";
import type { ParagraphInfo, RunInfo, TextModel } from "./text.js";
import {
  alignedParagraphProperties,
  bodyPrefix,
  bodySuffix,
  changedRunProperties,
  colorProblem,
  itemsOfSegment,
  normalizeText,
  paragraphAt,
  paragraphXml,
  renamedProperties,
  runProperties,
  sliceOf,
  styleAt,
  textProblem,
  PARAGRAPH_BREAK,
  type ParagraphDraft,
  type RunItem,
} from "./text-write.js";
import type {
  PptxReplaceTextOperation,
  PptxSetTextStyleOperation,
} from "./types.js";

/*
 * replaceText and setTextStyle: the paragraphs a range touches are rebuilt
 * from their runs — untouched runs as their original bytes, touched ones
 * re-serialized with their properties — and written back as one content
 * patch of the text body, verified by re-scan.
 */

interface TextTarget {
  readonly record: ShapeRecord;
  readonly part: XmlPart;
  readonly txBody: XmlElement;
  readonly model: TextModel;
}

/** Resolves a text target; reports why it cannot be edited. */
async function textTarget(
  id: string,
  context: PptxOperationContext,
  issue: Issue,
): Promise<TextTarget | undefined> {
  const record = await context.locate(id);
  if (!record) {
    issue("/target", "unknown-target", `No element ${id}`);
    return undefined;
  }
  if (record.readOnly || !record.txBody || !record.text) {
    issue(
      "/target",
      "invalid-target",
      `Element ${id} has no editable text body`,
    );
    return undefined;
  }
  return {
    record,
    part: record.part,
    txBody: record.txBody,
    model: record.text,
  };
}

/** Checks a range against its target; returns the clamped offsets. */
function checkRange(
  target: TextTarget,
  id: string,
  range: TextRange | undefined,
  issue: Issue,
): { start: number; end: number } | undefined {
  const length = target.model.text.length;
  if (!range) return { start: 0, end: length };
  if (range.start.elementId !== id || range.end.elementId !== id) {
    issue("/range", "invalid-range", "Both ends must lie on the target");
    return undefined;
  }
  const { start } = range;
  const { end } = range;
  if (
    start.offset > length ||
    end.offset > length ||
    start.offset > end.offset
  ) {
    issue(
      "/range",
      "invalid-range",
      `Offsets must satisfy 0 ≤ start ≤ end ≤ ${length}`,
    );
    return undefined;
  }
  const text = target.model.text;
  for (const offset of [start.offset, end.offset])
    if (
      offset > 0 &&
      offset < length &&
      (text.charCodeAt(offset) & 0xfc00) === 0xdc00 &&
      (text.charCodeAt(offset - 1) & 0xfc00) === 0xd800
    ) {
      issue("/range", "invalid-range", "A range cannot split a surrogate pair");
      return undefined;
    }
  for (const paragraph of target.model.paragraphs)
    for (const run of paragraph.runs)
      if (
        run.kind === "fld" &&
        ((run.start < start.offset && start.offset < run.end) ||
          (run.start < end.offset && end.offset < run.end))
      ) {
        issue(
          "/range",
          "invalid-range",
          "A range cannot start or end inside a field",
        );
        return undefined;
      }
  return { start: start.offset, end: end.offset };
}

/** Items of a run restricted to `[from, to)` of its text, as bytes when whole. */
function runItems(
  part: XmlPart,
  run: RunInfo,
  from: number,
  to: number,
): RunItem[] {
  if (to <= from) return [];
  if (from === run.start && to === run.end)
    return [{ kind: "raw", xml: sliceOf(part, run.node) }];
  if (run.kind === "br") return [{ kind: "br", rPr: runProperties(part, run) }];
  return [
    {
      kind: "run",
      rPr: runProperties(part, run),
      text: run.text.slice(from - run.start, to - run.start),
    },
  ];
}

function draftOf(part: XmlPart, paragraph: ParagraphInfo): ParagraphDraft {
  return {
    pPr: sliceOf(part, paragraph.pPr),
    items: [],
    endParaRPr: sliceOf(part, paragraph.endParaRPr),
  };
}

/** The whole body's new content with the paragraphs `[first, last]` replaced by `drafts`. */
function bodyContent(
  part: XmlPart,
  txBody: XmlElement,
  model: TextModel,
  first: number,
  last: number,
  drafts: readonly ParagraphDraft[],
): string {
  const paragraphs = model.paragraphs;
  const prefix = bodyPrefix(part, txBody, paragraphs[0]?.node);
  const suffix = bodySuffix(part, txBody, paragraphs.at(-1)?.node);
  const before = paragraphs
    .slice(0, first)
    .map((paragraph) => sliceOf(part, paragraph.node))
    .join("");
  const after = paragraphs
    .slice(last + 1)
    .map((paragraph) => sliceOf(part, paragraph.node))
    .join("");
  return `${prefix}${before}${drafts.map(paragraphXml).join("")}${after}${suffix}`;
}

function commitBody(
  context: PptxOperationContext,
  target: TextTarget,
  content: string,
): Promise<PptxOperationResult> {
  const transaction = context.pkg.transaction();
  transaction.patch(target.part, [
    patches.replaceContent(target.part, target.txBody, content),
  ]);
  return transaction.commit().then((change) => ({
    createdIds: [],
    changedPages: [target.record.element.pageIndex],
    warnings: change.warnings,
    parts: committedParts(change),
  }));
}

export const replaceTextHandler: PptxOperationHandler<PptxReplaceTextOperation> =
  {
    async validate(operation, context, issue) {
      const problem = textProblem(normalizeText(operation.text));
      if (problem) issue("/text", "invalid-text", `The text holds ${problem}`);
      const target = await textTarget(operation.target, context, issue);
      if (target) checkRange(target, operation.target, operation.range, issue);
    },
    async apply(operation, context) {
      const target = (await textTarget(operation.target, context, () => {}))!;
      const { start, end } = checkRange(
        target,
        operation.target,
        operation.range,
        () => {},
      )!;
      return commitBody(
        context,
        target,
        replacedBodyContent(
          target.part,
          target.txBody,
          target.model,
          normalizeText(operation.text),
          start,
          end,
          !operation.range,
        ),
      );
    },
  };

/**
 * The new content of a text body with `[start, end)` of its text replaced:
 * the paragraphs the range touches are rebuilt, the others keep their
 * bytes. A whole replacement styles the new text like the first run that
 * had text; a ranged one like the run at the start of the range.
 */
export function replacedBodyContent(
  part: XmlPart,
  txBody: XmlElement,
  model: TextModel,
  text: string,
  start: number,
  end: number,
  whole: boolean,
): string {
  const paragraphs = model.paragraphs;
  if (paragraphs.length === 0) {
    // A body without paragraphs: the new text becomes its paragraphs.
    const drafts = text.split(PARAGRAPH_BREAK).map((segment) => ({
      pPr: "",
      items: itemsOfSegment(segment, ""),
      endParaRPr: "",
    }));
    return bodyContent(part, txBody, model, 0, -1, drafts);
  }
  const first = paragraphAt(model, start);
  const last = paragraphAt(model, end);
  const firstParagraph = paragraphs[first]!;
  const lastParagraph = paragraphs[last]!;
  const style = whole
    ? firstTextStyle(part, model)
    : end > start
      ? (styleOfRunAt(part, firstParagraph, start) ??
        styleAt(part, firstParagraph, start))
      : styleAt(part, firstParagraph, start);
  const head: RunItem[] = [];
  for (const run of firstParagraph.runs)
    head.push(...runItems(part, run, run.start, Math.min(run.end, start)));
  const tail: RunItem[] = [];
  for (const run of lastParagraph.runs)
    tail.push(...runItems(part, run, Math.max(run.start, end), run.end));
  const segments = text.split(PARAGRAPH_BREAK);
  const drafts: ParagraphDraft[] = [];
  segments.forEach((segment, index) => {
    const draft = draftOf(part, firstParagraph);
    if (index === 0) draft.items.push(...head);
    draft.items.push(...itemsOfSegment(segment, style));
    if (index === segments.length - 1) {
      draft.items.push(...tail);
      draft.endParaRPr = sliceOf(part, lastParagraph.endParaRPr);
    }
    drafts.push(draft);
  });
  return bodyContent(part, txBody, model, first, last, drafts);
}

/** The rPr of the run that holds the first replaced character. */
function styleOfRunAt(
  part: XmlPart,
  paragraph: ParagraphInfo,
  offset: number,
): string | undefined {
  const run = paragraph.runs.find(
    (candidate) =>
      candidate.kind !== "br" &&
      candidate.start <= offset &&
      offset < candidate.end,
  );
  return run ? runProperties(part, run) : undefined;
}

function firstTextStyle(part: XmlPart, model: TextModel): string {
  for (const paragraph of model.paragraphs)
    for (const run of paragraph.runs)
      if (run.kind !== "br" && run.text.length > 0)
        return runProperties(part, run);
  const first = model.paragraphs[0];
  return first?.endParaRPr
    ? renamedProperties(part, first.endParaRPr, "a:rPr")
    : "";
}

export const setTextStyleHandler: PptxOperationHandler<PptxSetTextStyleOperation> =
  {
    async validate(operation, context, issue) {
      const { style } = operation;
      if (style.color !== undefined) {
        const problem = colorProblem(style.color);
        if (problem)
          issue("/style/color", "invalid-value", `Colour: ${problem}`);
      }
      if (style.fontFamily !== undefined) {
        const problem = textProblem(style.fontFamily);
        if (problem)
          issue("/style/fontFamily", "invalid-value", `Font: ${problem}`);
      }
      const target = await textTarget(operation.target, context, issue);
      if (target) checkRange(target, operation.target, operation.range, issue);
    },
    async apply(operation, context) {
      const target = (await textTarget(operation.target, context, () => {}))!;
      const { start, end } = checkRange(
        target,
        operation.target,
        operation.range,
        () => {},
      )!;
      const { part, model } = target;
      const { style } = operation;
      const { align, ...runChange } = style;
      const changesRuns = Object.keys(runChange).length > 0;
      // Nothing to write: the body keeps its bytes and its autofit scale.
      if (!changesRuns && align === undefined)
        return { createdIds: [], changedPages: [], warnings: [] };
      const paragraphs = model.paragraphs;
      if (paragraphs.length === 0)
        return { createdIds: [], changedPages: [], warnings: [] };
      const first = paragraphAt(model, start);
      const last = paragraphAt(model, end);
      const drafts: ParagraphDraft[] = [];
      for (let index = first; index <= last; index += 1) {
        const paragraph = paragraphs[index]!;
        const draft = draftOf(part, paragraph);
        if (align !== undefined)
          draft.pPr = alignedParagraphProperties(part, paragraph.pPr, align);
        for (const run of paragraph.runs) {
          const from = Math.max(run.start, start);
          const to = Math.min(run.end, end);
          if (!changesRuns || to <= from) {
            draft.items.push({ kind: "raw", xml: sliceOf(part, run.node) });
            continue;
          }
          draft.items.push(...runItems(part, run, run.start, from));
          draft.items.push(restyled(part, run, from, to, runChange));
          draft.items.push(...runItems(part, run, to, run.end));
        }
        if (changesRuns && end >= paragraph.end && paragraph.endParaRPr)
          draft.endParaRPr = changedRunProperties(
            part,
            paragraph.endParaRPr,
            runChange,
            "a:endParaRPr",
          );
        drafts.push(draft);
      }
      return commitBody(
        context,
        target,
        bodyContent(part, target.txBody, model, first, last, drafts),
      );
    },
  };

/** The covered part of a run with the change applied; a field keeps its own bytes around a new rPr. */
function restyled(
  part: XmlPart,
  run: RunInfo,
  from: number,
  to: number,
  change: Omit<PptxSetTextStyleOperation["style"], "align">,
): RunItem {
  const rPr = changedRunProperties(part, run.rPr, change);
  if (run.kind === "br") return { kind: "br", rPr };
  if (run.kind === "fld") {
    const xml = sliceOf(part, run.node);
    const offset = run.node.start;
    if (run.rPr)
      return {
        kind: "raw",
        xml:
          xml.slice(0, run.rPr.start - offset) +
          rPr +
          xml.slice(run.rPr.end - offset),
      };
    const at = run.node.contentStart - offset;
    return { kind: "raw", xml: `${xml.slice(0, at)}${rPr}${xml.slice(at)}` };
  }
  return {
    kind: "run",
    rPr,
    text: run.text.slice(from - run.start, to - run.start),
  };
}
