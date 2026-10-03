import { patches, type XmlPatch } from "../ooxml/patch.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";
import type { TextRange } from "../types.js";
import { W_NS } from "./ids.js";
import type { DocxModel, ParagraphRecord } from "./model.js";
import type {
  DocxOperationContext,
  DocxOperationHandler,
  DocxOperationResult,
  Issue,
} from "./operations.js";
import type { RunItem } from "./text.js";
import {
  insertedRunXml,
  markedParagraphProperties,
  revisionAttributes,
  trackedRangeProblem,
  trackedSplit,
} from "./tracked.js";
import type {
  DocxReplaceTextOperation,
  DocxSetParagraphStyleOperation,
  DocxSetTextStyleOperation,
  DocxTextStyleChange,
} from "./types.js";
import {
  changedParagraphProperties,
  changedRunProperties,
  attributeProblem,
  colorProblem,
  namespacePatches,
  normalizeText,
  paragraphMarkProperties,
  paragraphPropertiesWithoutSection,
  paragraphXml,
  runContentXml,
  runXml,
  sliceOf,
  textProblem,
} from "./write.js";

/*
 * replaceText, setTextStyle and setParagraphStyle: a paragraph is rebuilt
 * from units — its direct children, a complex field's runs grouped as one
 * — where every unit the edit does not touch keeps its bytes, a touched
 * run is re-serialized around the change, a hyperlink or content control
 * is entered when the change lies inside it, and the paragraph is written
 * back as one element with its id, verified by re-scan.
 */

/** A direct child of a container, or the children a complex field spans. */
export interface Unit {
  readonly nodes: readonly XmlElement[];
  readonly kind: "zero" | "run" | "wrapper" | "field";
  /** Half-open text span; `start === end` for a unit without text. */
  readonly start: number;
  readonly end: number;
  readonly items: readonly RunItem[];
}

interface TextTarget {
  readonly record: ParagraphRecord;
  readonly part: XmlPart;
}

/** The child of `container` an item sits under. */
function topOf(node: XmlElement, container: XmlElement): XmlElement {
  let current = node;
  while (current.parent && current.parent !== container)
    current = current.parent;
  return current;
}

/** The units of a container's children for the items it holds. */
export function unitsOf(
  container: XmlElement,
  items: readonly RunItem[],
): Unit[] {
  const byTop = new Map<XmlElement, RunItem[]>();
  for (const item of items) {
    const top = topOf(item.run, container);
    const list = byTop.get(top) ?? [];
    list.push(item);
    byTop.set(top, list);
  }
  // A complex field spans consecutive children: from its begin run's top
  // to its end run's top.
  const groupEnd = new Map<XmlElement, XmlElement>();
  for (const item of items)
    if (item.fieldRuns && item.fieldRuns.length > 1) {
      const first = topOf(item.fieldRuns[0]!, container);
      const last = topOf(item.fieldRuns.at(-1)!, container);
      if (first !== last) groupEnd.set(first, last);
    }
  const units: Unit[] = [];
  let position = 0;
  const children = container.children.filter(
    (child) => !(child.local === "pPr" && child.namespace === W_NS),
  );
  for (let index = 0; index < children.length; index += 1) {
    const child = children[index]!;
    const end = groupEnd.get(child);
    const nodes: XmlElement[] = [child];
    if (end) {
      let cursor = index + 1;
      while (cursor < children.length && children[cursor - 1] !== end)
        nodes.push(children[cursor++]!);
      index = cursor - 1;
    }
    const own = nodes.flatMap((node) => byTop.get(node) ?? []);
    const kind: Unit["kind"] =
      own.length === 0
        ? "zero"
        : end ||
            child.local === "fldSimple" ||
            own.some((item) => item.fieldRuns)
          ? "field"
          : child.local === "r"
            ? "run"
            : "wrapper";
    const start = own.length > 0 ? own[0]!.start : position;
    const stop = own.length > 0 ? own.at(-1)!.end : position;
    units.push({ nodes, kind, start, end: stop, items: own });
    position = stop;
  }
  return units;
}

export function rPrOf(part: XmlPart, item: RunItem | undefined): string {
  return item ? sliceOf(part, item.rPr) : "";
}

function startTagOf(part: XmlPart, node: XmlElement): string {
  return part.text.slice(node.start, node.contentStart);
}

function endTagOf(part: XmlPart, node: XmlElement): string {
  return part.text.slice(node.contentEnd, node.end);
}

/**
 * The content of a run's items in `[from, to]`: `w:t` text sliced, every
 * other child (tab, break, picture, symbol, note reference) as its bytes
 * when it lies inside. A zero-width item at `from` belongs to the slice
 * only when `includeStart` says so, which is how a caret or a range edge
 * hands such an item to exactly one side.
 */
export function contentOf(
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

/**
 * Inline objects a replacement of `[from, to)` drops: pictures, objects
 * and anchors inside the range (a zero-width one only strictly inside),
 * by element id, with the ids of every paragraph nested in them (a text
 * box) so the engine's id list follows.
 */
function droppedInlines(
  model: DocxModel,
  record: ParagraphRecord,
  items: readonly RunItem[],
  from: number,
  to: number,
): { removedIds: string[]; removedParagraphIds: string[] } {
  const removedIds: string[] = [];
  const removedParagraphIds: string[] = [];
  for (const item of items) {
    if (
      item.kind !== "picture" &&
      item.kind !== "object" &&
      item.kind !== "anchor"
    )
      continue;
    const inside =
      item.start === item.end
        ? from < item.start && item.start < to
        : item.start >= from && item.end <= to;
    if (!inside) continue;
    const inline = record.inlines.find(
      (candidate) => candidate.node === item.child,
    );
    if (inline) removedIds.push(inline.elementId);
    for (const paragraph of model.document.findAll("p", item.child)) {
      const id = model.paragraphIds.get(paragraph);
      if (id && paragraph.namespace === W_NS) removedParagraphIds.push(id);
    }
  }
  return { removedIds, removedParagraphIds };
}

interface Split {
  readonly before: string;
  readonly after: string;
  /** The `w:rPr` bytes new text takes. */
  readonly rPr: string;
  /** The insertion point lies inside a hyperlink or content control. */
  readonly nested: boolean;
}

/**
 * The container's content with `[start, end)` removed and `insert(rPr)`
 * written at `start`: `before` holds everything up to and including the
 * inserted text, `after` what follows. New text is styled like the run
 * that held the first replaced character, like the run before a caret,
 * else like the run after it.
 */
function splitAt(
  part: XmlPart,
  container: XmlElement,
  items: readonly RunItem[],
  start: number,
  end: number,
  insert: (rPr: string) => string,
  fallbackRPr: string,
): Split {
  let before = "";
  let after = "";
  let inserted = false;
  let nested = false;
  let rPr = fallbackRPr;
  let lastRPr: string | undefined;
  const raw = (unit: Unit): string =>
    unit.nodes.map((node) => sliceOf(part, node)).join("");
  const place = (style: string | undefined): void => {
    rPr = style ?? lastRPr ?? fallbackRPr;
    before += insert(rPr);
    inserted = true;
  };
  for (const unit of unitsOf(container, items)) {
    if (unit.kind === "zero") {
      if (inserted) after += raw(unit);
      else before += raw(unit);
      continue;
    }
    if (unit.end <= start) {
      before += raw(unit);
      lastRPr = rPrOf(part, unit.items.at(-1));
      if (!inserted && start === end && unit.end === start) place(lastRPr);
      continue;
    }
    if (unit.start >= end) {
      if (!inserted) place(lastRPr ?? rPrOf(part, unit.items[0]));
      after += raw(unit);
      continue;
    }
    if (unit.start >= start && unit.end <= end) {
      // Covered whole: dropped; the first such unit styles the new text.
      if (!inserted) place(rPrOf(part, unit.items[0]));
      continue;
    }
    if (unit.kind === "run") {
      const style = rPrOf(part, unit.items[0]);
      // The head keeps a zero-width item at `start`; the tail takes one at
      // `end` only when the range is not a caret (the head has it then).
      const head = runXml(
        style,
        contentOf(part, unit.items, unit.start, start, true),
      );
      const tail = runXml(
        style,
        contentOf(part, unit.items, end, unit.end, start < end),
      );
      if (!inserted) {
        before += head;
        place(style);
        after += tail;
      } else after += tail;
      continue;
    }
    if (unit.kind === "wrapper" && unit.start < start && end < unit.end) {
      // Both ends inside: the wrapper stays and its runs change.
      const wrapper = unit.nodes[0]!;
      const inner = splitAt(
        part,
        wrapper,
        unit.items,
        start,
        end,
        insert,
        fallbackRPr,
      );
      before +=
        startTagOf(part, wrapper) +
        inner.before +
        inner.after +
        endTagOf(part, wrapper);
      rPr = inner.rPr;
      inserted = true;
      nested = true;
      continue;
    }
    // A field cut, a wrapper edge cut: validation refuses these.
    throw new Error("A range cannot cut a field or a wrapper");
  }
  if (!inserted) place(lastRPr);
  return { before, after, rPr, nested };
}

async function textTarget(
  id: string,
  context: DocxOperationContext,
  issue: Issue,
): Promise<TextTarget | undefined> {
  const record = context.model.byId.get(id);
  if (!record || record.kind !== "paragraph") {
    issue(
      "/target",
      record ? "invalid-target" : "unknown-target",
      record ? `Element ${id} is not a paragraph` : `No element ${id}`,
    );
    return undefined;
  }
  if (record.readOnlyReason) {
    issue(
      "/target",
      "invalid-target",
      `Paragraph ${id} is read-only (${record.readOnlyReason})`,
    );
    return undefined;
  }
  return { record, part: context.model.document };
}

/** Checks a range against its target; returns the clamped offsets. */
function checkRange(
  target: TextTarget,
  id: string,
  range: TextRange | undefined,
  issue: Issue,
  structural: boolean,
): { start: number; end: number } | undefined {
  const text = target.record.text.text;
  const length = text.length;
  if (!range) return { start: 0, end: length };
  if (range.start.elementId !== id || range.end.elementId !== id) {
    issue("/range", "invalid-range", "Both ends must lie on the target");
    return undefined;
  }
  const { start, end } = { start: range.start.offset, end: range.end.offset };
  if (start > length || end > length || start > end) {
    issue(
      "/range",
      "invalid-range",
      `Offsets must satisfy 0 ≤ start ≤ end ≤ ${length}`,
    );
    return undefined;
  }
  for (const offset of [start, end])
    if (
      offset > 0 &&
      offset < length &&
      (text.charCodeAt(offset) & 0xfc00) === 0xdc00 &&
      (text.charCodeAt(offset - 1) & 0xfc00) === 0xd800
    ) {
      issue("/range", "invalid-range", "A range cannot split a surrogate pair");
      return undefined;
    }
  if (!structural) return { start, end };
  const inside = (unit: Unit, offset: number): boolean =>
    unit.start < offset && offset < unit.end;
  const check = (container: XmlElement, items: readonly RunItem[]): boolean => {
    for (const unit of unitsOf(container, items)) {
      if (unit.kind === "field" && (inside(unit, start) || inside(unit, end))) {
        issue(
          "/range",
          "invalid-range",
          "A range cannot start or end inside a field",
        );
        return false;
      }
      if (unit.kind !== "wrapper") continue;
      const startIn = inside(unit, start);
      const endIn = inside(unit, end);
      if (startIn !== endIn) {
        issue(
          "/range",
          "invalid-range",
          "A range cannot cut the edge of a hyperlink or content control",
        );
        return false;
      }
      if (startIn && !check(unit.nodes[0]!, unit.items)) return false;
    }
    return true;
  };
  return check(target.record.node, target.record.text.items)
    ? { start, end }
    : undefined;
}

function commitParagraph(
  context: DocxOperationContext,
  target: TextTarget,
  items: readonly XmlPatch[],
  extra: Partial<DocxOperationResult> = {},
): Promise<DocxOperationResult> {
  const { part, record } = target;
  const transaction = context.pkg.transaction();
  transaction.patch(part, [...namespacePatches(part), ...items]);
  return transaction.commit().then((change) => ({
    createdIds: [],
    warnings: change.warnings,
    ...(context.model.unauthoredSet.has(record.id)
      ? { stamped: [record.id] }
      : {}),
    reflowFrom: record.id,
    ...extra,
  }));
}

export const replaceTextHandler: DocxOperationHandler<DocxReplaceTextOperation> =
  {
    async validate(operation, context, issue) {
      const text = normalizeText(operation.text);
      const style = operation.insertionStyle;
      if (style !== undefined) {
        if (text.length === 0)
          issue(
            "/insertionStyle",
            "invalid-value",
            "An insertion style requires nonempty replacement text",
          );
        if (style.color !== undefined) {
          const invalid = colorProblem(style.color);
          if (invalid)
            issue(
              "/insertionStyle/color",
              "invalid-value",
              `Colour: ${invalid}`,
            );
        }
        if (style.fontFamily !== undefined) {
          const invalid = attributeProblem(style.fontFamily);
          if (invalid)
            issue(
              "/insertionStyle/fontFamily",
              "invalid-value",
              `Font: ${invalid}`,
            );
        }
      }
      const problem = textProblem(text);
      if (problem) issue("/text", "invalid-text", `The text holds ${problem}`);
      const target = await textTarget(operation.target, context, issue);
      if (!target) return;
      const range = checkRange(
        target,
        operation.target,
        operation.range,
        issue,
        true,
      );
      if (
        range &&
        context.tracked &&
        trackedRangeProblem(
          target.record,
          range.start,
          range.end,
          issue,
          "/range",
        )
      )
        return;
      if (!range || !text.includes("\n")) return;
      // A paragraph cannot be split inside a hyperlink or content control.
      const probe = splitAt(
        target.part,
        target.record.node,
        target.record.text.items,
        range.start,
        range.end,
        () => "",
        "",
      );
      if (probe.nested)
        issue(
          "/text",
          "invalid-range",
          "A paragraph break cannot be inserted inside a hyperlink or content control",
        );
    },
    async apply(operation, context) {
      const target = (await textTarget(operation.target, context, () => {}))!;
      const { start, end } = checkRange(
        target,
        operation.target,
        operation.range,
        () => {},
        true,
      )!;
      const { items, createdIds, removedIds, removedParagraphIds } =
        replacedParagraph(
          context,
          target.record,
          start,
          end,
          normalizeText(operation.text),
          operation.insertionStyle,
        );
      return commitParagraph(context, target, items, {
        createdIds,
        ...(removedIds.length > 0 ? { removedIds } : {}),
        ...(removedParagraphIds.length > 0 ? { removedParagraphIds } : {}),
      });
    },
  };

/**
 * The patches that replace `[start, end)` of a paragraph's text with
 * `text`, newlines splitting it into new paragraphs after it: the first
 * paragraph rebuilt in place, the others inserted with fresh ids.
 */
export function replacedParagraph(
  context: DocxOperationContext,
  record: ParagraphRecord,
  start: number,
  end: number,
  text: string,
  insertionStyle?: DocxTextStyleChange,
): {
  items: XmlPatch[];
  createdIds: string[];
  removedIds: string[];
  removedParagraphIds: string[];
} {
  const part = context.model.document;
  const segments = text.split("\n");
  const single = segments.length === 1;
  const items: XmlPatch[] = [];
  const createdIds: string[] = [];
  // The splitter chooses an existing run or paragraph mark's exact properties.
  // Merge only inserted runs and new marks; surrounding XML stays untouched.
  const sourceProperties = (raw: string): XmlElement | undefined => {
    const mark = record.pPr?.children.find(
      (child) => child.local === "rPr" && child.namespace === W_NS,
    );
    return (
      record.text.items.find((item) => sliceOf(part, item.rPr) === raw)?.rPr ??
      (sliceOf(part, mark) === raw ? mark : undefined)
    );
  };
  const insertion = (raw: string) =>
    insertionStyle
      ? {
          style: insertionStyle,
          styles: context.model.styles,
          source: sourceProperties(raw),
        }
      : undefined;
  const insertedProperties = (raw: string): string =>
    insertionStyle
      ? changedRunProperties(
          part,
          sourceProperties(raw),
          insertionStyle,
          context.model.styles,
        )
      : raw;
  if (context.tracked) {
    // The removed runs stay as a deletion, the new text goes in as an
    // insertion; a paragraph split marks the new paragraph marks inserted
    // and leaves the original mark on the last paragraph.
    const split = trackedSplit(
      context,
      record,
      start,
      end,
      paragraphMarkRPr(part, record),
    );
    const runProperties = insertedProperties(split.rPr);
    const inserted = (segment: string): string =>
      insertedRunXml(context, runProperties, segment);
    // The original mark (section properties included) ends the last
    // paragraph; every mark before it is an insertion.
    const originalPPr = sliceOf(part, record.pPr);
    const firstPPr = single
      ? originalPPr
      : markedParagraphProperties(
          context,
          record.pPr,
          "ins",
          insertion(split.rPr),
        );
    items.push(
      patches.replaceElement(
        part,
        record.node,
        paragraphXml(
          record.node,
          record.id,
          firstPPr,
          split.before +
            split.deleted +
            inserted(segments[0]!) +
            (single ? split.after : ""),
        ),
      ),
    );
    segments.slice(1).forEach((segment, index) => {
      const id = context.freshParagraphId();
      createdIds.push(`p:${id}`);
      const last = index === segments.length - 2;
      items.push(
        patches.insertAfter(
          part,
          record.node,
          paragraphXml(
            record.node,
            id,
            last
              ? originalPPr
              : markedParagraphProperties(
                  context,
                  record.pPr,
                  "ins",
                  insertion(split.rPr),
                ),
            inserted(segment) + (last ? split.after : ""),
          ),
        ),
      );
    });
  } else {
    const split = splitAt(
      part,
      record.node,
      record.text.items,
      start,
      end,
      (rPr) => runXml(insertedProperties(rPr), runContentXml(segments[0]!)),
      paragraphMarkRPr(part, record),
    );
    const pPr = sliceOf(part, record.pPr);
    const runProperties = insertedProperties(split.rPr);
    const copiedPPr = paragraphPropertiesWithoutSection(
      part,
      record.pPr,
      insertion(split.rPr),
    );
    items.push(
      patches.replaceElement(
        part,
        record.node,
        paragraphXml(
          record.node,
          record.id,
          pPr,
          split.before + (single ? split.after : ""),
        ),
      ),
    );
    segments.slice(1).forEach((segment, index) => {
      const id = context.freshParagraphId();
      createdIds.push(`p:${id}`);
      const last = index === segments.length - 2;
      items.push(
        patches.insertAfter(
          part,
          record.node,
          paragraphXml(
            record.node,
            id,
            copiedPPr,
            runXml(runProperties, runContentXml(segment)) +
              (last ? split.after : ""),
          ),
        ),
      );
    });
  }
  const dropped = droppedInlines(
    context.model,
    record,
    record.text.items,
    start,
    end,
  );
  const removedIds = dropped.removedIds;
  // Inline objects that moved to a new paragraph change their id.
  if (!single)
    for (const inline of record.inlines) {
      const item = record.text.items.find(
        (candidate) => candidate.child === inline.node,
      );
      if (item && item.start >= end && !removedIds.includes(inline.elementId))
        removedIds.push(inline.elementId);
    }
  return {
    items,
    createdIds,
    removedIds,
    removedParagraphIds: dropped.removedParagraphIds,
  };
}

/** The paragraph mark's `w:rPr`, the style of text in an empty paragraph. */
function paragraphMarkRPr(part: XmlPart, record: ParagraphRecord): string {
  const rPr = record.pPr?.children.find(
    (child) => child.local === "rPr" && child.namespace === W_NS,
  );
  return sliceOf(part, rPr);
}

/** A run re-serialized with new properties around its original content. */
function restyledRun(part: XmlPart, run: XmlElement, rPr: string): string {
  const content = run.children
    .filter((child) => !(child.local === "rPr" && child.namespace === W_NS))
    .map((child) => sliceOf(part, child))
    .join("");
  return `<w:r>${rPr}${content}</w:r>`;
}

/** The container's content with the runs in `[start, end)` restyled. */
function restyled(
  part: XmlPart,
  container: XmlElement,
  items: readonly RunItem[],
  start: number,
  end: number,
  change: DocxTextStyleChange,
  context: DocxOperationContext,
): string {
  let out = "";
  const changed = (rPr: XmlElement | undefined): string =>
    changedRunProperties(
      part,
      rPr,
      change,
      context.model.styles,
      context.tracked ? revisionAttributes(context) : undefined,
    );
  for (const unit of unitsOf(container, items)) {
    const raw = unit.nodes.map((node) => sliceOf(part, node)).join("");
    const overlaps =
      unit.start < end && start < unit.end
        ? true
        : unit.start === unit.end && start < unit.start && unit.start < end;
    if (unit.kind === "zero" || !overlaps) {
      out += raw;
      continue;
    }
    if (unit.kind === "run") {
      const run = unit.nodes[0]!;
      const own = rPrOf(part, unit.items[0]);
      const rPr = unit.items[0]?.rPr;
      out += runXml(
        own,
        contentOf(
          part,
          unit.items,
          unit.start,
          Math.min(start, unit.end),
          true,
        ),
      );
      out += runXml(
        changed(rPr),
        contentOf(
          part,
          unit.items,
          Math.max(unit.start, start),
          Math.min(unit.end, end),
          false,
        ),
      );
      out += runXml(
        own,
        contentOf(part, unit.items, Math.max(end, unit.start), unit.end, false),
      );
      continue;
    }
    if (unit.kind === "field") {
      // A field is styled whole: every run of it takes the change.
      for (const node of unit.nodes)
        if (node.local === "r" && node.namespace === W_NS)
          out += restyledRun(
            part,
            node,
            changed(
              node.children.find(
                (child) => child.local === "rPr" && child.namespace === W_NS,
              ),
            ),
          );
        else if (node.local === "fldSimple")
          out +=
            startTagOf(part, node) +
            restyled(part, node, unit.items, start, end, change, context) +
            endTagOf(part, node);
        else out += sliceOf(part, node);
      continue;
    }
    const wrapper = unit.nodes[0]!;
    out +=
      startTagOf(part, wrapper) +
      restyled(part, wrapper, unit.items, start, end, change, context) +
      endTagOf(part, wrapper);
  }
  return out;
}

export const setTextStyleHandler: DocxOperationHandler<DocxSetTextStyleOperation> =
  {
    async validate(operation, context, issue) {
      const { style } = operation;
      if (style.color !== undefined) {
        const problem = colorProblem(style.color);
        if (problem)
          issue("/style/color", "invalid-value", `Colour: ${problem}`);
      }
      if (style.fontFamily !== undefined) {
        const problem = attributeProblem(style.fontFamily);
        if (problem)
          issue("/style/fontFamily", "invalid-value", `Font: ${problem}`);
      }
      const target = await textTarget(operation.target, context, issue);
      if (target)
        checkRange(target, operation.target, operation.range, issue, false);
    },
    async apply(operation, context) {
      const target = (await textTarget(operation.target, context, () => {}))!;
      const { part, record } = target;
      const { start, end } = checkRange(
        target,
        operation.target,
        operation.range,
        () => {},
        false,
      )!;
      if (Object.keys(operation.style).length === 0)
        return { createdIds: [], warnings: [] };
      const content = restyled(
        part,
        record.node,
        record.text.items,
        start,
        end,
        operation.style,
        context,
      );
      const pPr =
        end >= record.text.text.length
          ? paragraphMarkProperties(
              part,
              record.pPr,
              operation.style,
              context.model.styles,
              context.tracked ? revisionAttributes(context) : undefined,
            )
          : sliceOf(part, record.pPr);
      return commitParagraph(context, target, [
        patches.replaceElement(
          part,
          record.node,
          paragraphXml(record.node, record.id, pPr, content),
        ),
      ]);
    },
  };

export const setParagraphStyleHandler: DocxOperationHandler<DocxSetParagraphStyleOperation> =
  {
    async validate(operation, context, issue) {
      await textTarget(operation.target, context, issue);
    },
    async apply(operation, context) {
      const target = (await textTarget(operation.target, context, () => {}))!;
      const { part, record } = target;
      const { style } = operation;
      if (
        style.align === undefined &&
        (style.spacing === undefined || Object.keys(style.spacing).length === 0)
      )
        return { createdIds: [], warnings: [] };
      const pPr = changedParagraphProperties(
        part,
        record.pPr,
        style,
        context.tracked ? revisionAttributes(context) : undefined,
      );
      const node = record.node;
      const content = node.selfClosing
        ? ""
        : part.text.slice(
            record.pPr ? record.pPr.end : node.contentStart,
            node.contentEnd,
          );
      return commitParagraph(context, target, [
        patches.replaceElement(
          part,
          node,
          paragraphXml(node, record.id, pPr, content),
        ),
      ]);
    },
  };
