import type { TextRun, TextSelection } from "../../contracts.js";
import { normalizeWithMap } from "../../search-text.js";
import type { PageRect, TextRange } from "../types.js";
import type { PageLayout, TextLayoutLine } from "./types.js";

/*
 * Maps the viewer's text selection — PDF.js runs in page space — to the
 * elements and text ranges the PDFium engine edits. The two engines extract
 * text differently, so a run is matched to a layout line by a ladder: the
 * line whose box the run covers by at least half, else the one line that
 * contains the run's box, else a line whose folded text contains the run's.
 * The ladder follows GenOffice's technique (genspark-ai/genoffice,
 * `apps/pdf/src/main/text-edit.ts`, Apache-2.0); this is web-doc's own
 * implementation, see THIRD_PARTY_NOTICES.md.
 */

interface Candidate {
  readonly order: number;
  readonly elementId: string;
  readonly line: TextLayoutLine;
}

/** Slack, in points, when a run's box is compared with a line's. */
const SLACK = 0.5;

/** Ranges of the selected text, merged per element, in reading order. */
export function resolveSelection(
  selection: TextSelection,
  pages: readonly PageLayout[],
): TextRange[] {
  const candidates: Candidate[] = [];
  for (const page of pages)
    for (const layout of page.layouts)
      for (const line of layout.lines)
        candidates.push({
          order: candidates.length,
          elementId: layout.elementId,
          line,
        });
  const found: { readonly order: number; readonly range: TextRange }[] = [];
  for (const run of selection.runs) {
    if (!run.text.trim()) continue;
    const match = matchRun(run, candidates);
    if (!match) continue;
    const range = rangeWithin(run, match.line, match.elementId);
    if (range) found.push({ order: match.order, range });
  }
  return merge(found);
}

function matchRun(
  run: TextRun,
  candidates: readonly Candidate[],
): Candidate | undefined {
  const box = { x: run.x, y: run.y, width: run.width, height: run.height };
  // 1. The line whose box the run covers by at least half.
  let best: Candidate | undefined;
  let bestOverlap = 0;
  for (const candidate of candidates) {
    const overlap = overlapArea(box, candidate.line.bounds);
    if (overlap >= area(candidate.line.bounds) / 2 && overlap > bestOverlap) {
      best = candidate;
      bestOverlap = overlap;
    }
  }
  if (best) return best;
  // 2. The single line that contains the run's box.
  const containing = candidates.filter((candidate) =>
    contains(candidate.line.bounds, box),
  );
  if (containing.length > 0)
    return containing.reduce((smallest, candidate) =>
      area(candidate.line.bounds) < area(smallest.line.bounds)
        ? candidate
        : smallest,
    );
  // 3. A line whose folded text contains the run's folded text.
  const needle = fold(run.text).text;
  if (!needle) return undefined;
  return candidates.find((candidate) =>
    fold(candidate.line.text).text.includes(needle),
  );
}

/**
 * The part of the line the run covers. The run's text decides when it occurs
 * in the line — the viewer cuts runs proportionally, so their boxes are only
 * approximate — and its box picks the occurrence nearest to the glyphs it
 * covers; without a textual match the covered glyphs stand on their own.
 */
function rangeWithin(
  run: TextRun,
  line: TextLayoutLine,
  elementId: string,
): TextRange | undefined {
  const box = { x: run.x, y: run.y, width: run.width, height: run.height };
  const covered = line.glyphs
    .filter((glyph) => contains(grow(box, SLACK), centre(glyph.box)))
    .map((glyph) => glyph.offset);
  const anchor =
    covered.length > 0 ? Math.min(...covered) : line.range.start.offset;
  const haystack = fold(line.text);
  const needle = fold(run.text).text;
  const base = line.range.start.offset;
  let best: TextRange | undefined;
  let bestDistance = Number.POSITIVE_INFINITY;
  for (
    let at = needle ? haystack.text.indexOf(needle) : -1;
    at >= 0;
    at = haystack.text.indexOf(needle, at + 1)
  ) {
    const start = base + haystack.starts[at]!;
    const distance = Math.abs(start - anchor);
    if (distance < bestDistance) {
      bestDistance = distance;
      best = {
        start: { elementId, offset: start },
        end: {
          elementId,
          offset: base + haystack.ends[at + needle.length - 1]!,
        },
      };
    }
  }
  if (best) return best;
  if (covered.length === 0) return undefined;
  return {
    start: { elementId, offset: Math.min(...covered) },
    end: { elementId, offset: Math.max(...covered) + 1 },
  };
}

/** NFKC-folded text without whitespace, with each folded unit's source offsets. */
function fold(text: string): {
  readonly text: string;
  readonly starts: readonly number[];
  readonly ends: readonly number[];
} {
  const mapped = normalizeWithMap(text, false);
  let folded = "";
  const starts: number[] = [];
  const ends: number[] = [];
  for (let index = 0; index < mapped.text.length; index += 1) {
    const unit = mapped.text[index]!;
    if (/\s/u.test(unit)) continue;
    folded += unit;
    starts.push(mapped.starts[index]!);
    ends.push(mapped.ends[index]!);
  }
  return { text: folded, starts, ends };
}

function merge(
  found: readonly { readonly order: number; readonly range: TextRange }[],
): TextRange[] {
  const sorted = [...found].sort(
    (a, b) => a.order - b.order || a.range.start.offset - b.range.start.offset,
  );
  const merged: TextRange[] = [];
  for (const { range } of sorted) {
    const last = merged.at(-1);
    if (
      last &&
      last.end.elementId === range.start.elementId &&
      range.start.offset <= last.end.offset
    ) {
      if (range.end.offset > last.end.offset)
        merged[merged.length - 1] = { start: last.start, end: range.end };
    } else merged.push(range);
  }
  return merged;
}

function area(rect: PageRect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

function overlapArea(a: PageRect, b: PageRect): number {
  const width = Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x);
  const height = Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y);
  return width > 0 && height > 0 ? width * height : 0;
}

function contains(
  outer: PageRect,
  inner: PageRect | { readonly x: number; readonly y: number },
): boolean {
  const width = "width" in inner ? inner.width : 0;
  const height = "height" in inner ? inner.height : 0;
  return (
    inner.x >= outer.x - SLACK &&
    inner.y >= outer.y - SLACK &&
    inner.x + width <= outer.x + outer.width + SLACK &&
    inner.y + height <= outer.y + outer.height + SLACK
  );
}

function grow(rect: PageRect, by: number): PageRect {
  return {
    x: rect.x - by,
    y: rect.y - by,
    width: rect.width + 2 * by,
    height: rect.height + 2 * by,
  };
}

function centre(rect: PageRect): { readonly x: number; readonly y: number } {
  return { x: rect.x + rect.width / 2, y: rect.y + rect.height / 2 };
}
