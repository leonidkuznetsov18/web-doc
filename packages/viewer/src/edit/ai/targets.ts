import { alignFuzzyPassage } from "../../fuzzy-alignment.js";
import {
  DEFAULT_FUZZY_SEARCH_OPTIONS,
  FuzzyPageIndex,
  type FuzzyPageText,
} from "../../fuzzy-search.js";
import { graphemeSegments, normalizeSearchText } from "../../search-text.js";
import type {
  EditElement,
  EditFindOptions,
  ElementQuery,
  ReadOptions,
  ReadResult,
  TextRange,
  TextTarget,
} from "../types.js";
import { foldText } from "./outline.js";
import type { TargetCandidate, TargetQuery } from "./types.js";

/*
 * A model names what it wants to edit by quoting it ("the sentence that
 * starts with…"), by citing it with a page, or by kind ("the first table on
 * slide 3"). Resolution turns that into element ids and ranges in named
 * steps, so a host knows how much to trust a candidate: the engine's exact
 * search first, then a match with whitespace, case, quotes and compatibility
 * forms folded, then the viewer's fuzzy citation matching over the
 * elements' text, and a kind lookup when there is no text at all.
 */

const DEFAULT_MAX_RESULTS = 5;
const SNIPPET_CONTEXT = 40;
/** Fuzzy candidates score from this down to 0.5 at the loosest accepted edit ratio. */
const FUZZY_TOP = 0.85;

/** What resolution needs of a session: its elements and its exact text search. */
export interface TargetSource {
  getElements(
    query?: ElementQuery,
    options?: ReadOptions,
  ): Promise<ReadResult<EditElement>>;
  findText(
    query: string,
    options?: EditFindOptions,
  ): Promise<ReadResult<TextTarget>>;
}

export async function resolveTargets(
  session: TargetSource,
  query: TargetQuery,
  options: ReadOptions = {},
): Promise<ReadResult<TargetCandidate>> {
  const needle = (query.text ?? query.citation?.text ?? "").trim();
  const hint =
    query.citation?.pageNumber !== undefined &&
    Number.isFinite(query.citation.pageNumber)
      ? Math.max(0, Math.trunc(query.citation.pageNumber) - 1)
      : undefined;
  const limit = Math.max(
    1,
    Math.trunc(query.maxResults ?? DEFAULT_MAX_RESULTS),
  );
  const readOptions = options.signal ? { signal: options.signal } : {};
  const elementQuery: ElementQuery =
    query.pageIndex === undefined ? {} : { pageIndex: query.pageIndex };

  // The two reads queue one after the other; a change landing between them
  // would make the elements stale, so they are read again until they agree.
  let elements: ReadResult<EditElement>;
  let found: ReadResult<TextTarget> | undefined;
  for (let attempt = 0; ; attempt += 1) {
    elements = await session.getElements(elementQuery, readOptions);
    found = needle
      ? await session.findText(needle, {
          ...readOptions,
          ...(query.pageIndex === undefined
            ? {}
            : { pageRange: [query.pageIndex, query.pageIndex] as const }),
        })
      : undefined;
    if (!found || found.revision === elements.revision || attempt >= 2) break;
  }
  const envelope = {
    sessionId: (found ?? elements).sessionId,
    revision: (found ?? elements).revision,
  };
  const allowed = allowedElements(elements.items, query);
  const byId = new Map(allowed.map((element) => [element.id, element]));

  let candidates: TargetCandidate[] = [];
  if (needle && found) {
    candidates = exactCandidates(found.items, byId);
    if (candidates.length === 0) {
      const corpus = buildCorpus(allowed);
      candidates = normalizedCandidates(corpus, needle);
      if (candidates.length === 0) candidates = fuzzyCandidates(corpus, needle);
    }
  } else if (!needle && query.kinds && query.kinds.length > 0) {
    candidates = allowed.map((element) => ({
      elementId: element.id,
      pageIndex: element.pageIndex,
      score: 0.5,
      reason: "kind-only" as const,
      snippet: snippetOf(element.text ?? "", 0, 0),
    }));
  }
  if (hint !== undefined)
    candidates = candidates
      .map((candidate, index) => ({ candidate, index }))
      .sort(
        (a, b) =>
          distance(a.candidate.pageIndex, hint) -
            distance(b.candidate.pageIndex, hint) || a.index - b.index,
      )
      .map((entry) => entry.candidate);
  return Object.freeze({
    ...envelope,
    items: Object.freeze(
      candidates.slice(0, limit).map((candidate) => Object.freeze(candidate)),
    ),
  });
}

function distance(pageIndex: number, hint: number): number {
  return pageIndex < 0 ? Number.POSITIVE_INFINITY : Math.abs(pageIndex - hint);
}

/** The elements a query may name: by kind, inside `within`, on the page. */
function allowedElements(
  elements: readonly EditElement[],
  query: TargetQuery,
): readonly EditElement[] {
  let allowed = elements;
  if (query.within !== undefined) {
    const inside = new Set<string>([query.within]);
    // Children follow their parents in reading order, so one pass closes
    // the set; a second catches a child listed before its parent.
    for (let pass = 0; pass < 2; pass += 1)
      for (const element of elements)
        if (element.parentId !== undefined && inside.has(element.parentId))
          inside.add(element.id);
    allowed = allowed.filter((element) => inside.has(element.id));
  }
  if (query.kinds) {
    const kinds = new Set(query.kinds);
    allowed = allowed.filter((element) => kinds.has(element.kind));
  }
  return allowed;
}

function exactCandidates(
  targets: readonly TextTarget[],
  byId: ReadonlyMap<string, EditElement>,
): TargetCandidate[] {
  const out: TargetCandidate[] = [];
  for (const target of targets) {
    if (target.elementIds.length === 0) continue;
    if (!target.elementIds.every((id) => byId.has(id))) continue;
    const first = target.ranges[0];
    const last = target.ranges.at(-1);
    const element = byId.get(target.elementIds[0]!)!;
    const range: TextRange | undefined =
      first && last ? { start: first.start, end: last.end } : undefined;
    out.push({
      elementId: element.id,
      ...(range && !coversWhole(range, element) ? { range } : {}),
      pageIndex: target.pageIndex,
      score: 1,
      reason: "exact",
      snippet:
        range &&
        element.text !== undefined &&
        range.start.elementId === element.id
          ? snippetOf(
              element.text,
              range.start.offset,
              range.end.elementId === element.id
                ? range.end.offset
                : element.text.length,
            )
          : foldText(target.text),
    });
  }
  return out;
}

function coversWhole(range: TextRange, element: EditElement): boolean {
  return (
    range.start.elementId === element.id &&
    range.end.elementId === element.id &&
    range.start.offset === 0 &&
    range.end.offset === (element.text?.length ?? 0)
  );
}

/** One element's text inside a chunk, by chunk offsets. */
interface Segment {
  readonly element: EditElement;
  readonly start: number;
  readonly end: number;
}

/** Consecutive elements of one page, joined by spaces, within Fuse's text budget. */
interface Chunk {
  readonly pageIndex: number;
  readonly text: string;
  readonly segments: readonly Segment[];
}

/**
 * The elements' text in reading order, cut into chunks that never span a
 * page. A container whose children carry text (a Word table, whose cells are
 * paragraphs) contributes nothing itself, so a match names the cell.
 */
function buildCorpus(elements: readonly EditElement[]): readonly Chunk[] {
  const textChildren = new Set<string>();
  for (const element of elements)
    if (element.parentId !== undefined && element.text)
      textChildren.add(element.parentId);
  const chunks: Chunk[] = [];
  let pageIndex: number | undefined;
  let text = "";
  let segments: Segment[] = [];
  const flush = (): void => {
    if (segments.length > 0 && pageIndex !== undefined)
      chunks.push({ pageIndex, text, segments });
    text = "";
    segments = [];
  };
  const budget = DEFAULT_FUZZY_SEARCH_OPTIONS.maxPageTextLength;
  for (const element of elements) {
    if (!element.text || textChildren.has(element.id)) continue;
    if (
      pageIndex !== element.pageIndex ||
      text.length + 1 + element.text.length > budget
    ) {
      flush();
      pageIndex = element.pageIndex;
    }
    const start = text.length === 0 ? 0 : text.length + 1;
    text = text.length === 0 ? element.text : `${text} ${element.text}`;
    segments.push({ element, start, end: start + element.text.length });
  }
  flush();
  return chunks;
}

/** Whitespace runs to one space, quotes and dashes to ASCII, NFKC, case folded; offsets kept. */
export function foldWithMap(text: string): {
  readonly text: string;
  readonly starts: readonly number[];
  readonly ends: readonly number[];
} {
  const out: string[] = [];
  const starts: number[] = [];
  const ends: number[] = [];
  let spaceStart = -1;
  let spaceEnd = -1;
  for (const segment of graphemeSegments(text)) {
    const value = foldPunctuation(normalizeSearchText(segment.value));
    if (value.length === 0) continue;
    if (/^\s+$/u.test(value)) {
      if (out.length === 0) continue;
      if (spaceStart < 0) spaceStart = segment.start;
      spaceEnd = segment.end;
      continue;
    }
    if (spaceStart >= 0) {
      out.push(" ");
      starts.push(spaceStart);
      ends.push(spaceEnd);
      spaceStart = -1;
    }
    for (const char of value) {
      out.push(char);
      starts.push(segment.start);
      ends.push(segment.end);
    }
  }
  return { text: out.join(""), starts, ends };
}

const PUNCTUATION: Readonly<Record<string, string>> = {
  "‘": "'",
  "’": "'",
  "‚": "'",
  "‛": "'",
  "′": "'",
  "“": '"',
  "”": '"',
  "„": '"',
  "‟": '"',
  "″": '"',
  "«": '"',
  "»": '"',
  "‐": "-",
  "‑": "-",
  "‒": "-",
  "–": "-",
  "—": "-",
  "―": "-",
  "−": "-",
  "…": "...",
  " ": " ",
};

function foldPunctuation(value: string): string {
  let out = "";
  for (const char of value) out += PUNCTUATION[char] ?? char;
  return out;
}

function normalizedCandidates(
  corpus: readonly Chunk[],
  needle: string,
): TargetCandidate[] {
  const pattern = foldWithMap(needle).text;
  if (pattern.length === 0) return [];
  const out: TargetCandidate[] = [];
  for (const chunk of corpus) {
    const folded = foldWithMap(chunk.text);
    let from = 0;
    while (from <= folded.text.length - pattern.length) {
      const at = folded.text.indexOf(pattern, from);
      if (at < 0) break;
      const start = folded.starts[at]!;
      const end = folded.ends[at + pattern.length - 1]!;
      const candidate = candidateOf(chunk, start, end, 0.9, "normalized");
      if (candidate) out.push(candidate);
      from = at + Math.max(1, pattern.length);
    }
  }
  return out;
}

function fuzzyCandidates(
  corpus: readonly Chunk[],
  needle: string,
): TargetCandidate[] {
  if (corpus.length === 0) return [];
  const fuzzy = DEFAULT_FUZZY_SEARCH_OPTIONS;
  const pattern = needle.slice(0, fuzzy.maxQueryLength);
  const pages: FuzzyPageText[] = corpus.map((chunk, index) => ({
    pageIndex: index,
    text: chunk.text,
  }));
  const index = new FuzzyPageIndex(pages, fuzzy, false);
  const out: TargetCandidate[] = [];
  for (const match of index.search(pattern, fuzzy.maxScore)) {
    const chunk = corpus[match.pageIndex]!;
    const passage = alignFuzzyPassage(
      chunk.text,
      pattern,
      false,
      fuzzy.maxScore,
    );
    if (!passage) continue;
    const ratio = passage.length === 0 ? 0 : passage.cost / passage.length;
    const score =
      FUZZY_TOP - (FUZZY_TOP - 0.5) * Math.min(1, ratio / fuzzy.maxScore);
    const candidate = candidateOf(
      chunk,
      passage.start,
      passage.end,
      Math.round(score * 100) / 100,
      "fuzzy",
    );
    if (candidate) out.push(candidate);
  }
  return out;
}

/** The candidate for a span of a chunk: the first element it touches, with a range unless it covers the whole element. */
function candidateOf(
  chunk: Chunk,
  start: number,
  end: number,
  score: number,
  reason: TargetCandidate["reason"],
): TargetCandidate | undefined {
  const touched = chunk.segments.filter(
    (segment) => segment.end > start && segment.start < end,
  );
  const first = touched[0];
  const last = touched.at(-1);
  if (!first || !last) return undefined;
  const startOffset = Math.max(0, start - first.start);
  const endOffset = Math.min(last.end, end) - last.start;
  const range: TextRange = {
    start: { elementId: first.element.id, offset: startOffset },
    end: { elementId: last.element.id, offset: endOffset },
  };
  return {
    elementId: first.element.id,
    ...(coversWhole(range, first.element) ? {} : { range }),
    pageIndex: first.element.pageIndex,
    score,
    reason,
    snippet: snippetOf(chunk.text, start, end),
  };
}

/** The matched text with a little context on each side, on one line. */
export function snippetOf(text: string, start: number, end: number): string {
  const from = Math.max(0, start - SNIPPET_CONTEXT);
  const to = Math.min(text.length, end + SNIPPET_CONTEXT);
  const head = from > 0 ? "…" : "";
  const tail = to < text.length ? "…" : "";
  return `${head}${foldText(text.slice(from, to))}${tail}`;
}
