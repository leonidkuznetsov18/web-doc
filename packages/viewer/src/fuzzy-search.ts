import Fuse, { type FuseIndex, type IFuseOptions } from "fuse.js";

import type { FuzzySearchOptions, SearchMatch } from "./contracts.js";
import { alignFuzzyPassage } from "./fuzzy-alignment.js";

/**
 * Fuse.js selects candidate pages (Bitap with a bounded edit budget per
 * 32-character chunk); contiguous alignment determines highlight boundaries.
 * This bridges the gaps NFKC + case folding leave open when
 * a query was not copied verbatim from the document: AI-generated citations
 * and OCR'd passages differ from the source in spacing, line breaks, list
 * bullets, table separators and typographic punctuation. Every hit maps back
 * to original UTF-16 offsets, so the viewport highlights the verbatim text.
 */
export interface ResolvedFuzzySearchOptions {
  readonly threshold: number;
  readonly maxScore: number;
  readonly maxQueryLength: number;
  readonly maxPageTextLength: number;
  readonly pagesPerBatch: number;
  readonly pageWindow: number;
  readonly worker: boolean;
}

export const DEFAULT_FUZZY_SEARCH_OPTIONS: ResolvedFuzzySearchOptions =
  Object.freeze({
    threshold: 0.3,
    maxScore: 0.4,
    // Bitap cost grows with the query; 600 characters still identifies a
    // passage while bounding both candidate scoring and alignment work.
    maxQueryLength: 600,
    maxPageTextLength: 20_000,
    pagesPerBatch: 2,
    pageWindow: 12,
    worker: true,
  });

/**
 * Layer fuzzy settings in precedence order (viewer defaults first, then the
 * per-call option). `true` enables the defaults, an object enables and
 * overrides them, `false` disables, `undefined` leaves the previous layer in
 * place. Returns `undefined` when fuzzy matching ends up disabled.
 */
export function resolveFuzzySearchOptions(
  ...layers: readonly (boolean | FuzzySearchOptions | undefined)[]
): ResolvedFuzzySearchOptions | undefined {
  let enabled = false;
  let resolved: ResolvedFuzzySearchOptions = DEFAULT_FUZZY_SEARCH_OPTIONS;
  for (const layer of layers) {
    if (layer === undefined) continue;
    if (typeof layer === "boolean") {
      enabled = layer;
      continue;
    }
    enabled = true;
    resolved = {
      threshold: unitInterval(layer.threshold, resolved.threshold),
      maxScore: unitInterval(layer.maxScore, resolved.maxScore),
      maxQueryLength: positiveInteger(
        layer.maxQueryLength,
        resolved.maxQueryLength,
      ),
      maxPageTextLength: positiveInteger(
        layer.maxPageTextLength,
        resolved.maxPageTextLength,
      ),
      pagesPerBatch: positiveInteger(
        layer.pagesPerBatch,
        resolved.pagesPerBatch,
      ),
      pageWindow: positiveInteger(layer.pageWindow, resolved.pageWindow),
      worker: layer.worker ?? resolved.worker,
    };
  }
  return enabled ? resolved : undefined;
}

export interface FuzzyPageText {
  readonly pageIndex: number;
  readonly text: string;
}

/** Fuse.js options shared by every fuzzy scan; `keys` and `threshold` are fixed per index. */
function fuseOptions(
  threshold: number,
  caseSensitive: boolean,
): IFuseOptions<FuzzyPageText> {
  return {
    keys: ["text"],
    isCaseSensitive: caseSensitive,
    ignoreDiacritics: false,
    includeScore: true,
    // A citation can sit anywhere on the page; Fuse's location bias would
    // otherwise penalize matches far from the start of the text.
    ignoreLocation: true,
    // Long page text must not dilute the score of a match inside it.
    ignoreFieldNorm: true,
    threshold,
    minMatchCharLength: 3,
    shouldSort: false,
  };
}

/**
 * A document's pages, indexed once for Fuse.js so that every fuzzy search
 * reuses the normalized records instead of rebuilding them: a citation
 * lookup tries several anchors in a row, and each used to pay for the index
 * again. A scan restricted to a page window reuses the same records through
 * `Fuse.parseIndex`, so only scoring and alignment run for those pages.
 */
export class FuzzyPageIndex {
  readonly #caseSensitive: boolean;
  readonly #pages: readonly FuzzyPageText[];
  readonly #index: FuseIndex<FuzzyPageText>;
  readonly #options: IFuseOptions<FuzzyPageText>;
  readonly #fuse: Fuse<FuzzyPageText>;

  constructor(
    pages: readonly FuzzyPageText[],
    options: { threshold: number; maxPageTextLength: number },
    caseSensitive = false,
  ) {
    this.#caseSensitive = caseSensitive;
    this.#pages = pages.map((page) => ({
      pageIndex: page.pageIndex,
      text: page.text.slice(0, options.maxPageTextLength),
    }));
    this.#options = fuseOptions(options.threshold, caseSensitive);
    this.#index = Fuse.createIndex(["text"], this.#pages);
    this.#fuse = new Fuse(this.#pages, this.#options, this.#index);
  }

  get pageCount(): number {
    return this.#pages.length;
  }

  /**
   * One match per page whose text holds the query within the edit budget:
   * the best contiguous alignment, so neighbouring text and repeated
   * occurrences cannot expand the highlight. `pageIndices` restricts the scan;
   * the result is in page order either way.
   */
  search(
    query: string,
    maxScore: number,
    pageIndices?: readonly number[],
  ): readonly SearchMatch[] {
    if (!query.trim()) return [];
    const fuse = pageIndices ? this.#subset(pageIndices) : this.#fuse;
    const matches: SearchMatch[] = [];
    for (const result of fuse.search(query)) {
      if ((result.score ?? 1) > maxScore) continue;
      const span = alignFuzzyPassage(
        result.item.text,
        query,
        this.#caseSensitive,
        maxScore,
      );
      if (!span) continue;
      const { start, end } = span;
      const text = result.item.text.slice(start, end);
      matches.push({ pageIndex: result.item.pageIndex, start, end, text });
    }
    return matches.sort((a, b) => a.pageIndex - b.pageIndex);
  }

  #subset(pageIndices: readonly number[]): Fuse<FuzzyPageText> {
    const wanted = new Set(pageIndices);
    const { keys, records } = this.#index.toJSON();
    const kept = records.filter((record) =>
      wanted.has(this.#pages[record.i]?.pageIndex ?? -1),
    );
    return new Fuse(
      this.#pages,
      this.#options,
      Fuse.parseIndex<FuzzyPageText>({ keys, records: kept }),
    );
  }
}

/**
 * One-off scan of a few pages, for callers without a standing index. The
 * viewer keeps a {@link FuzzyPageIndex} per document instead.
 */
export function findFuzzyPageMatches(
  pages: readonly FuzzyPageText[],
  query: string,
  options: ResolvedFuzzySearchOptions,
  caseSensitive = false,
): readonly SearchMatch[] {
  return new FuzzyPageIndex(pages, options, caseSensitive).search(
    query.slice(0, options.maxQueryLength),
    options.maxScore,
  );
}

/** Page indices of `[first, last]` ordered by distance from `nearPage`, ties earlier-first. */
export function pagesNearestFirst(
  first: number,
  last: number,
  nearPage: number | undefined,
): number[] {
  const pages = Array.from({ length: last - first + 1 }, (_, i) => first + i);
  if (nearPage === undefined) return pages;
  return pages.sort(
    (a, b) => Math.abs(a - nearPage) - Math.abs(b - nearPage) || a - b,
  );
}

/** Index of the match closest to `nearPage`; the first match when there is no hint. */
export function nearestMatchIndex(
  matches: readonly SearchMatch[],
  nearPage: number | undefined,
): number {
  if (matches.length === 0) return -1;
  if (nearPage === undefined) return 0;
  let best = 0;
  for (let index = 1; index < matches.length; index += 1)
    if (
      Math.abs(matches[index]!.pageIndex - nearPage) <
      Math.abs(matches[best]!.pageIndex - nearPage)
    )
      best = index;
  return best;
}

function unitInterval(value: number | undefined, fallback: number): number {
  return value !== undefined &&
    Number.isFinite(value) &&
    value >= 0 &&
    value <= 1
    ? value
    : fallback;
}

function positiveInteger(value: number | undefined, fallback: number): number {
  return value !== undefined && Number.isInteger(value) && value > 0
    ? value
    : fallback;
}
