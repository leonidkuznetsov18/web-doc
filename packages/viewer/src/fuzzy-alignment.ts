import { normalizeSearchText, normalizeWithMap } from "./search-text.js";

/**
 * Fuse's indices describe character masks, not a contiguous edit alignment.
 * Refine an accepted page with semi-global Levenshtein alignment: consume the
 * whole query, allowing free text before and after one occurrence. Tracking
 * the start alongside each cost needs O(query length) working memory rather
 * than a page × query traceback matrix. Equal-cost occurrences prefer the
 * earliest end, so an unmatched character after a passage cannot extend it.
 */
export function alignFuzzyPassage(
  text: string,
  query: string,
  caseSensitive: boolean,
  maxScore: number,
): { start: number; end: number } | undefined {
  const source = normalizeWithMap(text, caseSensitive);
  const pattern = normalizeSearchText(query, caseSensitive);
  const length = pattern.length;
  if (!length || !source.text.length) return undefined;

  const costs = new Uint32Array(length + 1);
  const starts = new Uint32Array(length + 1);
  for (let i = 0; i <= length; i += 1) costs[i] = i;
  let bestCost = length;
  let bestStart = 0;
  let bestEnd = 0;

  for (let end = 1; end <= source.text.length; end += 1) {
    let diagonalCost = costs[0]!;
    let diagonalStart = starts[0]!;
    costs[0] = 0;
    starts[0] = end;
    for (let i = 1; i <= length; i += 1) {
      const previousCost = costs[i]!;
      const previousStart = starts[i]!;
      let cost =
        diagonalCost + (pattern[i - 1] === source.text[end - 1] ? 0 : 1);
      let start = diagonalStart;
      // Prefer dropping an unmatched query character over substituting a
      // neighbouring source character when both alignments cost the same.
      const deletion = costs[i - 1]! + 1;
      if (deletion <= cost) {
        cost = deletion;
        start = starts[i - 1]!;
      }
      const insertion = previousCost + 1;
      if (insertion < cost) {
        cost = insertion;
        start = previousStart;
      }
      costs[i] = cost;
      starts[i] = start;
      diagonalCost = previousCost;
      diagonalStart = previousStart;
    }
    if (costs[length]! < bestCost) {
      bestCost = costs[length]!;
      bestStart = starts[length]!;
      bestEnd = end;
    }
  }

  // A page-level Fuse score alone can accept disconnected 32-character
  // chunks. Require the single passage to meet the same score ceiling.
  if (bestEnd <= bestStart || bestCost / length > maxScore) return undefined;
  const start = source.starts[bestStart];
  const end = source.ends[bestEnd - 1];
  return start === undefined || end === undefined ? undefined : { start, end };
}
