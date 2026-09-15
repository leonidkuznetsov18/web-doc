import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  DEFAULT_FUZZY_SEARCH_OPTIONS,
  findFuzzyPageMatches,
  handleFuzzyWorkerRequest,
  type FuzzyWorkerState,
} from "../src/index.js";

// Synthetic content: no customer document is needed to reproduce the leak.
const passage =
  "The shift from manual review to automated checks is the difference between incremental improvement and reliable delivery.";
const prefix = "Earlier paragraph.";
const suffix = "The next chapter";

function search(text: string, query: string) {
  return findFuzzyPageMatches(
    [{ pageIndex: 0, text }],
    query,
    DEFAULT_FUZZY_SEARCH_OPTIONS,
  );
}

describe("citation highlight boundaries", () => {
  for (const [label, query] of [
    ["extra whitespace", passage.replace("and reliable", "and  reliable")],
    ["line break", passage.replace("and reliable", "and\nreliable")],
    ["typo", passage.replace("automated", "automatd")],
  ] as const) {
    it(`excludes the next paragraph with ${label}`, () => {
      assert.deepEqual(search(prefix + passage + suffix, query), [
        {
          pageIndex: 0,
          start: prefix.length,
          end: prefix.length + passage.length,
          text: passage,
        },
      ]);
    });
  }

  it("chooses one occurrence instead of spanning repeated passages", () => {
    assert.deepEqual(
      search(
        prefix + passage + suffix + passage,
        passage.replace("automated", "automatd"),
      ),
      [
        {
          pageIndex: 0,
          start: prefix.length,
          end: prefix.length + passage.length,
          text: passage,
        },
      ],
    );
  });

  it("rejects disconnected query chunks separated by unrelated content", () => {
    assert.deepEqual(
      search(
        "Red birds fly over the mountain. An unrelated paragraph about shipping inventory. Blue boats sail across the ocean. Another unrelated inventory paragraph. Green trains cross the open desert.",
        "Red birds fly over the mountain. Blue boats sail across the ocean. Green trains cross the open desert.",
      ),
      [],
    );
  });

  it("preserves original UTF-16 offsets with case-fold expansion and astral characters", () => {
    const source = "İstanbul 🧭 " + passage;
    const before = "🧩 Intro.";
    const query = source
      .toLocaleLowerCase("und")
      .replace("automated", "automatd");
    assert.deepEqual(search(before + source + suffix, query), [
      {
        pageIndex: 0,
        start: before.length,
        end: before.length + source.length,
        text: source,
      },
    ]);
  });

  it("returns the same bounded passage through the worker protocol", () => {
    const state: FuzzyWorkerState = {};
    handleFuzzyWorkerRequest(state, {
      kind: "index",
      id: 1,
      pages: [{ pageIndex: 0, text: prefix + passage + suffix }],
      threshold: 0.3,
      maxPageTextLength: 20_000,
      caseSensitive: false,
    });
    const result = handleFuzzyWorkerRequest(state, {
      kind: "search",
      id: 2,
      query: passage.replace("and reliable", "and  reliable"),
      maxScore: 0.4,
    });
    assert.equal(result.kind, "matches");
    if (result.kind === "matches")
      assert.deepEqual(
        result.matches,
        search(prefix + passage + suffix, passage),
      );
  });
});
