import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { foldWithMap, snippetOf } from "../src/edit/ai/targets.js";
import { EditSessionController } from "../src/edit/session.js";
import type { EditOperation, TargetCandidate } from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";
import { docxSession, run } from "./fixtures/docx-session.js";
import {
  encodePages,
  FakeEditEngine,
  FakeHost,
} from "./fixtures/fake-edit-engine.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";
import { buildDeck, table, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

/*
 * Task 60 of the ai-edit module: `resolveTargets()` turns a quoted text, a
 * citation or a kind into element ids and ranges in named passes (exact,
 * normalized, fuzzy, kind-only), stopping at the first that yields, with
 * scores a host can act on.
 */

function brief(candidate: TargetCandidate): unknown[] {
  return [
    candidate.reason,
    candidate.elementId,
    candidate.range
      ? `${candidate.range.start.elementId}@${candidate.range.start.offset}-${candidate.range.end.elementId}@${candidate.range.end.offset}`
      : "",
    candidate.score,
  ];
}

describe("foldWithMap", () => {
  it("folds whitespace, case, quotes, dashes and compatibility forms, keeping offsets", () => {
    const text = "  The “Quick”\n\tbrown—fox ﬁle ";
    const folded = foldWithMap(text);
    assert.equal(folded.text, 'the "quick" brown-fox file');
    // Every folded character maps back into the original text.
    const at = folded.text.indexOf("brown");
    assert.equal(text.slice(folded.starts[at]!, folded.ends[at + 4]!), "brown");
    const quote = folded.text.indexOf('"');
    assert.equal(text[folded.starts[quote]!], "“");
    assert.equal(foldWithMap("   ").text, "");
  });
});

describe("snippetOf", () => {
  it("adds a little context and marks the cuts", () => {
    const text = `${"a".repeat(50)}MATCH${"b".repeat(50)}`;
    assert.equal(
      snippetOf(text, 50, 55),
      `…${"a".repeat(40)}MATCH${"b".repeat(40)}…`,
    );
    assert.equal(snippetOf("short\ttext", 0, 5), "short | text");
  });
});

describe("resolveTargets on the core", () => {
  it("finds exact text through the engine with the read envelope", async () => {
    const original = encodePages(["one two", "three four"]);
    const engine = new FakeEditEngine(original);
    const core = new EditSessionController(engine, new FakeHost(), original, 2);
    try {
      const exact = await core.resolveTargets({ text: "three" });
      assert.equal(exact.sessionId, core.sessionId);
      assert.equal(exact.revision, 0);
      assert.deepEqual(exact.items.map(brief), [["exact", "p1w0", "", 1]]);
      assert.equal(exact.items[0]!.pageIndex, 1);
      assert.equal(exact.items[0]!.snippet, "three");
      assert.equal(Object.isFrozen(exact.items[0]), true);
      // A kind the query excludes drops the engine's match.
      const none = await core.resolveTargets({
        text: "three",
        kinds: ["image"],
      });
      assert.deepEqual(none.items, []);
      // Without text, kinds list elements in reading order with a low score.
      const kinds = await core.resolveTargets({
        kinds: ["word"],
        maxResults: 3,
      });
      assert.deepEqual(kinds.items.map(brief), [
        ["kind-only", "p0w0", "", 0.5],
        ["kind-only", "p0w1", "", 0.5],
        ["kind-only", "p1w0", "", 0.5],
      ]);
      assert.deepEqual((await core.resolveTargets({})).items, []);
      // A page bounds every pass.
      const paged = await core.resolveTargets({
        kinds: ["word"],
        pageIndex: 1,
      });
      assert.deepEqual(
        paged.items.map((item) => item.elementId),
        ["p1w0", "p1w1"],
      );
      // A read queued behind a change describes the state after it.
      const shout = { op: "setText", pageIndex: 0, text: "ONE two" };
      const pending = core.apply([shout as EditOperation]);
      const after = core.resolveTargets({ text: "ONE" });
      await pending;
      assert.equal((await after).revision, 1);
    } finally {
      await core.end();
    }
  });

  it("falls back to folded and fuzzy matches over the elements' text", async () => {
    const original = encodePages([
      "The “quarterly” review covers three regions and two quarters",
    ]);
    const engine = new FakeEditEngine(original);
    const core = new EditSessionController(engine, new FakeHost(), original, 1);
    try {
      // The fake engine only matches a whole-page substring; a quoted word
      // with straight quotes fails there and is found folded.
      const folded = await core.resolveTargets({ text: '"QUARTERLY" review' });
      assert.equal(folded.items.length, 1);
      const [first] = folded.items;
      assert.equal(first!.reason, "normalized");
      assert.equal(first!.score, 0.9);
      assert.equal(first!.elementId, "p0w1");
      assert.deepEqual(first!.range, {
        start: { elementId: "p0w1", offset: 0 },
        end: { elementId: "p0w2", offset: 6 },
      });
      assert.equal(
        first!.snippet,
        "The “quarterly” review covers three regions and two quarters",
      );
      // A citation with a typo is found fuzzily, scored below a folded match.
      const fuzzy = await core.resolveTargets({
        citation: { text: "covers thre regions", pageNumber: 1 },
      });
      assert.equal(fuzzy.items.length, 1);
      assert.equal(fuzzy.items[0]!.reason, "fuzzy");
      assert.ok(fuzzy.items[0]!.score >= 0.5 && fuzzy.items[0]!.score < 0.9);
      assert.equal(fuzzy.items[0]!.elementId, "p0w3");
      assert.equal(fuzzy.items[0]!.range?.end.elementId, "p0w5");
      // Nothing like it: no candidates, no exception.
      const nothing = await core.resolveTargets({
        text: "completely unrelated words here",
      });
      assert.deepEqual(nothing.items, []);
    } finally {
      await core.end();
    }
  });
});

describe("resolveTargets on format sessions", () => {
  it("resolves inside PDF text boxes with ranges and whole-element matches", async () => {
    const { session, end } = await pdfSession(
      await buildPdf(["First page", "Second page"]),
    );
    try {
      await session.insertTextBox({
        pageIndex: 1,
        rect: { x: 72, y: 100, width: 300, height: 60 },
        text: "Quarterly review of the northern region",
      });
      const part = await session.resolveTargets({ text: "northern region" });
      assert.equal(part.items.length, 1);
      const [candidate] = part.items;
      assert.equal(candidate!.reason, "exact");
      assert.equal(candidate!.pageIndex, 1);
      assert.equal(candidate!.range?.start.offset, 24);
      assert.equal(candidate!.range?.end.offset, 39);
      const box = await session.getElement(candidate!.elementId);
      assert.equal(box.item?.kind, "textBox");
      // The whole text of a text object: no range, so the operation takes the element.
      const whole = await session.resolveTargets({ text: "First page" });
      assert.equal(whole.items[0]!.reason, "exact");
      assert.equal(whole.items[0]!.range, undefined);
      // A folded query with a line break inside still names the box.
      const folded = await session.resolveTargets({
        text: "quarterly\nREVIEW",
      });
      assert.equal(folded.items[0]?.reason, "normalized");
      assert.equal(folded.items[0]?.elementId, candidate!.elementId);
      // The citation hint orders candidates by page distance.
      await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 100, width: 300, height: 60 },
        text: "Quarterly review of the southern region",
      });
      const near = await session.resolveTargets({
        citation: { text: "Quarterly review", pageNumber: 2 },
      });
      assert.deepEqual(
        near.items.map((item) => item.pageIndex),
        [1, 0],
      );
      const far = await session.resolveTargets({
        citation: { text: "Quarterly review", pageNumber: 1 },
      });
      assert.deepEqual(
        far.items.map((item) => item.pageIndex),
        [0, 1],
      );
    } finally {
      await end();
    }
  });

  it("resolves deck shapes by kind, within a slide and by table text", async () => {
    const deck = buildDeck({
      slides: [
        {
          shapes: [
            textShape({
              id: 2,
              name: "Title 1",
              x: 0,
              y: 0,
              cx: 914400,
              cy: 457200,
              paragraphs: [["Agenda"]],
            }),
            table({
              id: 3,
              name: "Table 2",
              x: 0,
              y: 914400,
              cx: 1828800,
              cy: 914400,
              rows: [
                ["Region", "Q1"],
                ["North", "120"],
              ],
            }),
          ],
        },
        {
          shapes: [
            textShape({
              id: 2,
              name: "Title 1",
              x: 0,
              y: 0,
              cx: 914400,
              cy: 457200,
              paragraphs: [["Results"]],
            }),
            table({
              id: 3,
              name: "Table 2",
              x: 0,
              y: 914400,
              cx: 1828800,
              cy: 914400,
              rows: [["Region", "Q2"]],
            }),
          ],
        },
      ],
    });
    const { session, end } = await pptxSession(deck);
    try {
      const tables = await session.resolveTargets({
        kinds: ["table"],
        pageIndex: 1,
      });
      assert.deepEqual(tables.items.map(brief), [
        ["kind-only", "sld2:3", "", 0.5],
      ]);
      assert.equal(tables.items[0]!.snippet, "Region | Q2");
      const cell = await session.resolveTargets({ text: "North" });
      assert.deepEqual(cell.items.map(brief), [
        ["exact", "sld1:3", "sld1:3@10-sld1:3@15", 1],
      ]);
      const title = await session.resolveTargets({
        text: "results",
        kinds: ["shape"],
      });
      assert.deepEqual(title.items.map(brief), [["exact", "sld2:2", "", 1]]);
      assert.equal(title.items[0]!.pageIndex, 1);
    } finally {
      await end();
    }
  });

  it("resolves Word paragraphs, cells within a table and unplaced pages", async () => {
    const bytes = buildDocx({
      body:
        paragraph("Intro paragraph about the review") +
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
        `<w:tr><w:tc>${paragraph("Region")}</w:tc><w:tc>${paragraph("Review score")}</w:tc></w:tr></w:tbl>` +
        paragraph("Closing “review” remarks") +
        sectPr(),
    });
    const { session, end } = await docxSession(bytes, [[]]);
    try {
      const outline = await session.getOutline();
      const [intro, tbl, closing] = outline.items.map((node) => node.id);
      const cellId = outline.items[1]!.children![1]!.id;
      const all = await session.resolveTargets({ text: "review" });
      assert.deepEqual(
        all.items.map((item) => [item.reason, item.elementId]),
        [
          ["exact", intro],
          ["exact", cellId],
          ["exact", closing],
        ],
      );
      assert.equal(all.items[0]!.pageIndex, -1);
      // Inside the table only: the cell paragraph, never the table's own text.
      const inside = await session.resolveTargets({
        text: "review",
        within: tbl!,
      });
      assert.deepEqual(
        inside.items.map((item) => item.elementId),
        [cellId],
      );
      // Folded quotes find the closing paragraph through the second pass.
      const quoted = await session.resolveTargets({ text: '"review" remarks' });
      assert.deepEqual(quoted.items.map(brief), [
        ["normalized", closing!, `${closing}@8-${closing}@24`, 0.9],
      ]);
      // A placed page gives the candidate its page.
      const placed = await docxSession(bytes, [
        [run(intro!.slice(2), "Intro", 96, 96)],
      ]);
      try {
        const found = await placed.session.resolveTargets({
          text: "Intro paragraph",
          pageIndex: 0,
        });
        assert.equal(found.items[0]?.pageIndex, 0);
        assert.equal(found.items[0]?.elementId, intro);
      } finally {
        await placed.end();
      }
    } finally {
      await end();
    }
  });
});
