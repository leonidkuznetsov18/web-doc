import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  buildOutline,
  foldText,
  labelOf,
  renderDescription,
} from "../src/edit/ai/outline.js";
import { EditSessionController } from "../src/edit/session.js";
import {
  defaultResourceLimits,
  resolveLimits,
  ViewerError,
  type EditElement,
  type EditOperation,
  type OutlineNode,
} from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";
import { docxSession, run } from "./fixtures/docx-session.js";
import {
  encodePages,
  FakeEditEngine,
  FakeHost,
} from "./fixtures/fake-edit-engine.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";
import { buildDeck, group, table, textShape } from "./fixtures/pptx-builder.js";
import { pptxSession } from "./fixtures/pptx-session.js";

/*
 * Task 59 of the ai-edit module: `getOutline()` nests elements by parent,
 * numbers them in reading order, labels them and cuts their text;
 * `describe()` renders the outline in the fixed line grammar within a
 * character budget. Both are reads with the usual envelope.
 */

const RECT = { x: 0, y: 0, width: 10, height: 10 };

function element(
  id: string,
  kind: string,
  fields: Partial<EditElement> & Record<string, unknown> = {},
): EditElement {
  return {
    id,
    kind,
    pageIndex: 0,
    bounds: RECT,
    operations: ["deleteElement"],
    ...fields,
  };
}

const ELEMENTS: readonly EditElement[] = [
  element("p:1", "paragraph", { text: "Intro" }),
  element("tbl:2", "table", {
    text: "A1\tB1\nA2\tB2",
    table: {
      rows: [
        ["A1", "B1"],
        ["A2", "B2"],
      ],
    },
  }),
  element("p:3", "paragraph", { text: "A1", parentId: "tbl:2" }),
  element("p:4", "paragraph", { text: "B1", parentId: "tbl:2" }),
  element("p:5", "paragraph", { text: "A2", parentId: "tbl:2", pageIndex: 1 }),
  element("p:6", "paragraph", { text: "B2", parentId: "tbl:2", pageIndex: 1 }),
  element("p:7", "paragraph", { text: "Picture here", pageIndex: 1 }),
  element("img:8", "image", { parentId: "p:7", pageIndex: 1 }),
  element("p:9", "paragraph", { text: "", pageIndex: 2 }),
];

function flat(nodes: readonly OutlineNode[]): string[] {
  return nodes.flatMap((node) => [
    `${node.ordinal} ${node.id}`,
    ...flat(node.children ?? []),
  ]);
}

describe("buildOutline", () => {
  it("nests by parent and numbers nodes in reading order", () => {
    const built = buildOutline("docx", ELEMENTS, {}, 5000);
    assert.deepEqual(flat(built.nodes), [
      "1 p:1",
      "2 tbl:2",
      "2.1 p:3",
      "2.2 p:4",
      "2.3 p:5",
      "2.4 p:6",
      "3 p:7",
      "3.1 img:8",
      "4 p:9",
    ]);
    assert.equal(built.nodeCount, 9);
    assert.equal(built.truncated, false);
    const tableNode = built.nodes[1]!;
    assert.equal(tableNode.label, "Table (2×2)");
    assert.equal(tableNode.text, "A1\tB1\nA2\tB2");
    assert.equal(tableNode.textLength, 11);
    assert.deepEqual(tableNode.operations, ["deleteElement"]);
    // No geometry or style leaks into a node.
    assert.deepEqual(Object.keys(tableNode).sort(), [
      "children",
      "id",
      "kind",
      "label",
      "operations",
      "ordinal",
      "pageIndex",
      "text",
      "textLength",
      "truncated",
    ]);
    assert.equal(Object.isFrozen(tableNode), true);
    const empty = built.nodes[3]!;
    assert.equal(empty.text, "");
    assert.equal(empty.textLength, 0);
    assert.equal("children" in built.nodes[0]!, false);
  });

  it("keeps ordinals stable under filters and keeps the containers of kept nodes", () => {
    const byKind = buildOutline("docx", ELEMENTS, { kinds: ["image"] }, 5000);
    assert.deepEqual(flat(byKind.nodes), ["3 p:7", "3.1 img:8"]);
    assert.equal(byKind.nodeCount, 2);
    const byPage = buildOutline("docx", ELEMENTS, { pageRange: [1, 1] }, 5000);
    assert.deepEqual(flat(byPage.nodes), [
      "2 tbl:2",
      "2.3 p:5",
      "2.4 p:6",
      "3 p:7",
      "3.1 img:8",
    ]);
    const both = buildOutline(
      "docx",
      ELEMENTS,
      { pageRange: [1, 2], kinds: ["paragraph"] },
      5000,
    );
    assert.deepEqual(flat(both.nodes), [
      "2 tbl:2",
      "2.3 p:5",
      "2.4 p:6",
      "3 p:7",
      "4 p:9",
    ]);
    assert.deepEqual(
      buildOutline("docx", ELEMENTS, { kinds: ["shape"] }, 5000).nodes,
      [],
    );
  });

  it("cuts text per node without splitting a surrogate pair", () => {
    const long = element("p:1", "paragraph", { text: "ab😀cd" });
    const cut = buildOutline("docx", [long], { maxTextChars: 3 }, 5000)
      .nodes[0]!;
    assert.equal(cut.text, "ab");
    assert.equal(cut.textLength, 6);
    assert.equal(cut.truncated, true);
    const whole = buildOutline("docx", [long], { maxTextChars: 6 }, 5000)
      .nodes[0]!;
    assert.equal(whole.text, "ab😀cd");
    assert.equal(whole.truncated, false);
    const none = buildOutline("docx", [long], { maxTextChars: 0 }, 5000)
      .nodes[0]!;
    assert.equal(none.text, "");
    assert.equal(none.truncated, true);
  });

  it("cuts the outline at the node limit in pre-order", () => {
    const built = buildOutline("docx", ELEMENTS, {}, 4);
    assert.deepEqual(flat(built.nodes), [
      "1 p:1",
      "2 tbl:2",
      "2.1 p:3",
      "2.2 p:4",
    ]);
    assert.equal(built.nodeCount, 4);
    assert.equal(built.truncated, true);
    const exact = buildOutline("docx", ELEMENTS, {}, 9);
    assert.equal(exact.truncated, false);
  });

  it("lists elements whose parents form a cycle after the roots", () => {
    const cyclic = [
      element("a", "shape", { parentId: "b" }),
      element("b", "shape", { parentId: "a" }),
      element("c", "shape"),
    ];
    assert.deepEqual(flat(buildOutline("pptx", cyclic, {}, 5000).nodes), [
      "1 c",
      "2 a",
      "2.1 b",
    ]);
  });

  it("carries the read-only reason and the hidden flag", () => {
    const nodes = buildOutline(
      "docx",
      [
        element("p:1", "paragraph", {
          text: "locked",
          readOnlyReason: "tracked-changes",
          operations: ["insertParagraph"],
        }),
        element("sld1:2", "shape", { hidden: true }),
      ],
      {},
      5000,
    ).nodes;
    assert.equal(nodes[0]!.readOnlyReason, "tracked-changes");
    assert.equal("hidden" in nodes[0]!, false);
    assert.equal(nodes[1]!.hidden, true);
    assert.equal("readOnlyReason" in nodes[1]!, false);
  });
});

describe("labelOf", () => {
  it("names deck shapes by their name, else their placeholder", () => {
    assert.equal(
      labelOf("pptx", element("sld1:2", "shape", { name: "Title 1" })),
      "Title 1",
    );
    assert.equal(
      labelOf(
        "pptx",
        element("sld1:2", "shape", {
          name: "",
          placeholder: { type: "ctrTitle" },
        }),
      ),
      "Title",
    );
    assert.equal(
      labelOf(
        "pptx",
        element("sld1:2", "shape", { name: " ", placeholder: { type: "x" } }),
      ),
      "x placeholder",
    );
    assert.equal(
      labelOf(
        "pptx",
        element("sld1:5", "table", {
          name: "Table 4",
          table: { rows: [["a", "b", "c"]] },
        }),
      ),
      "Table 4 (1×3)",
    );
    assert.equal(
      labelOf("pptx", element("sld1:3", "image", { name: "" })),
      undefined,
    );
  });

  it("names Word paragraphs by their style and tables by their size", () => {
    const styled = (styleId?: string, numbering?: { level: number }) =>
      element("p:1", "paragraph", {
        paragraphStyle: {
          ...(styleId ? { styleId } : {}),
          align: "left",
          spacing: {},
          ...(numbering ? { numbering: { numId: 1, ...numbering } } : {}),
        },
      });
    assert.equal(labelOf("docx", styled("Heading1")), "Heading 1");
    assert.equal(labelOf("docx", styled("ListParagraph")), "List Paragraph");
    assert.equal(labelOf("docx", styled("TOC2")), "TOC 2");
    assert.equal(labelOf("docx", styled("Title")), "Title");
    assert.equal(labelOf("docx", styled("Normal")), undefined);
    assert.equal(labelOf("docx", styled()), undefined);
    assert.equal(
      labelOf("docx", styled("Normal", { level: 1 })),
      "List item (level 2)",
    );
    assert.equal(
      labelOf(
        "docx",
        element("tbl:1", "table", { table: { rows: [["a"], ["b", "c"]] } }),
      ),
      "Table (2×2)",
    );
    assert.equal(labelOf("docx", element("img:1", "image")), undefined);
  });

  it("names PDF tables only", () => {
    assert.equal(
      labelOf("pdf", element("p0:t1", "table", { table: { rows: [["a"]] } })),
      "Table (1×1)",
    );
    assert.equal(labelOf("pdf", element("p0:o1", "textBox")), undefined);
    assert.equal(labelOf("pdf", element("p0:o2", "text")), undefined);
  });
});

describe("renderDescription", () => {
  it("prints the header and one line per node in the grammar", () => {
    const docx = [
      element("p:1", "paragraph", {
        text: "Heading",
        paragraphStyle: { styleId: "Heading1", align: "left", spacing: {} },
      }),
      element("tbl:2", "table", {
        text: "A1\tB1\nA2\tB2",
        table: {
          rows: [
            ["A1", "B1"],
            ["A2", "B2"],
          ],
        },
      }),
      element("p:3", "paragraph", { text: "A1", parentId: "tbl:2" }),
      element("p:4", "paragraph", {
        text: "locked",
        parentId: "tbl:2",
        readOnlyReason: "tracked-changes",
      }),
      element("p:5", "paragraph", { text: "", pageIndex: -1 }),
      element("sld:6", "shape", {
        text: 'He said "hi"',
        name: 'Say "hi"',
        hidden: true,
      }),
    ];
    const built = buildOutline("docx", docx, { maxTextChars: 5 }, 5000);
    const description = renderDescription("docx", 3, built, 50_000, 5000);
    assert.equal(
      description.text,
      [
        "docx: 3 pages, 6 elements",
        '[p:1] page 1 paragraph "Heading 1": Headi…',
        '[tbl:2] page 1 table "Table (2×2)": A1 | B1…',
        "  [p:3] page 1 paragraph: A1",
        "  [p:4] page 1 paragraph (read-only: tracked-changes): locke…",
        "[p:5] paragraph",
        "[sld:6] page 1 shape (hidden): He sa…",
      ].join("\n"),
    );
    assert.equal(description.format, "docx");
    assert.equal(description.pageCount, 3);
    assert.equal(description.elementCount, 6);
    assert.equal(description.truncated, false);
    const deck = renderDescription(
      "pptx",
      1,
      buildOutline(
        "pptx",
        [element("sld1:2", "shape", { name: 'A "B"', text: "x" })],
        {},
        5000,
      ),
      50_000,
      5000,
    );
    assert.equal(
      deck.text,
      "pptx: 1 slide, 1 element\n[sld1:2] slide 1 shape \"A 'B'\": x",
    );
  });

  it("cuts whole lines to the budget and says how many were left", () => {
    const built = buildOutline("docx", ELEMENTS, {}, 5000);
    const full = renderDescription("docx", 3, built, 50_000, 5000);
    const lines = full.text.split("\n");
    const budget = lines.slice(0, 4).join("\n").length + 2;
    const cut = renderDescription("docx", 3, built, budget, 5000);
    assert.ok(cut.text.length <= budget, `${cut.text.length} > ${budget}`);
    const cutLines = cut.text.split("\n");
    assert.equal(cutLines.at(-1), "… (7 more lines)");
    assert.deepEqual(cutLines.slice(0, -1), lines.slice(0, 3));
    assert.equal(cut.truncated, true);
    assert.equal(cut.elementCount, 9);
    // A budget too small for anything hands back the header's head.
    const tiny = renderDescription("docx", 3, built, 8, 5000);
    assert.equal(tiny.text, "docx: 3 ");
    assert.equal(tiny.truncated, true);
  });

  it("says when the node limit cut the outline", () => {
    const built = buildOutline("docx", ELEMENTS, {}, 2);
    const description = renderDescription("docx", 3, built, 50_000, 2);
    assert.equal(
      description.text.split("\n").at(-1),
      "… (outline cut at 2 nodes)",
    );
    assert.equal(description.truncated, true);
    assert.equal(description.elementCount, 2);
  });

  it("folds text to one line", () => {
    assert.equal(foldText("a\tb\r\nc\nd e\u0007f"), "a | b ⏎ c ⏎ d ⏎ e f");
  });
});

describe("outline reads on sessions", () => {
  it("reads through the core with the envelope of the state described", async () => {
    const original = encodePages(["one two", "three"]);
    const engine = new FakeEditEngine(original);
    const host = new FakeHost();
    const core = new EditSessionController(engine, host, original, 2);
    try {
      const outline = await core.getOutline();
      assert.equal(outline.sessionId, core.sessionId);
      assert.equal(outline.revision, 0);
      assert.deepEqual(
        outline.items.map((node) => [node.ordinal, node.id, node.text]),
        [
          ["1", "p0w0", "one"],
          ["2", "p0w1", "two"],
          ["3", "p1w0", "three"],
        ],
      );
      assert.equal(outline.nodeCount, 3);
      // A read queued behind a change describes the document after it.
      const shout = { op: "setText", pageIndex: 1, text: "THREE" };
      const pending = core.apply([shout as EditOperation]);
      const after = core.describe({ maxChars: 1000 });
      await pending;
      const description = await after;
      assert.equal(description.revision, 1);
      assert.equal(
        description.item?.text,
        "pdf: 2 pages, 3 elements\n[p0w0] page 1 word: one\n[p0w1] page 1 word: two\n[p1w0] page 2 word: THREE",
      );
      // A single-page range is read as that page.
      engine.calls.length = 0;
      const page = await core.getOutline({ pageRange: [1, 1] });
      assert.deepEqual(
        page.items.map((node) => node.id),
        ["p1w0"],
      );
      assert.deepEqual(engine.calls, ["getElements"]);
      // The limits bound the outline and the description.
      const limited = new EditSessionController(
        new FakeEditEngine(original),
        new FakeHost({
          limits: resolveLimits(
            {},
            { maxOutlineNodes: 2, maxDescribeChars: 40 },
          ),
        }),
        original,
        2,
      );
      const small = await limited.getOutline();
      assert.equal(small.truncated, true);
      assert.equal(small.nodeCount, 2);
      const short = await limited.describe({ maxChars: 10_000 });
      assert.ok(short.item!.text.length <= 40);
      assert.equal(short.item!.truncated, true);
      await limited.end();
      // The signal cancels the read like any other.
      const controller = new AbortController();
      controller.abort();
      await assert.rejects(
        core.describe({ signal: controller.signal }),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "aborted",
      );
    } finally {
      await core.end();
    }
  });

  it("outlines a PDF with its text boxes and tables", async () => {
    const { session, end } = await pdfSession(
      await buildPdf([
        { texts: [{ text: "Hello", x: 72, y: 700, fontSize: 24 }] },
        "Second",
      ]),
    );
    try {
      await session.insertTextBox({
        pageIndex: 0,
        rect: { x: 72, y: 100, width: 200, height: 40 },
        text: "A box of text",
      });
      await session.insertTable({
        pageIndex: 1,
        at: { x: 72, y: 100 },
        width: 300,
        rows: [
          ["Region", "Q1"],
          ["North", "120"],
        ],
      });
      const outline = await session.getOutline();
      assert.equal(outline.revision, 2);
      assert.deepEqual(
        outline.items.map((node) => [
          node.ordinal,
          node.kind,
          node.pageIndex,
          node.label ?? "",
          node.text,
        ]),
        [
          ["1", "text", 0, "", "Hello"],
          ["2", "textBox", 0, "", "A box of text"],
          ["3", "text", 1, "", "Second"],
          ["4", "table", 1, "Table (2×2)", "Region\tQ1\nNorth\t120"],
        ],
      );
      assert.ok(outline.items[1]!.operations.includes("replaceText"));
      const description = await session.describe();
      const lines = description.item!.text.split("\n");
      assert.equal(lines[0], "pdf: 2 pages, 4 elements");
      assert.match(lines[2]!, /^\[[^\]]+\] page 1 textBox: A box of text$/);
      assert.match(
        lines[4]!,
        /^\[[^\]]+\] page 2 table "Table \(2×2\)": Region \| Q1 ⏎ North \| 120$/,
      );
    } finally {
      await end();
    }
  });

  it("outlines a deck with names, placeholders, groups and tables", async () => {
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
              placeholder: { type: "title" },
              paragraphs: [["Quarterly review"]],
            }),
            group({
              id: 3,
              name: "Group 2",
              x: 0,
              y: 914400,
              cx: 1828800,
              cy: 914400,
              child: { x: 0, y: 0, cx: 1828800, cy: 914400 },
              children: [
                textShape({
                  id: 4,
                  name: "Note 3",
                  x: 0,
                  y: 0,
                  cx: 914400,
                  cy: 457200,
                  paragraphs: [["Inside the group"]],
                }),
              ],
            }),
            table({
              id: 5,
              name: "Table 4",
              x: 0,
              y: 2743200,
              cx: 1828800,
              cy: 914400,
              rows: [
                ["Region", "Q1", "Q2"],
                ["North", "120", "130"],
              ],
            }),
          ],
        },
        {
          shapes: [
            textShape({
              id: 2,
              name: "",
              x: 0,
              y: 0,
              cx: 914400,
              cy: 457200,
              placeholder: { type: "subTitle" },
              paragraphs: [["Sub"]],
            }),
          ],
        },
      ],
    });
    const { session, end } = await pptxSession(deck);
    try {
      const outline = await session.getOutline();
      assert.deepEqual(
        outline.items.map((node) => [
          node.ordinal,
          node.id,
          node.label,
          node.text ?? "",
        ]),
        [
          ["1", "sld1:2", "Title 1", "Quarterly review"],
          ["2", "sld1:3", "Group 2", ""],
          ["3", "sld1:5", "Table 4 (2×3)", "Region\tQ1\tQ2\nNorth\t120\t130"],
          ["4", "sld2:2", "Subtitle", "Sub"],
        ],
      );
      assert.deepEqual(
        outline.items[1]!.children!.map((node) => [
          node.ordinal,
          node.id,
          node.label,
          node.text,
        ]),
        [["2.1", "sld1:4", "Note 3", "Inside the group"]],
      );
      assert.equal(outline.nodeCount, 5);
      const description = await session.describe({ pageRange: [0, 0] });
      assert.equal(
        description.item!.text,
        [
          "pptx: 2 slides, 4 elements",
          '[sld1:2] slide 1 shape "Title 1": Quarterly review',
          '[sld1:3] slide 1 group "Group 2"',
          '  [sld1:4] slide 1 shape "Note 3": Inside the group',
          '[sld1:5] slide 1 table "Table 4 (2×3)": Region | Q1 | Q2 ⏎ North | 120 | 130',
        ].join("\n"),
      );
    } finally {
      await end();
    }
  });

  it("outlines a Word document with styles, cells, pictures and unplaced pages", async () => {
    const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
    const bytes = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      styles:
        '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:keepNext/></w:pPr></w:style>',
      body:
        `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>` +
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
        `<w:tr><w:tc>${paragraph("A1")}</w:tc><w:tc>${paragraph("B1")}</w:tc></w:tr></w:tbl>` +
        `<w:p><w:r><w:t>Picture </w:t></w:r>${inlinePicture(914400, 457200)}</w:p>` +
        `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Item</w:t></w:r></w:p>` +
        sectPr(),
    });
    const { session, end } = await docxSession(bytes, [[]]);
    try {
      const outline = await session.getOutline();
      const ids = outline.items.map((node) => node.id);
      assert.equal(outline.nodeCount, 7);
      assert.deepEqual(
        outline.items.map((node) => [
          node.ordinal,
          node.kind,
          node.pageIndex,
          node.label ?? "",
          node.text ?? "",
        ]),
        [
          ["1", "paragraph", -1, "Heading 1", "Heading"],
          ["2", "table", -1, "Table (1×2)", "A1\tB1"],
          ["3", "paragraph", -1, "", "Picture \uFFFC"],
          ["4", "paragraph", -1, "List item (level 1)", "Item"],
        ],
      );
      assert.deepEqual(
        outline.items[1]!.children!.map((node) => [
          node.ordinal,
          node.kind,
          node.text,
        ]),
        [
          ["2.1", "paragraph", "A1"],
          ["2.2", "paragraph", "B1"],
        ],
      );
      assert.deepEqual(
        outline.items[2]!.children!.map((node) => [node.ordinal, node.kind]),
        [["3.1", "image"]],
      );
      // Placed pages give the nodes their page; the grammar then prints it.
      const placed = await docxSession(bytes, [
        [run(ids[0]!.slice(2), "Heading", 96, 96)],
      ]);
      try {
        const description = await placed.session.describe({
          pageRange: [0, 0],
        });
        assert.equal(
          description.item!.text,
          `docx: 1 page, 1 element\n[${ids[0]}] page 1 paragraph "Heading 1": Heading`,
        );
      } finally {
        await placed.end();
      }
      const description = await session.describe();
      assert.equal(
        description.item!.text.split("\n")[1],
        `[${ids[0]}] paragraph "Heading 1": Heading`,
      );
    } finally {
      await end();
    }
  });
});
