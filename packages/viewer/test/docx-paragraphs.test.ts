import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  createDocxParagraphIdResolver,
  paragraphIdOf,
  type DocxModelBlock,
  type DocxModelDocument,
} from "../src/adapters/docx-paragraphs.js";

/*
 * Task 53 of the DOCX engine upgrade: a run's source (story, instance and
 * block path in the engine's model) leads to the paragraph whose pre-pass
 * bookmark names its w:p, for every story the engine lays out, including
 * paragraphs the engine splits around a page break.
 */

const paragraph = (
  id: string | undefined,
  extra: Partial<DocxModelBlock> = {},
): DocxModelBlock => ({
  type: "paragraph",
  ...(id === undefined ? {} : { bookmarks: ["_GoBack", `_wd${id}`] }),
  ...extra,
});

const model: DocxModelDocument = {
  body: [
    paragraph("1A000000"),
    {
      type: "table",
      rows: [
        { cells: [{ content: [paragraph("1A000001")] }] },
        {
          cells: [
            { content: [paragraph("1A000002")] },
            {
              content: [
                {
                  type: "table",
                  rows: [{ cells: [{ content: [paragraph("1A000003")] }] }],
                },
              ],
            },
          ],
        },
      ],
    },
    paragraph("1A000004"),
    { type: "pageBreak" },
    paragraph(undefined),
    { type: "pageBreak" },
    paragraph(undefined),
    {
      type: "sectionBreak",
      headers: { default: { body: [paragraph("1A000005")] } },
      footers: { first: { body: [paragraph("1A000006")] } },
    },
    paragraph("1A000007", {
      runs: [
        { type: "text" },
        { type: "shape", textBoxContent: [paragraph("1A000008")] },
      ],
    } as Partial<DocxModelBlock>),
    paragraph(undefined),
    { type: "paragraph", paragraphId: "2B000000" },
  ],
  headers: { first: { body: [paragraph("1A000009")] } },
  footers: { default: { body: [paragraph("1A00000A")] }, even: null },
  footnotes: [{ id: "1", content: [paragraph("1A00000B")] }],
  endnotes: [{ id: "7", content: [paragraph("1A00000C")] }],
};

const body = (path: number[]) => ({
  story: "body",
  storyInstance: "body",
  path,
});

describe("DOCX paragraph bridge (docx-engine-upgrade)", () => {
  const resolve = createDocxParagraphIdResolver(model);

  it("maps body paragraphs and table cells by their block path", () => {
    assert.equal(resolve(body([0, 0])), "1A000000");
    assert.equal(resolve(body([1, 0, 0, 0, 2])), "1A000001");
    assert.equal(resolve(body([1, 1, 0, 0])), "1A000002");
    assert.equal(resolve(body([1, 1, 1, 0, 0, 0, 0, 1])), "1A000003");
    assert.equal(resolve(body([2])), "1A000004");
  });

  it("gives a paragraph split around page breaks the id of its first part", () => {
    assert.equal(resolve(body([4, 0])), "1A000004");
    assert.equal(resolve(body([6, 3])), "1A000004");
  });

  it("leaves an unmarked paragraph that is not a continuation without an id", () => {
    assert.equal(resolve(body([9, 0])), undefined);
  });

  it("prefers a paragraph id the engine read from w14:paraId", () => {
    assert.equal(resolve(body([10, 0])), "2B000000");
    assert.equal(
      paragraphIdOf({
        type: "paragraph",
        paragraphId: "",
        bookmarks: ["_wdX"],
      }),
      "X",
    );
  });

  it("walks the document's and a section's headers and footers", () => {
    const run = (story: string, storyInstance: string, path: number[]) => ({
      story,
      storyInstance,
      path,
    });
    assert.equal(resolve(run("header", "first", [0, 0])), "1A000009");
    assert.equal(resolve(run("footer", "default", [0])), "1A00000A");
    assert.equal(resolve(run("footer", "even", [0])), undefined);
    assert.equal(resolve(run("header", "section:7:default", [0])), "1A000005");
    assert.equal(resolve(run("footer", "section:7:first", [0])), "1A000006");
    assert.equal(resolve(run("header", "section:7:first", [0])), undefined);
    assert.equal(resolve(run("header", "section:0:default", [0])), undefined);
    assert.equal(resolve(run("header", "weird", [0])), undefined);
  });

  it("walks footnotes, endnotes and text boxes through their instance", () => {
    assert.equal(
      resolve({ story: "footnote", storyInstance: "1", path: [0, 0] }),
      "1A00000B",
    );
    assert.equal(
      resolve({ story: "endnote", storyInstance: "7", path: [0] }),
      "1A00000C",
    );
    assert.equal(
      resolve({ story: "endnote", storyInstance: "8", path: [0] }),
      undefined,
    );
    assert.equal(
      resolve({
        story: "textbox",
        storyInstance: "body:body:8.1",
        path: [0, 0],
      }),
      "1A000008",
    );
    assert.equal(
      resolve({ story: "textbox", storyInstance: "body:body:8.0", path: [0] }),
      undefined,
    );
    assert.equal(
      resolve({ story: "textbox", storyInstance: "nonsense", path: [0] }),
      undefined,
    );
  });

  it("returns nothing for paths that lead outside the model or to no paragraph", () => {
    assert.equal(resolve(body([3])), undefined);
    assert.equal(resolve(body([42])), undefined);
    assert.equal(resolve(body([1, 5, 0, 0])), undefined);
    assert.equal(resolve(body([1])), undefined);
    assert.equal(resolve(body([])), undefined);
    assert.equal(resolve(undefined), undefined);
    assert.equal(
      resolve({ story: "comment", storyInstance: "1", path: [0] }),
      undefined,
    );
    assert.equal(
      createDocxParagraphIdResolver(undefined)(body([0])),
      undefined,
    );
  });

  it("caches per source and survives a model that throws while walked", () => {
    // One read when the resolver is created, one per uncached source.
    let reads = 0;
    const lazy: DocxModelDocument = {
      get body(): readonly DocxModelBlock[] {
        reads += 1;
        if (reads > 2) throw new Error("gone");
        return [paragraph("1A0000FF")];
      },
    };
    const resolver = createDocxParagraphIdResolver(lazy);
    assert.equal(resolver(body([0])), "1A0000FF");
    assert.equal(resolver(body([0])), "1A0000FF");
    assert.equal(resolver(body([0, 1])), undefined);
    assert.equal(reads, 3);
  });
});
