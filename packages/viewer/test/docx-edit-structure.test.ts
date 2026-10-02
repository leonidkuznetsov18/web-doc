import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { DocxEditEngine } from "../src/edit/docx/engine.js";
import type { DocxOperation } from "../src/edit/docx/types.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { IMAGE_RELATIONSHIP_TYPE } from "../src/edit/ooxml/opc.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";
import { docxSession, run } from "./fixtures/docx-session.js";

/*
 * Task 56 of the DOCX module: paragraphs inserted next to paragraphs and
 * tables, paragraphs, tables and pictures deleted, paragraphs and tables
 * moved within their container, pictures inserted as paragraphs of their
 * own; every id stable and every untouched byte kept.
 */

const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3,
]);
const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xe0, 4, 5, 6]);

async function open(bytes: Uint8Array): Promise<DocxEditEngine> {
  return DocxEditEngine.open(bytes, limits, signal);
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  return new TextDecoder().decode(await pkg.part("/word/document.xml"));
}

async function bodyOf(engine: DocxEditEngine): Promise<string> {
  const xml = await documentXml(await engine.materialize("save", {}, signal));
  return xml.slice(xml.indexOf("<w:body>") + 8, xml.indexOf("</w:body>"));
}

async function listing(
  engine: DocxEditEngine,
): Promise<[string, string, string | undefined][]> {
  return (await engine.getElements({}, signal)).map((element) => [
    element.kind,
    element.id,
    element.text,
  ]);
}

function apply(engine: DocxEditEngine, operations: readonly DocxOperation[]) {
  return engine.apply(operations, signal);
}

async function expectIssue(
  engine: DocxEditEngine,
  operation: DocxOperation,
  code: string,
  path?: string,
): Promise<void> {
  const issues = await engine.validate([operation], signal);
  assert.ok(
    issues.some(
      (issue) => issue.code === code && (!path || issue.path === path),
    ),
    `expected ${code} at ${path ?? "any"}, got ${JSON.stringify(issues)}`,
  );
}

const TABLE = (cells: readonly string[]): string =>
  `<w:tbl><w:tblPr/><w:tblGrid>${cells.map(() => '<w:gridCol w:w="2000"/>').join("")}</w:tblGrid><w:tr>${cells
    .map((cell) => `<w:tc>${paragraph(cell)}</w:tc>`)
    .join("")}</w:tr></w:tbl>`;

describe("DOCX structure (docx-edit T56)", () => {
  it("inserts paragraphs before and after paragraphs and tables with copied or given style", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p><w:pPr><w:pStyle w:val="Quote"/>${sectPr()}</w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>one</w:t></w:r></w:p>` +
          TABLE(["cell"]) +
          paragraph("last") +
          sectPr(),
      }),
    );
    const [one, table, cell, last] = (await listing(engine)).map((e) => e[1]);
    const change = await apply(engine, [
      { op: "insertParagraph", after: one!, text: "A\nB" },
      {
        op: "insertParagraph",
        before: table!,
        text: "C",
        style: { bold: true },
      },
      { op: "insertParagraph", after: table!, text: "D" },
      { op: "insertParagraph", after: last!, text: "E" },
      { op: "insertParagraph", before: cell!, text: "F" },
    ]);
    assert.equal(change.createdIds.length, 6);
    const [a, b, c, d, e, f] = change.createdIds;
    // The table is named after its first paragraph, which F now is.
    assert.deepEqual(change.remappedIds, { [table!]: `tbl:${f!.slice(2)}` });
    assert.deepEqual(
      (await listing(engine)).map((entry) => [entry[1], entry[2]]),
      [
        [one, "one"],
        [a, "A"],
        [b, "B"],
        [c, "C"],
        [`tbl:${f!.slice(2)}`, "F\ncell"],
        [f, "F"],
        [cell, "cell"],
        [d, "D"],
        [last, "last"],
        [e, "E"],
      ],
    );
    const body = await bodyOf(engine);
    // The copied properties leave the section break behind; the given style merges in.
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${a!.slice(2)}"><w:pPr><w:pStyle w:val="Quote"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>A</w:t></w:r></w:p>`,
      ),
      body,
    );
    // A table reference gives a plain paragraph styled like the table's first run.
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${c!.slice(2)}"><w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>C</w:t></w:r></w:p>`,
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${d!.slice(2)}"><w:r><w:t>D</w:t></w:r></w:p>`,
      ),
      body,
    );
    // The body's own section properties stay last.
    assert.ok(
      body.endsWith(
        `<w:p w14:paraId="${e!.slice(2)}"><w:r><w:t>E</w:t></w:r></w:p>${sectPr()}`,
      ),
      body,
    );
    await expectIssue(
      engine,
      { op: "insertParagraph", text: "x" },
      "invalid-value",
      "/before",
    );
    await expectIssue(
      engine,
      { op: "insertParagraph", before: one!, after: last!, text: "x" },
      "invalid-value",
    );
    await expectIssue(
      engine,
      { op: "insertParagraph", after: "p:00000000", text: "x" },
      "unknown-target",
      "/after",
    );
    await engine.dispose();
  });

  it("deletes paragraphs, tables and pictures, refusing the last paragraph and section breaks", async () => {
    const original = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      body:
        paragraph("one") +
        TABLE(["a", "b"]) +
        `<w:p><w:r><w:t>pic </w:t></w:r>${inlinePicture(914400, 914400)}<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>` +
        TABLE(["x"]) +
        TABLE(["y"]) +
        `<w:p><w:pPr>${sectPr()}</w:pPr><w:r><w:t>section</w:t></w:r></w:p>` +
        sectPr(),
    });
    const engine = await open(original);
    const before = await listing(engine);
    const [
      one,
      table,
      a,
      b,
      pic,
      image,
      tableX,
      cellX,
      tableY,
      cellY,
      section,
    ] = before.map((entry) => entry[1]);
    assert.equal(before.length, 11);
    await expectIssue(
      engine,
      { op: "deleteElement", target: section! },
      "section-break",
    );
    await expectIssue(
      engine,
      { op: "deleteElement", target: a! },
      "last-paragraph",
    );
    const other = (await engine.getElements({ kinds: ["other"] }, signal))[0];
    void other;
    // The picture goes with the relationship only it used; the run text stays.
    const picture = await apply(engine, [
      { op: "deleteElement", target: image! },
    ]);
    assert.deepEqual(picture.removedIds, [image]);
    let body = await bodyOf(engine);
    assert.ok(!body.includes("w:drawing"), body);
    assert.ok(
      body.includes("<w:t>pic </w:t>") &&
        body.includes('<w:t xml:space="preserve"> end</w:t>'),
    );
    const relsPkg = await OoxmlPackage.open(
      await engine.materialize("save", {}, signal),
      { limits },
    );
    const rels = await relsPkg.relationships("/word/document.xml");
    assert.equal(rels.byId("rId9"), undefined);
    // Deleting a table and a paragraph drops their ids; a body ending with a table gains a paragraph.
    const dropped = await apply(engine, [
      { op: "deleteElement", target: table! },
      { op: "deleteElement", target: one! },
    ]);
    assert.deepEqual(dropped.removedIds, [table, a, b, one]);
    assert.equal(dropped.reflowFrom, table!.slice(4));
    await apply(engine, [{ op: "deleteElement", target: section! }]).catch(
      () => {},
    );
    const trailing = await apply(engine, [
      { op: "deleteElement", target: pic! },
    ]);
    assert.deepEqual(trailing.createdIds, []);
    assert.deepEqual(
      (await listing(engine)).map((entry) => entry[1]),
      [tableX, cellX, tableY, cellY, section],
    );
    // The section paragraph cannot go; the last table before it can, and a
    // body that would end with a table gets an empty paragraph.
    body = await bodyOf(engine);
    assert.ok(body.includes("<w:t>section</w:t>"));
    await engine.dispose();

    const tail = await open(
      buildDocx({
        body: paragraph("p") + TABLE(["x"]) + TABLE(["y"]) + sectPr(),
      }),
    );
    const [, , , tableY2] = (await listing(tail)).map((entry) => entry[1]);
    const change = await apply(tail, [
      { op: "deleteElement", target: tableY2! },
    ]);
    assert.equal(change.createdIds.length, 1);
    assert.ok(
      (await bodyOf(tail)).endsWith(
        `</w:tbl><w:p w14:paraId="${change.createdIds[0]!.slice(2)}"></w:p>${sectPr()}`,
      ),
      await bodyOf(tail),
    );
    await tail.dispose();
  });

  it("moves paragraphs and tables within their container, keeping ids", async () => {
    const engine = await open(
      buildDocx({
        body:
          paragraph("one") +
          paragraph("two") +
          `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>${paragraph("c1")}${paragraph("c2")}</w:tc></w:tr></w:tbl>` +
          paragraph("three") +
          `<w:p><w:pPr>${sectPr()}</w:pPr><w:r><w:t>s</w:t></w:r></w:p>` +
          sectPr(),
      }),
    );
    const ids = (await listing(engine)).map((entry) => entry[1]);
    const [one, two, table, c1, c2, three, section] = ids;
    const change = await apply(engine, [
      { op: "moveElement", target: one!, after: three! },
      { op: "moveElement", target: table!, before: two! },
      { op: "moveElement", target: c2!, before: c1! },
    ]);
    assert.deepEqual(change.remappedIds, { [table!]: `tbl:${c2!.slice(2)}` });
    assert.deepEqual(
      (await listing(engine)).map((entry) => [entry[1], entry[2]]),
      [
        [`tbl:${c2!.slice(2)}`, "c2\nc1"],
        [c2, "c2"],
        [c1, "c1"],
        [two, "two"],
        [three, "three"],
        [one, "one"],
        [section, "s"],
      ],
    );
    assert.equal(change.reflowFrom, one!.slice(2));
    const body = await bodyOf(engine);
    // Moved paragraphs carry their ids; the untouched ones keep their bytes.
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${one!.slice(2)}"><w:r><w:t>one</w:t></w:r></w:p>`,
      ),
      body,
    );
    assert.ok(body.includes(`<w:p><w:r><w:t>two</w:t></w:r></w:p>`), body);
    assert.ok(
      body.includes(
        `<w:tc><w:p w14:paraId="${c2!.slice(2)}"><w:r><w:t>c2</w:t></w:r></w:p><w:p w14:paraId="${c1!.slice(2)}"><w:r><w:t>c1</w:t></w:r></w:p></w:tc>`,
      ),
      body,
    );
    // The shown copy stamps every paragraph with the same ids the engine reports.
    const shown = await documentXml(
      await engine.materialize("show", {}, signal),
    );
    for (const id of [one, two, c1, c2, three, section])
      assert.ok(shown.includes(`w14:paraId="${id!.slice(2)}"`), String(id));
    await expectIssue(
      engine,
      { op: "moveElement", target: c1!, after: two! },
      "invalid-target",
      "/after",
    );
    await expectIssue(
      engine,
      { op: "moveElement", target: section!, before: one! },
      "section-break",
    );
    await expectIssue(
      engine,
      { op: "moveElement", target: one!, after: one! },
      "invalid-target",
    );
    await expectIssue(
      engine,
      { op: "moveElement", target: one! },
      "invalid-value",
    );
    // Replaying from the original gives the same order and ids.
    await engine.restore(
      [
        [
          { op: "moveElement", target: one!, after: three! } as DocxOperation,
          { op: "moveElement", target: table!, before: two! } as DocxOperation,
          { op: "moveElement", target: c2!, before: c1! } as DocxOperation,
        ],
      ],
      signal,
    );
    assert.deepEqual(
      (await listing(engine)).map((entry) => entry[1]),
      [`tbl:${c2!.slice(2)}`, c2, c1, two, three, one, section],
    );
    await engine.dispose();
  });

  it("inserts pictures as paragraphs, storing the bytes once and relating them from the document", async () => {
    const engine = await open(
      buildDocx({ body: paragraph("one") + paragraph("two") + sectPr() }),
    );
    const [one, two] = (await listing(engine)).map((entry) => entry[1]);
    const first = await apply(engine, [
      {
        op: "insertImage",
        after: one!,
        data: PNG,
        mimeType: "image/png",
        size: { width: 72, height: 36 },
      },
      {
        op: "insertImage",
        before: "$0",
        data: PNG,
        mimeType: "image/png",
        size: { width: 10, height: 10 },
      },
      {
        op: "insertImage",
        after: two!,
        data: JPEG,
        mimeType: "image/jpeg",
        size: { width: 5, height: 5 },
      },
    ]);
    assert.equal(first.createdIds.length, 6);
    const items = await listing(engine);
    assert.deepEqual(
      items.map((entry) => entry[0]),
      [
        "paragraph",
        "paragraph",
        "image",
        "paragraph",
        "image",
        "paragraph",
        "paragraph",
        "image",
      ],
    );
    assert.equal(items[1]![2], "￼");
    const saved = await engine.materialize("save", {}, signal);
    const pkg = await OoxmlPackage.open(saved, { limits });
    const media = pkg.currentPartNames.filter((name) =>
      name.startsWith("/word/media/"),
    );
    assert.deepEqual(media.sort(), [
      "/word/media/image1.jpeg",
      "/word/media/image1.png",
    ]);
    const xml = await documentXml(saved);
    assert.match(xml, /<wp:extent cx="914400" cy="457200"\/>/);
    assert.match(xml, /<wp:docPr id="1" name="Picture 1"\/>/);
    assert.match(xml, /<wp:docPr id="2" name="Picture 2"\/>/);
    assert.match(xml, /<wp:docPr id="3" name="Picture 3"\/>/);
    const rels = await pkg.relationships("/word/document.xml");
    const images = rels.byType(IMAGE_RELATIONSHIP_TYPE);
    assert.equal(images.length, 3);
    assert.equal(new Set(images.map((item) => item.target)).size, 2);
    await expectIssue(
      engine,
      {
        op: "insertImage",
        after: one!,
        data: JPEG,
        mimeType: "image/png",
        size: { width: 1, height: 1 },
      },
      "invalid-value",
      "/data",
    );
    await expectIssue(
      engine,
      {
        op: "insertImage",
        after: one!,
        data: "asset:missing",
        mimeType: "image/png",
        size: { width: 1, height: 1 },
      },
      "unknown-asset",
      "/data",
    );
    await engine.dispose();
  });

  it("works through the session with reflow pages and typed methods", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + sectPr(),
    });
    const probe = await open(original);
    const [a, b] = (await listing(probe)).map((entry) => entry[1]!.slice(2));
    await probe.dispose();
    const { session, end } = await docxSession(original, [
      [run(a!, "one", 72, 72)],
      [run(b!, "two", 72, 72)],
    ]);
    try {
      const inserted = await session.insertParagraph({
        after: `p:${b}`,
        text: "three",
      });
      assert.deepEqual(inserted.changedPages, [1]);
      const moved = await session.moveElement({
        target: `p:${b}`,
        before: `p:${a}`,
      });
      assert.deepEqual(moved.changedPages, [0, 1]);
      const image = await session.insertImage({
        before: `p:${a}`,
        data: PNG,
        mimeType: "image/png",
        size: { width: 20, height: 20 },
      });
      assert.equal(image.createdIds.length, 2);
      const deleted = await session.deleteElement({
        target: image.createdIds[1]!,
      });
      assert.deepEqual(deleted.removedIds, [image.createdIds[1]]);
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.text),
        ["two", "", "one", "three"],
      );
      await assert.rejects(
        session.deleteElement({ target: "p:00000000" }),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      await session.reset();
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });
});
