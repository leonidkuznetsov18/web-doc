import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { DocxEditEngine } from "../src/edit/docx/engine.js";
import type { DocxOperation } from "../src/edit/docx/types.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { defaultResourceLimits } from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";
import { docxSession, run } from "./fixtures/docx-session.js";

/*
 * Task 57 of the DOCX module: tables inserted next to body blocks with a
 * grid over the section's content width and one paragraph per cell, and
 * cell text replaced in the cell's first paragraph.
 */

const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);

async function open(bytes: Uint8Array): Promise<DocxEditEngine> {
  return DocxEditEngine.open(bytes, limits, signal);
}

async function bodyOf(engine: DocxEditEngine): Promise<string> {
  const pkg = await OoxmlPackage.open(
    await engine.materialize("save", {}, signal),
    { limits },
  );
  const xml = new TextDecoder().decode(await pkg.part("/word/document.xml"));
  return xml.slice(xml.indexOf("<w:body>") + 8, xml.indexOf("</w:body>"));
}

function apply(engine: DocxEditEngine, operations: readonly DocxOperation[]) {
  return engine.apply(operations, signal);
}

async function ids(engine: DocxEditEngine): Promise<string[]> {
  return (await engine.getElements({}, signal)).map((element) => element.id);
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

describe("DOCX tables (docx-edit T57)", () => {
  it("inserts a table over the section width with cell paragraphs and a trailing paragraph where needed", async () => {
    const engine = await open(
      buildDocx({
        styles:
          '<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style>',
        body:
          paragraph("one") +
          `<w:p><w:pPr>${sectPr({ width: 8000, margin: 500 })}</w:pPr><w:r><w:t>narrow section</w:t></w:r></w:p>` +
          paragraph("two") +
          sectPr(),
      }),
    );
    const [one, narrow, two] = await ids(engine);
    const change = await apply(engine, [
      {
        op: "insertTable",
        after: one!,
        rows: [
          ["a", "b\nc"],
          ["d", ""],
        ],
        columnWidths: [1, 3],
      },
      { op: "insertTable", after: two!, rows: [["end"]] },
      { op: "insertTable", before: "$1", rows: [["mid"]] },
    ]);
    // Each table names itself, then its cell paragraphs, then a trailing
    // paragraph when the body needs one (the second and third do).
    assert.equal(change.createdIds.length, 5 + 3 + 3);
    const [table1, a, b, d, e] = change.createdIds;
    assert.equal(table1, `tbl:${a!.slice(2)}`);
    const body = await bodyOf(engine);
    // The first section is 7000 twips wide: weights 1 and 3 give 1750 and 5250.
    assert.ok(
      body.includes(
        `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1750"/><w:gridCol w:w="5250"/></w:tblGrid>` +
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="1750" w:type="dxa"/></w:tcPr><w:p w14:paraId="${a!.slice(2)}"><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="5250" w:type="dxa"/></w:tcPr><w:p w14:paraId="${b!.slice(2)}"><w:r><w:t>b</w:t><w:br/><w:t>c</w:t></w:r></w:p></w:tc></w:tr>` +
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="1750" w:type="dxa"/></w:tcPr><w:p w14:paraId="${d!.slice(2)}"><w:r><w:t>d</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="5250" w:type="dxa"/></w:tcPr><w:p w14:paraId="${e!.slice(2)}"></w:p></w:tc></w:tr></w:tbl>`,
      ),
      body,
    );
    // After "one" the next block is a paragraph: no trailing paragraph; the
    // table after "two" would meet the body's sectPr and gets one; "mid"
    // sits before a table and gets one too.
    const texts = (await engine.getElements({}, signal)).map((element) => [
      element.kind,
      element.text,
    ]);
    assert.deepEqual(texts, [
      ["paragraph", "one"],
      ["table", "a\tb\u000bc\nd\t"],
      ["paragraph", "a"],
      ["paragraph", "b\u000bc"],
      ["paragraph", "d"],
      ["paragraph", ""],
      ["paragraph", "narrow section"],
      ["paragraph", "two"],
      ["table", "mid"],
      ["paragraph", "mid"],
      ["paragraph", ""],
      ["table", "end"],
      ["paragraph", "end"],
      ["paragraph", ""],
    ]);
    assert.ok(
      body.endsWith(
        `</w:tbl><w:p w14:paraId="${change.createdIds[7]!.slice(2)}"></w:p>${sectPr()}`,
      ),
      body.slice(-300),
    );
    assert.equal(change.reflowFrom, one!.slice(2));
    assert.equal(narrow, (await ids(engine))[6]);
    // Without a TableGrid style the table carries single borders.
    const plain = await open(buildDocx({ body: paragraph("x") + sectPr() }));
    await apply(plain, [
      {
        op: "insertTable",
        after: (await ids(plain))[0]!,
        rows: [["1", "2", "3"]],
      },
    ]);
    const plainBody = await bodyOf(plain);
    assert.ok(
      plainBody.includes(
        '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>',
      ),
      plainBody,
    );
    assert.ok(
      plainBody.includes(
        '<w:gridCol w:w="3120"/><w:gridCol w:w="3120"/><w:gridCol w:w="3120"/>',
      ),
    );
    await plain.dispose();
    await expectIssue(
      engine,
      { op: "insertTable", after: a!, rows: [["x"]] },
      "invalid-target",
      "/after",
    );
    await expectIssue(
      engine,
      { op: "insertTable", after: one!, rows: [["x", "y"], ["z"]] },
      "invalid-value",
      "/rows",
    );
    await expectIssue(
      engine,
      { op: "insertTable", after: one!, rows: [["x"]], columnWidths: [1, 2] },
      "invalid-value",
      "/columnWidths",
    );
    await expectIssue(
      engine,
      { op: "insertTable", after: one!, rows: [["x\u0000"]] },
      "invalid-text",
      "/rows/0/0",
    );
    await engine.dispose();
  });

  it("replaces a cell's text in its first paragraph and removes the cell's other paragraphs", async () => {
    const engine = await open(
      buildDocx({
        media: [{ name: "word/media/image1.png", data: PNG }],
        body:
          `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
          `<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>first</w:t></w:r></w:p>${paragraph("second")}<w:p>${inlinePicture(100, 100)}</w:p></w:tc>` +
          `<w:tc>${paragraph("other")}</w:tc></w:tr></w:tbl>` +
          paragraph("after") +
          sectPr(),
      }),
    );
    const [table, first, second, picParagraph, picture, other] =
      await ids(engine);
    const change = await apply(engine, [
      {
        op: "setTableCell",
        target: table!,
        row: 0,
        column: 0,
        text: "new\nline",
      },
    ]);
    assert.deepEqual(change.removedIds, [second, picParagraph, picture]);
    const body = await bodyOf(engine);
    assert.ok(
      body.includes(
        `<w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr><w:p w14:paraId="${first!.slice(2)}"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>new</w:t><w:br/><w:t>line</w:t></w:r></w:p></w:tc>`,
      ),
      body,
    );
    assert.deepEqual(await ids(engine), [
      table,
      first,
      other,
      (await ids(engine))[3],
    ]);
    const element = (await engine.getElement(table!, signal))!;
    assert.deepEqual(element.table, { rows: [["new\u000bline", "other"]] });
    // An empty text leaves an empty paragraph with its properties.
    await apply(engine, [
      { op: "setTableCell", target: table!, row: 0, column: 1, text: "" },
    ]);
    assert.ok(
      (await bodyOf(engine)).includes(
        `<w:p w14:paraId="${other!.slice(2)}"></w:p></w:tc>`,
      ),
    );
    await expectIssue(
      engine,
      { op: "setTableCell", target: table!, row: 1, column: 0, text: "x" },
      "range",
      "/row",
    );
    await expectIssue(
      engine,
      { op: "setTableCell", target: table!, row: 0, column: 2, text: "x" },
      "range",
      "/column",
    );
    await expectIssue(
      engine,
      { op: "setTableCell", target: first!, row: 0, column: 0, text: "x" },
      "invalid-target",
    );
    await expectIssue(
      engine,
      { op: "setTableCell", target: table!, row: 0, column: 0, text: "\u0001" },
      "invalid-text",
      "/text",
    );
    await engine.dispose();
  });

  it("works through the session: a new table reflows from its reference and replays", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + sectPr(),
    });
    const probe = await open(original);
    const [a, b] = (await ids(probe)).map((id) => id.slice(2));
    await probe.dispose();
    const { session, end } = await docxSession(original, [
      [run(a!, "one", 72, 72)],
      [run(b!, "two", 72, 72)],
    ]);
    try {
      const inserted = await session.insertTable({
        after: `p:${b}`,
        rows: [["x", "y"]],
      });
      assert.deepEqual(inserted.changedPages, [1]);
      assert.equal(inserted.createdIds.length, 4);
      const table = inserted.createdIds[0]!;
      const cell = await session.setTableCell({
        target: table,
        row: 0,
        column: 1,
        text: "Y",
      });
      // The fake renderer never laid the new table out, so the host cannot
      // place its paragraph and the whole document repaints.
      assert.deepEqual(cell.changedPages, [0, 1]);
      const element = (await session.getElement(table)).item!;
      assert.deepEqual(element.table, { rows: [["x", "Y"]] });
      await session.undo();
      await session.undo();
      assert.deepEqual((await session.save()).bytes, original);
      await session.redo();
      await session.redo();
      assert.deepEqual(
        ((await session.getElement(table)).item as { table: unknown }).table,
        { rows: [["x", "Y"]] },
      );
    } finally {
      await end();
    }
  });
});
