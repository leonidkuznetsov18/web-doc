import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import { prepareDocxForDisplay } from "../src/adapters/docx-prepass.js";
import { DocxEditEngine } from "../src/edit/docx/engine.js";
import { docxOperationSchemaDrafts } from "../src/edit/docx/schemas.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { assertSupportedSchema } from "../src/edit/schema.js";
import {
  defaultResourceLimits,
  type DocxElement,
  type EditOperation,
} from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";

/*
 * Task 54 of the DOCX module: the engine opens a document, lists every
 * paragraph, table, inline picture and other object of the body with ids
 * that match the display pre-pass, text models, resolved styles and the
 * operations each accepts; text search; identity without changes; restore.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const SAMPLE = new URL("../../.cache/corpus/sample.docx", PACKAGE_DIR);
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const W14 =
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

async function elementsOf(bytes: Uint8Array): Promise<{
  engine: DocxEditEngine;
  elements: readonly DocxElement[];
}> {
  const engine = await DocxEditEngine.open(bytes, limits, signal);
  const elements = await engine.getElements({}, signal);
  return { engine, elements };
}

/** The `_wd` bookmark ids the pre-pass writes into the main part, in order. */
async function markedIds(bytes: Uint8Array): Promise<string[]> {
  const display = await prepareDocxForDisplay(bytes, limits);
  const pkg = await OoxmlPackage.open(display.bytes, { limits });
  const xml = new TextDecoder().decode(await pkg.part("/word/document.xml"));
  return [...xml.matchAll(/w:name="_wd([0-9A-F]{8})"/g)].map((m) => m[1]!);
}

describe("DOCX edit engine: inspection (docx-edit T54)", () => {
  it("lists paragraphs, tables, cell paragraphs and pictures with pre-pass ids", async () => {
    const bytes = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      body:
        paragraph("First") +
        `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
        `<w:tr><w:tc>${paragraph("A1")}</w:tc><w:tc>${paragraph("B1")}${paragraph("B1 more")}</w:tc></w:tr>` +
        `<w:tr><w:tc>${paragraph("A2")}</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>` +
        `<w:p><w:r><w:t>Picture </w:t></w:r>${inlinePicture(914400, 457200)}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>` +
        `<w:sdt><w:sdtPr/><w:sdtContent>${paragraph("Inside sdt")}</w:sdtContent></w:sdt>` +
        sectPr(),
    });
    const { engine, elements } = await elementsOf(bytes);
    try {
      const ids = await markedIds(bytes);
      assert.equal(ids.length, 8);
      assert.deepEqual(
        elements.map((element) => [element.kind, element.id]),
        [
          ["paragraph", `p:${ids[0]}`],
          ["table", `tbl:${ids[1]}`],
          ["paragraph", `p:${ids[1]}`],
          ["paragraph", `p:${ids[2]}`],
          ["paragraph", `p:${ids[3]}`],
          ["paragraph", `p:${ids[4]}`],
          ["paragraph", `p:${ids[5]}`],
          ["paragraph", `p:${ids[6]}`],
          ["image", `img:${ids[6]}.0`],
          ["paragraph", `p:${ids[7]}`],
        ],
      );
      const table = elements[1]!;
      assert.equal(table.text, "A1\tB1\nB1 more\nA2\t");
      assert.deepEqual(table.table, {
        rows: [
          ["A1", "B1\nB1 more"],
          ["A2", ""],
        ],
      });
      assert.equal(elements[2]!.parentId, table.id);
      assert.equal(elements[6]!.parentId, table.id);
      const picture = elements[7]!;
      assert.equal(picture.text, "Picture ￼ after");
      assert.equal(elements[8]!.parentId, picture.id);
      assert.equal(elements[9]!.text, "Inside sdt");
      for (const element of elements) {
        assert.equal(element.pageIndex, -1);
        assert.deepEqual(element.bounds, { x: 0, y: 0, width: 0, height: 0 });
        assert.deepEqual(element.fragments, []);
        assert.deepEqual(element.story, { kind: "body" });
        // No handler has shipped yet: no element accepts an operation.
        assert.deepEqual(element.operations, []);
      }
      assert.equal(
        (await engine.getElement(picture.id, signal))?.text,
        picture.text,
      );
      assert.equal(await engine.getElement("p:nope", signal), undefined);
      assert.deepEqual(
        (await engine.getElements({ kinds: ["image", "table"] }, signal)).map(
          (element) => element.kind,
        ),
        ["table", "image"],
      );
    } finally {
      await engine.dispose();
    }
  });

  it("keeps authored w14:paraId values and numbers the rest like the pre-pass", async () => {
    const bytes = buildDocx({
      rootAttributes: W14,
      body:
        paragraph("Authored", 'w14:paraId="1A00ABCD"') +
        paragraph("Generated") +
        paragraph("Authored too", 'w14:paraId="0000ffff"') +
        sectPr(),
    });
    const { engine, elements } = await elementsOf(bytes);
    try {
      const ids = await markedIds(bytes);
      assert.deepEqual(
        elements.map((element) => element.id),
        ids.map((id) => `p:${id}`),
      );
      assert.equal(elements[0]!.id, "p:1A00ABCD");
      assert.equal(elements[2]!.id, "p:0000FFFF");
    } finally {
      await engine.dispose();
    }
  });

  it("reads the text model: tabs, breaks, fields, hyperlinks, inline sdt, symbols, notes", async () => {
    const body =
      `<w:p><w:r><w:t>A</w:t><w:tab/><w:t>B</w:t><w:br/><w:t>C</w:t><w:br w:type="page"/><w:t xml:space="preserve">D </w:t></w:r>` +
      `<w:hyperlink r:id="rId9"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>link</w:t></w:r></w:hyperlink>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:fldSimple w:instr="AUTHOR"><w:r><w:t>Ann</w:t></w:r></w:fldSimple>` +
      `<w:sdt><w:sdtContent><w:r><w:t>sdt</w:t></w:r></w:sdtContent></w:sdt>` +
      `<w:r><w:sym w:font="Wingdings" w:char="0041"/><w:noBreakHyphen/><w:footnoteReference w:id="1"/><w:t>!</w:t></w:r>` +
      `</w:p>` +
      sectPr();
    const { engine, elements } = await elementsOf(buildDocx({ body }));
    try {
      assert.equal(elements.length, 1);
      assert.equal(elements[0]!.text, "A\tB\u000bC\fD link7AnnsdtA‑!");
      const model = await engine.model(signal);
      const items = model.paragraphs[0]!.text.items;
      assert.deepEqual(
        items.map((item) => item.kind),
        [
          "text",
          "tab",
          "text",
          "break",
          "text",
          "pageBreak",
          "text",
          "text",
          "field",
          "field",
          "text",
          "text",
          "text",
          "note",
          "text",
        ],
      );
      const complex = items[8]!;
      assert.equal(complex.text, "7");
      assert.equal(complex.fieldRuns?.length, 5);
      assert.equal(items[7]!.wrapper?.local, "hyperlink");
      assert.equal(items[10]!.wrapper?.local, "sdt");
    } finally {
      await engine.dispose();
    }
  });

  it("marks tracked, section-break and field-only paragraphs read-only and lists other objects", async () => {
    const body =
      `<w:p><w:ins w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>inserted</w:t></w:r></w:ins><w:del w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>` +
      `<w:p><w:pPr>${sectPr()}</w:pPr><w:r><w:t>ends a section</w:t></w:r></w:p>` +
      `<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>3</w:t></w:r></w:fldSimple></w:p>` +
      `<w:p><w:r><w:drawing><wp:anchor><wp:extent cx="10" cy="10"/></wp:anchor></w:drawing></w:r><w:r><w:object><v:shape xmlns:v="urn:schemas-microsoft-com:vml"/></w:object></w:r><w:r><w:t>text</w:t></w:r></w:p>` +
      `<w:p><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:oMath/></m:oMathPara></w:p>` +
      sectPr();
    const { engine, elements } = await elementsOf(buildDocx({ body }));
    try {
      assert.deepEqual(
        elements.map((element) => [element.kind, element.readOnlyReason]),
        [
          ["paragraph", "tracked-changes"],
          ["paragraph", "section-break"],
          ["paragraph", "unsupported-content"],
          ["paragraph", undefined],
          ["other", undefined],
          ["other", undefined],
          ["paragraph", undefined],
          ["other", undefined],
        ],
      );
      assert.equal(elements[0]!.text, "inserted");
      assert.equal(elements[3]!.text, "￼text");
      assert.equal(elements[4]!.id, `other:${elements[3]!.id.slice(2)}.0`);
      assert.equal(elements[5]!.id, `other:${elements[3]!.id.slice(2)}.1`);
      assert.equal(elements[6]!.text, "￼");
    } finally {
      await engine.dispose();
    }
  });

  it("resolves run and paragraph styles through styles, defaults and the theme", async () => {
    const styles =
      '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="0"/><w:jc w:val="center"/></w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/><w:b/><w:color w:val="2F5496" w:themeColor="accent1" w:themeShade="BF"/><w:sz w:val="32"/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Strong"><w:rPr><w:b/><w:i w:val="0"/><w:u w:val="single"/><w:highlight w:val="yellow"/></w:rPr></w:style>';
    const body =
      paragraph("Plain") +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:jc w:val="right"/><w:spacing w:before="120" w:line="480" w:lineRule="exact"/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Strong"/><w:rFonts w:ascii="Georgia"/><w:color w:val="FF0000"/><w:sz w:val="20"/></w:rPr><w:t>Styled</w:t></w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:b w:val="0"/><w:i/><w:u w:val="none"/><w:color w:val="auto"/></w:rPr><w:t>Toggles</w:t></w:r></w:p>` +
      sectPr();
    const { engine, elements } = await elementsOf(
      buildDocx({
        body,
        styles,
        theme: { major: "Aptos Display", minor: "Aptos" },
      }),
    );
    try {
      assert.deepEqual(elements[0]!.textStyle, {
        fontFamily: "Aptos",
        fontSize: 11,
        bold: false,
        italic: false,
        underline: false,
        color: "auto",
      });
      assert.deepEqual(elements[0]!.paragraphStyle, {
        align: "left",
        spacing: { after: 8, line: 259 / 240, lineRule: "auto" },
      });
      assert.deepEqual(elements[1]!.textStyle, {
        fontFamily: "Aptos Display",
        fontSize: 16,
        bold: true,
        italic: false,
        underline: false,
        color: { theme: "accent1" },
      });
      assert.deepEqual(elements[1]!.paragraphStyle, {
        styleId: "Heading1",
        align: "center",
        spacing: { before: 12, after: 0, line: 259 / 240, lineRule: "auto" },
      });
      assert.deepEqual(elements[2]!.textStyle, {
        fontFamily: "Georgia",
        fontSize: 10,
        bold: true,
        italic: false,
        underline: true,
        color: "#FF0000",
        highlight: "yellow",
      });
      assert.deepEqual(elements[2]!.paragraphStyle, {
        align: "right",
        spacing: { before: 6, after: 8, line: 24, lineRule: "exact" },
        numbering: { numId: 3, level: 1 },
      });
      assert.deepEqual(elements[3]!.textStyle, {
        fontFamily: "Aptos",
        fontSize: 11,
        bold: false,
        italic: true,
        underline: false,
        color: "auto",
      });
    } finally {
      await engine.dispose();
    }
  });

  it("finds text in document order with element ranges and no geometry", async () => {
    const body =
      paragraph("alpha beta Alpha") +
      `<w:tbl><w:tr><w:tc>${paragraph("beta alpha")}</w:tc></w:tr></w:tbl>` +
      sectPr();
    const { engine, elements } = await elementsOf(buildDocx({ body }));
    try {
      const hits = await engine.findText("alpha", {}, signal);
      assert.deepEqual(
        hits.map((hit) => [
          hit.elementIds[0],
          hit.ranges[0]!.start.offset,
          hit.text,
        ]),
        [
          [elements[0]!.id, 0, "alpha"],
          [elements[0]!.id, 11, "Alpha"],
          [elements[2]!.id, 5, "alpha"],
        ],
      );
      assert.deepEqual(hits[0]!.rects, []);
      assert.equal(hits[0]!.pageIndex, -1);
      assert.equal(
        (await engine.findText("alpha", { caseSensitive: true }, signal))
          .length,
        2,
      );
      assert.equal(
        (await engine.findText("alpha", { maxResults: 1 }, signal)).length,
        1,
      );
      assert.deepEqual(await engine.findText("", {}, signal), []);
      assert.deepEqual(await engine.elementsAt(0, { x: 1, y: 1 }, signal), []);
    } finally {
      await engine.dispose();
    }
  });

  it("returns the original bytes without changes, rejects unknown operations and restores", async () => {
    const bytes = buildDocx({ body: paragraph("x") + sectPr() });
    const engine = await DocxEditEngine.open(bytes, limits, signal);
    try {
      assert.deepEqual(await engine.materialize(signal), bytes);
      const issues = await engine.validate(
        [{ op: "replaceText", target: "p:1", text: "y" } as EditOperation],
        signal,
      );
      assert.deepEqual(
        issues.map((issue) => issue.code),
        ["unknown-operation"],
      );
      await assert.rejects(
        engine.apply([{ op: "nope" } as EditOperation], signal),
        (error: unknown) =>
          error instanceof Error && error.message.includes("nope"),
      );
      assert.deepEqual(await engine.materialize(signal), bytes);
      await engine.restore({ batches: [] }, signal);
      assert.equal((await engine.getElements({}, signal)).length, 1);
      await engine.dispose();
      await assert.rejects(engine.getElements({}, signal));
    } finally {
      await engine.dispose();
    }
  });

  it("refuses a package without a document part", async () => {
    const bytes = buildDocx({ body: paragraph("x") });
    const pkg = await OoxmlPackage.open(bytes, { limits });
    const transaction = pkg.transaction();
    transaction.removePart("/word/document.xml");
    await transaction.commit();
    await assert.rejects(
      DocxEditEngine.open(await pkg.save(), limits, signal),
      (error: unknown) =>
        error instanceof Error && /document part/.test(error.message),
    );
  });

  it("declares every operation with a schema the validator supports", () => {
    for (const [name, schema] of Object.entries(docxOperationSchemaDrafts)) {
      assertSupportedSchema(schema);
      assert.equal(
        (schema.properties as Record<string, { const?: string }>).op?.const,
        name,
      );
    }
  });

  it("lists the corpus document with ids the pre-pass agrees on", async (t) => {
    if (!existsSync(SAMPLE)) {
      t.skip("corpus not fetched");
      return;
    }
    const bytes = new Uint8Array(readFileSync(SAMPLE));
    const { engine, elements } = await elementsOf(bytes);
    try {
      const ids = await markedIds(bytes);
      const paragraphs = elements.filter((e) => e.kind === "paragraph");
      assert.ok(paragraphs.length > 0);
      // The main part's paragraphs, in order; the pre-pass also marks the
      // header and footer paragraphs, which this module does not list.
      assert.deepEqual(
        paragraphs.map((element) => element.id.slice(2)),
        ids,
      );
      assert.ok(paragraphs.every((element) => element.textStyle!.fontSize > 0));
    } finally {
      await engine.dispose();
    }
  });
});
