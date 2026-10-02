import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import {
  generatedParagraphId,
  PARAGRAPH_BOOKMARK_PREFIX,
  prepareDocxForDisplay,
} from "../src/adapters/docx-prepass.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { defaultResourceLimits } from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";

/*
 * Task 51 of the DOCX engine upgrade: the display pre-pass scales
 * oversized inline pictures to their section's content box and marks every
 * paragraph with a hidden bookmark carrying its id (the file's w14:paraId
 * or a deterministic one), as patches of the XML the renderer loads; a
 * document it cannot read passes through.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const FIXTURE = new URL(
  "../../tests/fixtures/docx/oversized-inline-image.docx",
  PACKAGE_DIR,
);
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const limits = defaultResourceLimits;
const NAME = /w:name="_wd([0-9A-F]{8})"/g;

async function partXml(
  bytes: Uint8Array,
  part = "/word/document.xml",
): Promise<string> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  return new TextDecoder().decode(await pkg.part(part));
}

describe("DOCX display pre-pass (docx-engine-upgrade)", () => {
  it("scales oversized inline pictures to their section and leaves anchored ones alone", async () => {
    const bytes = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      body:
        `<w:p>${inlinePicture(19202400, 4800600)}</w:p>` +
        `<w:p>${inlinePicture(19202400, 4800600, "rId9", true)}</w:p>` +
        `<w:p>${inlinePicture(914400, 914400)}</w:p>` +
        `<w:tbl><w:tr><w:tc><w:p>${inlinePicture(9144000, 18288000)}</w:p></w:tc></w:tr></w:tbl>` +
        sectPr(),
    });
    const result = await prepareDocxForDisplay(bytes, limits);
    assert.equal(result.scaledImages, 2);
    const xml = await partXml(result.bytes);
    // 6.5 inches of content: 19202400 × (5943600 / 19202400) = 5943600, height in step.
    assert.ok(xml.includes('<wp:extent cx="5943600" cy="1485900"/>'), xml);
    assert.ok(xml.includes('<a:ext cx="5943600" cy="1485900"/>'));
    assert.ok(
      xml.includes("<wp:anchor") &&
        xml.includes('<wp:extent cx="19202400" cy="4800600"/>'),
      "the anchored picture keeps its extent",
    );
    assert.ok(
      xml.includes('<wp:extent cx="914400" cy="914400"/>'),
      "a fitting picture is untouched",
    );
    // 9 inches of content height: 18288000 × (8229600 / 18288000) = 8229600, width in step.
    assert.ok(
      xml.includes('<wp:extent cx="4114800" cy="8229600"/>'),
      "the table picture fits the height",
    );
  });

  it("uses the geometry of the section each paragraph belongs to", async () => {
    const bytes = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      body:
        `<w:p>${inlinePicture(9144000, 914400)}</w:p>` +
        `<w:p><w:pPr>${sectPr({ width: 15840, height: 12240, margin: 720 })}</w:pPr></w:p>` +
        `<w:p>${inlinePicture(9144000, 914400)}</w:p>` +
        sectPr({ width: 12240, height: 15840, margin: 1440 }),
    });
    const result = await prepareDocxForDisplay(bytes, limits);
    assert.equal(result.scaledImages, 1);
    const xml = await partXml(result.bytes);
    const extents = [
      ...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g),
    ].map((match) => match[1]);
    // The first section is landscape with half-inch margins: 14.1 inches of content, the picture fits.
    // The last section is portrait: 6.5 inches, so the second picture shrinks.
    assert.deepEqual(extents, ["9144000", "5943600"]);
  });

  it("marks every paragraph with an id bookmark, keeping authored ids and existing bookmarks", async () => {
    const bytes = buildDocx({
      body:
        paragraph("One") +
        '<w:sdt><w:sdtContent><w:p w14:paraId="1A00ABCD"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Two</w:t></w:r></w:p></w:sdtContent></w:sdt>' +
        `<w:tbl><w:tr><w:tc>${paragraph("Three")}</w:tc></w:tr></w:tbl>` +
        '<w:p><w:bookmarkStart w:id="7000005" w:name="_GoBack"/><w:bookmarkEnd w:id="7000005"/><w:r><w:t>Four</w:t></w:r></w:p>' +
        "<w:p/>" +
        "<w:p><w:pPr/></w:p>" +
        sectPr(),
      rootAttributes:
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
      header: paragraph("Header"),
      footer: paragraph("Footer", 'w14:paraId="00000001"'),
      footnotes: `<w:footnote w:id="1">${paragraph("Note")}</w:footnote>`,
    });
    const first = await prepareDocxForDisplay(bytes, limits);
    assert.equal(first.markedParagraphs, 9);
    assert.equal(first.generatedIds, 7);
    const xml = await partXml(first.bytes);
    const names = [...xml.matchAll(NAME)].map((match) => match[1]);
    assert.deepEqual(names, [
      generatedParagraphId(0, new Set()),
      "1A00ABCD",
      generatedParagraphId(1, new Set()),
      generatedParagraphId(2, new Set()),
      generatedParagraphId(3, new Set()),
      generatedParagraphId(4, new Set()),
    ]);
    assert.ok(names.every((name) => Number.parseInt(name!, 16) < 0x80000000));
    // After w:pPr where there is one, first otherwise; empty paragraphs are opened.
    assert.ok(
      xml.includes(
        `<w:pPr><w:jc w:val="center"/></w:pPr><w:bookmarkStart w:id="7000007" w:name="${PARAGRAPH_BOOKMARK_PREFIX}1A00ABCD"/><w:bookmarkEnd w:id="7000007"/><w:r><w:t>Two</w:t></w:r>`,
      ),
      xml,
    );
    assert.ok(
      /<w:p><w:bookmarkStart w:id="\d+" w:name="_wd[0-9A-F]{8}"\/><w:bookmarkEnd w:id="\d+"\/><\/w:p>/.test(
        xml,
      ),
      "an empty paragraph is opened for its bookmark",
    );
    assert.ok(/<w:p><w:pPr\/><w:bookmarkStart /.test(xml));
    const ids = [...xml.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map(
      (match) => Number(match[1]),
    );
    assert.ok(ids.includes(7000005), "the existing bookmark stays");
    assert.equal(new Set(ids).size, ids.length, "bookmark ids are unique");
    assert.ok(
      ids.filter((id) => id !== 7000005).every((id) => id > 7000005),
      "new ids are above the existing ones",
    );
    assert.ok(
      (await partXml(first.bytes, "/word/header1.xml")).includes('w:name="_wd'),
    );
    assert.ok(
      (await partXml(first.bytes, "/word/footnotes.xml")).includes(
        'w:name="_wd',
      ),
    );
    assert.ok(
      (await partXml(first.bytes, "/word/footer1.xml")).includes(
        'w:name="_wd00000001"',
      ),
      "an authored id names the bookmark",
    );
    // The same input gives the same output.
    const second = await prepareDocxForDisplay(bytes, limits);
    assert.deepEqual(second.bytes, first.bytes);
  });

  it("skips ids the file already uses", async () => {
    const bytes = buildDocx({
      body:
        paragraph(
          "Taken",
          `w14:paraId="${generatedParagraphId(0, new Set())}"`,
        ) +
        paragraph("Fresh") +
        sectPr(),
      rootAttributes:
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"',
    });
    const result = await prepareDocxForDisplay(bytes, limits);
    const names = [...(await partXml(result.bytes)).matchAll(NAME)].map(
      (match) => match[1],
    );
    assert.equal(names.length, 2);
    assert.notEqual(names[0], names[1]);
    const taken = new Set([generatedParagraphId(0, new Set())]);
    assert.equal(names[1], generatedParagraphId(0, taken));
  });

  it("returns the input when the document cannot be read", async () => {
    const garbage = new Uint8Array([1, 2, 3, 4, 5]);
    assert.equal((await prepareDocxForDisplay(garbage, limits)).bytes, garbage);
    const malformed = buildDocx({ body: "<w:p><w:r><w:t>Open" });
    assert.equal(
      (await prepareDocxForDisplay(malformed, limits)).bytes,
      malformed,
    );
    const empty = buildDocx({ body: sectPr() });
    const same = await prepareDocxForDisplay(empty, limits);
    assert.equal(same.bytes, empty);
    assert.deepEqual([same.scaledImages, same.markedParagraphs], [0, 0]);
  });

  it("fits the oversized-image fixture and marks its paragraph", async (t) => {
    if (!existsSync(FIXTURE)) {
      t.skip("fixture missing");
      return;
    }
    const bytes = new Uint8Array(readFileSync(FIXTURE));
    const result = await prepareDocxForDisplay(bytes, limits);
    assert.deepEqual(
      [result.scaledImages, result.markedParagraphs, result.generatedIds],
      [1, 1, 1],
    );
    const xml = await partXml(result.bytes);
    assert.ok(xml.includes('<wp:extent cx="5943600" cy="1485900"/>'));
    assert.ok(
      /<w:p><w:bookmarkStart w:id="7000000" w:name="_wd[0-9A-F]{8}"\/><w:bookmarkEnd w:id="7000000"\/><w:r><w:drawing>/.test(
        xml,
      ),
      xml.slice(0, 300),
    );
  });
});
