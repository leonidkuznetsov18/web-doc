import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  generatedParagraphId,
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
 * Review of the pre-pass's coverage: the branches the task test does not
 * reach — a body without w:sectPr, sections without a content box, pictures
 * without a usable extent, paragraphs nested in tables, text boxes and both
 * branches of mc:AlternateContent, bookmark ids above and beside the file's
 * own, the endnotes story, missing targets and an aborted signal.
 */

const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const limits = defaultResourceLimits;
const NAME = /w:name="_wd([0-9A-F]{8})"/g;
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

async function partXml(
  bytes: Uint8Array,
  part = "/word/document.xml",
): Promise<string> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  return new TextDecoder().decode(await pkg.part(part));
}

/** The package with parts replaced or removed byte for byte, scanning nothing. */
async function withParts(
  bytes: Uint8Array,
  parts: Readonly<Record<string, string>>,
  remove: readonly string[] = [],
): Promise<Uint8Array> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  const set = new Map<string, Uint8Array>();
  for (const [name, data] of Object.entries(parts))
    set.set(name, new TextEncoder().encode(data));
  pkg.applyOverlay({ set, remove: new Set(remove) });
  return pkg.save();
}

function extents(xml: string): string[] {
  return [...xml.matchAll(/<wp:extent cx="(\d+)" cy="(\d+)"\/>/g)].map(
    (match) => `${match[1]}x${match[2]}`,
  );
}

describe("DOCX display pre-pass review (docx-engine-upgrade)", () => {
  it("falls back to Letter geometry without a body sectPr or usable page values, skips sections without a content box, and leaves aborted input alone", async () => {
    const oversized = `<w:p>${inlinePicture(19202400, 4800600)}</w:p>`;
    const media = [{ name: "word/media/image1.png", data: PNG }];

    const noSectPr = await prepareDocxForDisplay(
      buildDocx({ media, body: oversized }),
      limits,
    );
    assert.deepEqual(
      [noSectPr.scaledImages, noSectPr.markedParagraphs],
      [1, 1],
    );
    assert.deepEqual(extents(await partXml(noSectPr.bytes)), [
      "5943600x1485900",
    ]);

    // Missing and negative values fall back attribute by attribute.
    const partial = await prepareDocxForDisplay(
      buildDocx({
        media,
        body:
          oversized +
          '<w:sectPr><w:pgSz/><w:pgMar w:left="-5" w:right="abc"/></w:sectPr>',
      }),
      limits,
    );
    assert.equal(partial.scaledImages, 1);
    assert.deepEqual(extents(await partXml(partial.bytes)), [
      "5943600x1485900",
    ]);

    // Margins wider than the page: no content box, so nothing is scaled, but
    // the paragraph is still marked.
    const squeezed = await prepareDocxForDisplay(
      buildDocx({
        media,
        body: oversized + sectPr({ width: 2000, margin: 1440 }),
      }),
      limits,
    );
    assert.deepEqual(
      [squeezed.scaledImages, squeezed.markedParagraphs],
      [0, 1],
    );
    assert.deepEqual(extents(await partXml(squeezed.bytes)), [
      "19202400x4800600",
    ]);

    const input = buildDocx({ media, body: oversized });
    const aborted = await prepareDocxForDisplay(
      input,
      limits,
      AbortSignal.abort(),
    );
    assert.equal(aborted.bytes, input);
    assert.deepEqual([aborted.scaledImages, aborted.markedParagraphs], [0, 0]);
  });

  it("marks paragraphs nested in tables, text boxes and both AlternateContent branches in document order, scaling nested pictures", async () => {
    const textBox = (content: string) =>
      `<w:txbxContent><w:p><w:r><w:t>${content}</w:t></w:r></w:p></w:txbxContent>`;
    const bytes = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      rootAttributes:
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
        'xmlns:v="urn:schemas-microsoft-com:vml"',
      body:
        paragraph("Top") +
        "<w:tbl><w:tr><w:tc>" +
        paragraph("Outer") +
        `<w:tbl><w:tr><w:tc><w:p>${inlinePicture(19202400, 4800600)}</w:p></w:tc></w:tr></w:tbl>` +
        "<w:p/>" +
        "</w:tc></w:tr></w:tbl>" +
        "<w:p><w:r><mc:AlternateContent>" +
        '<mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="914400" cy="914400"/><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx>' +
        textBox("Box") +
        "</wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>" +
        "<mc:Fallback><w:pict><v:shape><v:textbox>" +
        textBox("Box") +
        "</v:textbox></v:shape></w:pict></mc:Fallback>" +
        "</mc:AlternateContent></w:r></w:p>" +
        sectPr(),
    });
    const result = await prepareDocxForDisplay(bytes, limits);
    assert.deepEqual(
      [result.scaledImages, result.markedParagraphs, result.generatedIds],
      [1, 7, 7],
    );
    const xml = await partXml(result.bytes);
    // The nested cell's picture is fitted to the body's section.
    assert.deepEqual(extents(xml), ["5943600x1485900", "914400x914400"]);
    // Ids follow the order of the paragraphs in the text: the paragraph that
    // holds the drawing is marked at its start, before the text boxes in it.
    const names = [...xml.matchAll(NAME)].map((match) => match[1]);
    assert.deepEqual(
      names,
      Array.from({ length: 7 }, (_, index) =>
        generatedParagraphId(index, new Set()),
      ),
    );
    const order = [
      "Top",
      "Outer",
      "<w:drawing><wp:inline",
      "<w:p><w:bookmarkStart",
      "<mc:AlternateContent>",
      "<wps:txbx><w:txbxContent><w:p><w:bookmarkStart",
      "<v:textbox><w:txbxContent><w:p><w:bookmarkStart",
    ];
    let cursor = 0;
    for (const marker of order) {
      const at = xml.indexOf(marker, cursor);
      assert.ok(at > cursor, `${marker} follows the previous paragraph`);
      cursor = at;
    }
    assert.ok(
      xml.includes(
        `<w:txbxContent><w:p><w:bookmarkStart w:id="7000005" w:name="_wd${names[5]}"/><w:bookmarkEnd w:id="7000005"/><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent>`,
      ),
      xml,
    );
    const ids = [...xml.matchAll(/<w:bookmarkStart w:id="(\d+)"/g)].map(
      (match) => Number(match[1]),
    );
    assert.deepEqual(
      ids,
      [7000000, 7000001, 7000002, 7000003, 7000004, 7000005, 7000006],
    );
  });

  it("continues bookmark ids above the highest numeric id, ignores non-numeric ones and avoids ids the file's own _wd bookmarks use", async () => {
    const first = generatedParagraphId(0, new Set());
    const bytes = buildDocx({
      body:
        `<w:p><w:bookmarkStart w:id="7500000" w:name="_wd${first.toLowerCase()}"/><w:bookmarkEnd w:id="7500000"/><w:r><w:t>A</w:t></w:r></w:p>` +
        '<w:p><w:bookmarkStart w:id="x9" w:name="_GoBack"/><w:bookmarkEnd w:id="x9"/><w:r><w:t>B</w:t></w:r></w:p>' +
        paragraph("C") +
        sectPr(),
    });
    const result = await prepareDocxForDisplay(bytes, limits);
    assert.deepEqual([result.markedParagraphs, result.generatedIds], [3, 3]);
    const xml = await partXml(result.bytes);
    const names = [...xml.matchAll(NAME)].map((match) => match[1]);
    const taken = new Set([first]);
    assert.deepEqual(names, [
      generatedParagraphId(0, taken),
      generatedParagraphId(1, taken),
      generatedParagraphId(2, taken),
    ]);
    assert.notEqual(names[0], first, "the lowercase name counts as taken");
    assert.ok(xml.includes(`w:name="_wd${first.toLowerCase()}"`), "kept");
    assert.ok(xml.includes('w:id="x9"'), "a non-numeric id is left alone");
    // A paragraph without w:pPr gets its pair first, before its own bookmarks.
    const ids = [...xml.matchAll(/<w:bookmarkStart w:id="([^"]+)"/g)].map(
      (match) => match[1],
    );
    assert.deepEqual(ids, ["7500001", "7500000", "7500002", "x9", "7500003"]);
    // Generated ids stay eight hex digits below 0x80000000, however many.
    for (const index of [0, 1, 1000, 200_000, 5_000_000]) {
      const id = generatedParagraphId(index, new Set());
      assert.match(id, /^[0-9A-F]{8}$/);
      assert.ok(Number.parseInt(id, 16) < 0x80000000);
    }
  });

  it("skips pictures without a usable extent, keeps a:ext only when it repeats the extent, marks endnotes and ignores missing parts", async () => {
    const media = [{ name: "word/media/image1.png", data: PNG }];
    const bytes = buildDocx({
      media,
      body:
        `<w:p>${inlinePicture(0, 4800600)}</w:p>` +
        `<w:p>${inlinePicture(19202400, 4800600).replace(/<wp:extent[^>]*\/>/, "")}</w:p>` +
        `<w:p>${inlinePicture(19202400, 4800600).replace(
          '<a:ext cx="19202400" cy="4800600"/>',
          '<a:ext cx="1234" cy="5678"/>',
        )}</w:p>` +
        sectPr(),
    });
    const result = await prepareDocxForDisplay(bytes, limits);
    assert.deepEqual([result.scaledImages, result.markedParagraphs], [1, 3]);
    const xml = await partXml(result.bytes);
    assert.deepEqual(extents(xml), ["0x4800600", "5943600x1485900"]);
    assert.ok(
      xml.includes('<a:ext cx="1234" cy="5678"/>'),
      "an a:ext that does not repeat the extent is left alone",
    );
    assert.ok(!xml.includes('<a:ext cx="5943600"'));

    // An endnotes part is a story; a header relationship whose part is
    // missing is skipped; neither needs a content type to be read.
    const original = buildDocx({ body: paragraph("Body") + sectPr() });
    const rels = (
      await partXml(original, "/word/_rels/document.xml.rels")
    ).replace(
      "</Relationships>",
      `<Relationship Id="rId8" Type="${R}/endnotes" Target="endnotes.xml"/>` +
        `<Relationship Id="rId9" Type="${R}/header" Target="header9.xml"/></Relationships>`,
    );
    const withEndnotes = await withParts(original, {
      "/word/_rels/document.xml.rels": rels,
      "/word/endnotes.xml": `${XML}<w:endnotes xmlns:w="${W}"><w:endnote w:id="1">${paragraph("Note")}${paragraph("More")}</w:endnote></w:endnotes>`,
    });
    const stories = await prepareDocxForDisplay(withEndnotes, limits);
    assert.deepEqual([stories.markedParagraphs, stories.generatedIds], [3, 3]);
    const endnotes = await partXml(stories.bytes, "/word/endnotes.xml");
    assert.equal([...endnotes.matchAll(NAME)].length, 2);
    assert.ok(
      endnotes.includes(
        `<w:endnote w:id="1"><w:p><w:bookmarkStart w:id="7000001" w:name="_wd${generatedParagraphId(1, new Set())}"/>`,
      ),
      "ids and bookmark numbers continue across the document's parts",
    );

    // Without the main part there is nothing to prepare.
    const headless = await withParts(original, {}, ["/word/document.xml"]);
    const same = await prepareDocxForDisplay(headless, limits);
    assert.equal(same.bytes, headless);
    assert.deepEqual([same.scaledImages, same.markedParagraphs], [0, 0]);
  });
});
