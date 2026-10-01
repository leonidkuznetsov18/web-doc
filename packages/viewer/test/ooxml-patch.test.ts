import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  applyPatches,
  patches,
  type XmlPatch,
} from "../src/edit/ooxml/patch.js";
import { scanXml } from "../src/edit/ooxml/xml.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { buildZip } from "./fixtures/zip-builder.js";

/*
 * Task 42 of the OOXML package layer: patch builders and their read-back
 * verification, atomic transactions across parts, snapshots, media
 * deduplication and dangling-relationship warnings.
 */

const limits = defaultResourceLimits;
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/></p:nvSpPr></p:sp></p:spTree></p:cSld></p:sld>`;

function code(error: unknown): string {
  return error instanceof ViewerError ? error.code : String(error);
}

function deck(): Uint8Array {
  return buildZip([
    {
      name: "[Content_Types].xml",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>',
    },
    {
      name: "_rels/.rels",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>',
    },
    {
      name: "ppt/presentation.xml",
      data: '<p:presentation xmlns:p="x" xmlns:r="y"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/></p:sldIdLst></p:presentation>',
    },
    {
      name: "ppt/_rels/presentation.xml.rels",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/></Relationships>',
    },
    { name: "ppt/slides/slide1.xml", data: SLIDE },
    { name: "ppt/slides/slide2.xml", data: SLIDE.replace("Hello", "Second") },
    {
      name: "ppt/slides/_rels/slide2.xml.rels",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    },
  ]);
}

describe("patches and transactions (ooxml package)", () => {
  it("builds patches whose read-back passes, for every builder", () => {
    const part = scanXml("/ppt/slides/slide1.xml", SLIDE);
    const text = part.find("a:t")!;
    const cNvPr = part.find("cNvPr")!;
    const spPr = part.find("p:spPr")!;
    const body = part.findAll("p:sp")[1]!;
    const items: XmlPatch[] = [
      patches.replaceContent(part, text, patches.text("Hi <there> & bye")),
      patches.setAttribute(part, cNvPr, "name", 'Title "A" & B'),
      patches.setAttribute(part, cNvPr, "hidden", "1"),
      patches.replaceContent(part, spPr, "<a:xfrm/>"),
      patches.insertAfter(part, body, "<p:pic><p:nvPicPr/></p:pic>"),
      patches.removeElement(part, body),
    ];
    const patched = applyPatches(part, items, 1);
    assert.equal(patched.revision, 1);
    assert.equal(patched.textOf(patched.find("a:t")!), "Hi <there> & bye");
    assert.equal(
      patched.attribute(patched.find("cNvPr")!, "name"),
      'Title "A" & B',
    );
    assert.equal(patched.attribute(patched.find("cNvPr")!, "hidden"), "1");
    assert.equal(patched.find("p:spPr")!.selfClosing, false);
    assert.equal(patched.find("a:xfrm")!.parent?.name, "p:spPr");
    assert.deepEqual(patched.findAll("p:sp").length, 1);
    assert.equal(patched.find("p:pic")!.parent?.name, "p:spTree");
    // Everything outside the ranges is byte for byte the original.
    assert.ok(
      patched.text.startsWith(SLIDE.slice(0, SLIDE.indexOf('name="Title 1"'))),
    );
    assert.ok(patched.text.endsWith("</p:spTree></p:cSld></p:sld>"));
    // Builders on the patched tree: append, insert before, remove attribute, replace element.
    const second = applyPatches(
      patched,
      [
        patches.appendChild(
          patched,
          patched.find("p:spTree")!,
          "<p:sp><p:nvSpPr/></p:sp>",
        ),
        patches.insertBefore(
          patched,
          patched.find("p:cSld")!,
          "<p:clrMapOvr/>",
        ),
        patches.removeAttribute(patched, patched.find("cNvPr")!, "hidden"),
        patches.replaceElement(
          patched,
          patched.find("a:bodyPr")!,
          '<a:bodyPr wrap="none"/>',
        ),
      ],
      2,
    );
    assert.equal(second.findAll("p:sp").length, 2);
    assert.equal(second.root.children[0]!.name, "p:clrMapOvr");
    assert.equal(second.attribute(second.find("cNvPr")!, "hidden"), undefined);
    assert.equal(second.attribute(second.find("a:bodyPr")!, "wrap"), "none");
    assert.equal(
      second.text.includes(" hidden="),
      false,
      "the attribute's whitespace went with it",
    );
  });

  it("refuses overlapping, out-of-range, malformed and mismatching patches", () => {
    const part = scanXml("/s.xml", SLIDE);
    const text = part.find("a:t")!;
    const bad = (items: XmlPatch[], expected: string, label: string): void => {
      try {
        applyPatches(part, items, 1);
      } catch (error) {
        assert.equal(
          code(error),
          "invalid-patch",
          `${label}: ${String(error)}`,
        );
        assert.match((error as Error).message, new RegExp(expected), label);
        return;
      }
      assert.fail(`${label}: no error`);
    };
    bad(
      [
        patches.replaceContent(part, text, "a"),
        patches.replaceElement(part, text, "<a:t>b</a:t>"),
      ],
      "overlap",
      "overlap",
    );
    bad(
      [{ start: 5, end: SLIDE.length + 1, text: "", expect: { kind: "none" } }],
      "outside",
      "out of range",
    );
    bad(
      [{ start: 9, end: 3, text: "", expect: { kind: "none" } }],
      "outside",
      "reversed",
    );
    bad(
      [patches.replaceContent(part, text, "<a:b>")],
      "well-formed",
      "malformed fragment",
    );
    bad(
      [patches.replaceElement(part, text, "<a:t>x</a:t><a:t>y</a:t>")],
      "one element",
      "two elements for one",
    );
    bad(
      [patches.insertAfter(part, text, "plain text")],
      "one element",
      "text where an element is expected",
    );
    const named = part.find("cNvPr")!;
    bad(
      [{ ...patches.setAttribute(part, named, "name", "x"), text: 'name="y"' }],
      "reads back",
      "attribute mismatch",
    );
    bad(
      [
        {
          ...patches.replaceContent(part, text, "x"),
          expect: { kind: "content", at: 7 },
        },
      ],
      "gone",
      "content of a missing element",
    );
  });

  it("commits patches, parts and media atomically, with snapshots and warnings", async () => {
    const pkg = await OoxmlPackage.open(deck(), { limits });
    const before = pkg.snapshot();
    const slide = await pkg.xml("/ppt/slides/slide1.xml");
    const transaction = pkg.transaction();
    transaction.patch(slide, [
      patches.replaceContent(
        slide,
        slide.find("a:t")!,
        patches.text("Patched"),
      ),
    ]);
    const media = await transaction.addMedia(
      "/ppt/slides/slide1.xml",
      "/ppt/media/",
      Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 1, 2, 3),
      "image/png",
    );
    assert.equal(media.part, "/ppt/media/image1.png");
    assert.equal(media.rId, "rId1");
    const again = await transaction.addMedia(
      "/ppt/slides/slide1.xml",
      "/ppt/media/",
      Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 1, 2, 3),
      "image/png",
    );
    assert.equal(again.part, media.part, "the same bytes are stored once");
    assert.equal(again.rId, "rId2", "but related again");
    const other = await transaction.addMedia(
      "/ppt/slides/slide1.xml",
      "/ppt/media/",
      Uint8Array.of(0xff, 0xd8, 0xff),
      "image/jpeg",
    );
    assert.equal(
      other.part,
      "/ppt/media/image1.jpeg",
      "numbered per extension",
    );
    assert.equal(
      transaction.uniquePartName("/ppt/slides/slide", ".xml"),
      "/ppt/slides/slide3.xml",
    );
    const change = await transaction.commit();
    assert.deepEqual([...change.changedParts].sort(), [
      "/[Content_Types].xml",
      "/ppt/slides/slide1.xml",
    ]);
    assert.deepEqual([...change.addedParts].sort(), [
      "/ppt/media/image1.jpeg",
      "/ppt/media/image1.png",
      "/ppt/slides/_rels/slide1.xml.rels",
    ]);
    assert.deepEqual(change.warnings, []);
    assert.equal(
      (await pkg.xml("/ppt/slides/slide1.xml")).textOf(
        (await pkg.xml("/ppt/slides/slide1.xml")).find("a:t")!,
      ),
      "Patched",
    );
    const types = decoder.decode(await pkg.part("/[Content_Types].xml"));
    assert.ok(
      types.includes('<Default Extension="png" ContentType="image/png"/>'),
    );
    assert.ok(
      types.includes('<Default Extension="jpeg" ContentType="image/jpeg"/>'),
    );
    const rels = await pkg.relationships("/ppt/slides/slide1.xml");
    assert.deepEqual(
      rels.items.map((item) => [item.id, item.targetPart]),
      [
        ["rId1", "/ppt/media/image1.png"],
        ["rId2", "/ppt/media/image1.png"],
        ["rId3", "/ppt/media/image1.jpeg"],
      ],
    );
    assert.ok(pkg.has("/ppt/media/image1.png"));
    // A stale scan cannot be patched; a committed transaction cannot be reused.
    assert.throws(
      () => pkg.transaction().patch(slide, []),
      (error: unknown) => code(error) === "invalid-patch",
    );
    assert.throws(
      () => transaction.removePart("/x"),
      (error: unknown) => code(error) === "lifecycle-error",
    );

    // Removing a slide: its own .rels goes too, the Override goes, the presentation's relationship dangles.
    const removal = pkg.transaction();
    removal.removePart("/ppt/slides/slide2.xml");
    const removed = await removal.commit();
    assert.deepEqual([...removed.removedParts].sort(), [
      "/ppt/slides/_rels/slide2.xml.rels",
      "/ppt/slides/slide2.xml",
    ]);
    assert.equal(removed.warnings.length, 1);
    assert.equal(removed.warnings[0]!.details?.reason, "dangling-relationship");
    assert.equal(removed.warnings[0]!.details?.source, "/ppt/presentation.xml");
    assert.equal(removed.warnings[0]!.details?.id, "rId2");
    assert.equal(
      decoder
        .decode(await pkg.part("/[Content_Types].xml"))
        .includes("slide2.xml"),
      false,
    );
    assert.equal(pkg.has("/ppt/slides/slide2.xml"), false);
    // Removing the relationship in the same transaction gives no warning.
    pkg.restore(before);
    const clean = pkg.transaction();
    clean.removePart("/ppt/slides/slide2.xml");
    clean.removeRelationship("/ppt/presentation.xml", "rId2");
    const presentation = await pkg.xml("/ppt/presentation.xml");
    clean.patch(presentation, [
      patches.removeElement(presentation, presentation.findAll("p:sldId")[1]!),
    ]);
    const cleaned = await clean.commit();
    assert.deepEqual(cleaned.warnings, []);
    assert.equal(
      (await pkg.relationships("/ppt/presentation.xml")).items.length,
      1,
    );
    assert.equal(
      (await pkg.xml("/ppt/presentation.xml")).findAll("p:sldId").length,
      1,
    );
    // The saved package reopens with the changes, and the original is untouched.
    const reopened = await OoxmlPackage.open(await pkg.save(), { limits });
    assert.equal(reopened.has("/ppt/slides/slide2.xml"), false);
    assert.ok(reopened.has("/ppt/slides/slide1.xml"));
    pkg.restore(before);
    assert.deepEqual(await pkg.save(), deck());
  });

  it("leaves the package unchanged when the last part of a transaction fails", async () => {
    const pkg = await OoxmlPackage.open(deck(), { limits });
    const slide1 = await pkg.xml("/ppt/slides/slide1.xml");
    const slide2 = await pkg.xml("/ppt/slides/slide2.xml");
    const transaction = pkg.transaction();
    transaction.patch(slide1, [
      patches.replaceContent(slide1, slide1.find("a:t")!, "ok"),
    ]);
    transaction.setPart(
      "/ppt/slides/slide9.xml",
      encoder.encode("<p:sld/>"),
      "application/xml",
    );
    transaction.patch(slide2, [
      patches.replaceContent(slide2, slide2.find("a:t")!, "<broken"),
    ]);
    await assert.rejects(
      transaction.commit(),
      (error: unknown) => code(error) === "invalid-patch",
    );
    assert.equal(pkg.revision, 0);
    assert.deepEqual(pkg.changedParts, []);
    assert.equal(pkg.has("/ppt/slides/slide9.xml"), false);
    assert.deepEqual(await pkg.save(), deck());
    // A new XML part that does not scan fails the same way.
    const bad = pkg.transaction();
    bad.setPart(
      "/ppt/slides/slide9.xml",
      encoder.encode("<p:sld>"),
      "application/xml",
    );
    await assert.rejects(
      bad.commit(),
      (error: unknown) => code(error) === "malformed-xml",
    );
    // A transaction started before another commit is a conflict.
    const early = pkg.transaction();
    const late = pkg.transaction();
    late.setPart("/ppt/x.xml", encoder.encode("<x/>"));
    await late.commit();
    early.setPart("/ppt/y.xml", encoder.encode("<y/>"));
    await assert.rejects(
      early.commit(),
      (error: unknown) => code(error) === "edit-conflict",
    );
  });
});
