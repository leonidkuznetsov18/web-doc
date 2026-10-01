import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  extensionForMime,
  parseContentTypes,
  parseRelationships,
  sourceOfRelationshipsPart,
} from "../src/edit/ooxml/opc.js";
import { relativeTarget } from "../src/edit/ooxml/transaction.js";
import { scanXml } from "../src/edit/ooxml/xml.js";
import { defaultResourceLimits } from "../src/index.js";
import { buildZip } from "./fixtures/zip-builder.js";

/*
 * Task 40 of the OOXML package layer: content types and relationships as
 * models over scanned parts, target resolution and id allocation, and how a
 * transaction writes them back as patches that keep untouched bytes.
 */

const limits = defaultResourceLimits;
const decoder = new TextDecoder();
const encoder = new TextEncoder();

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="PNG" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>`;

const SLIDE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/x" TargetMode="External"/></Relationships>`;

function deck(): Uint8Array {
  return buildZip([
    { name: "[Content_Types].xml", data: CONTENT_TYPES },
    { name: "_rels/.rels", data: ROOT_RELS },
    { name: "ppt/presentation.xml", data: '<p:presentation xmlns:p="x"/>' },
    {
      name: "ppt/slides/slide1.xml",
      data: '<p:sld xmlns:p="x"><p:cSld/></p:sld>',
    },
    { name: "ppt/slides/_rels/slide1.xml.rels", data: SLIDE_RELS },
    {
      name: "ppt/slideLayouts/slideLayout1.xml",
      data: '<p:sldLayout xmlns:p="x"/>',
    },
    {
      name: "ppt/media/image1.png",
      data: Uint8Array.of(0x89, 0x50, 0x4e, 0x47),
      method: 0,
    },
  ]);
}

describe("opc model (ooxml package)", () => {
  it("reads content types by override and by extension, case-insensitively", () => {
    const types = parseContentTypes(
      scanXml("/[Content_Types].xml", CONTENT_TYPES),
    );
    assert.equal(
      types.typeOf("/ppt/slides/slide1.xml"),
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    assert.equal(
      types.typeOf("/PPT/SLIDES/SLIDE1.XML"),
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    assert.equal(
      types.typeOf("/ppt/slideLayouts/slideLayout1.xml"),
      "application/xml",
    );
    assert.equal(types.typeOf("/ppt/media/image1.png"), "image/png");
    assert.equal(types.typeOf("/ppt/media/image1.Png"), "image/png");
    assert.equal(types.typeOf("/ppt/media/movie.mp4"), undefined);
    assert.equal(
      types.overrideNode("/ppt/presentation.xml")?.local,
      "Override",
    );
    assert.equal(types.defaultNode("png")?.local, "Default");
  });

  it("reads relationships with resolved targets, indexes and the next free id", () => {
    const rels = parseRelationships(
      "/ppt/slides/slide1.xml",
      "/ppt/slides/_rels/slide1.xml.rels",
      scanXml("/ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS),
    );
    assert.deepEqual(
      rels.items.map((item) => item.id),
      ["rId3", "rId1", "rId2"],
      "document order",
    );
    assert.equal(rels.byId("rId3")?.targetPart, "/ppt/media/image1.png");
    assert.equal(
      rels.byId("rId1")?.targetPart,
      "/ppt/slideLayouts/slideLayout1.xml",
    );
    assert.equal(rels.byId("rId2")?.targetMode, "External");
    assert.equal(rels.byId("rId2")?.targetPart, undefined);
    assert.equal(
      rels.byType(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
      ).length,
      1,
    );
    assert.equal(rels.nextId(), "rId4", "the smallest free id");
    assert.equal(rels.nextId(new Set(["rId4", "rId5"])), "rId6");
    const empty = parseRelationships("/x.xml", undefined, undefined);
    assert.equal(empty.nextId(), "rId1");
    const root = parseRelationships(
      "/",
      "/_rels/.rels",
      scanXml("/_rels/.rels", ROOT_RELS),
    );
    assert.equal(root.byId("rId1")?.targetPart, "/ppt/presentation.xml");
  });

  it("maps .rels parts back to their sources and media types to extensions", () => {
    assert.equal(sourceOfRelationshipsPart("/_rels/.rels"), "/");
    assert.equal(
      sourceOfRelationshipsPart("/ppt/slides/_rels/slide1.xml.rels"),
      "/ppt/slides/slide1.xml",
    );
    assert.equal(
      sourceOfRelationshipsPart("/ppt/slides/slide1.xml"),
      undefined,
    );
    assert.equal(extensionForMime("image/png"), "png");
    assert.equal(extensionForMime("image/JPEG"), "jpeg");
    assert.equal(extensionForMime("image/svg+xml"), "svg");
    assert.equal(extensionForMime("application/octet-stream"), "bin");
    assert.equal(
      relativeTarget("/ppt/slides/slide1.xml", "/ppt/media/image1.png"),
      "../media/image1.png",
    );
    assert.equal(
      relativeTarget("/ppt/slides/slide1.xml", "/ppt/slides/slide2.xml"),
      "slide2.xml",
    );
    assert.equal(
      relativeTarget("/", "/ppt/presentation.xml"),
      "ppt/presentation.xml",
    );
    assert.equal(
      relativeTarget("/word/document.xml", "/word/media/image1.png"),
      "media/image1.png",
    );
  });

  it("serves content types and relationships from the package and resolves targets", async () => {
    const pkg = await OoxmlPackage.open(deck(), { limits });
    assert.equal(
      await pkg.contentTypeOf("/ppt/slides/slide1.xml"),
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    const rels = await pkg.relationships("/ppt/slides/slide1.xml");
    assert.equal(rels.partName, "/ppt/slides/_rels/slide1.xml.rels");
    assert.equal(rels.items.length, 3);
    assert.equal(
      (await pkg.relationships("/")).byId("rId1")?.targetPart,
      "/ppt/presentation.xml",
    );
    assert.equal(
      (await pkg.relationships("/ppt/slideLayouts/slideLayout1.xml")).items
        .length,
      0,
      "a part without a .rels",
    );
    assert.equal(
      pkg.resolve("/ppt/slides/slide1.xml", "../media/image1.png"),
      "/ppt/media/image1.png",
    );
  });

  it("writes relationship and content-type changes as patches that keep the rest", async () => {
    const pkg = await OoxmlPackage.open(deck(), { limits });
    const transaction = pkg.transaction();
    const id = await transaction.addRelationship(
      "/ppt/slides/slide1.xml",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
      "../notesSlides/notesSlide1.xml",
    );
    assert.equal(id, "rId4");
    const second = await transaction.addRelationship(
      "/ppt/slides/slide1.xml",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
      "https://example.com/y",
      "External",
    );
    assert.equal(second, "rId5");
    transaction.removeRelationship("/ppt/slides/slide1.xml", "rId2");
    transaction.removeRelationship("/ppt/slides/slide1.xml", "rId5");
    transaction.setPart(
      "/ppt/notesSlides/notesSlide1.xml",
      encoder.encode('<p:notes xmlns:p="x"/>'),
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    );
    transaction.setPart(
      "/ppt/slides/slide2.xml",
      encoder.encode('<p:sld xmlns:p="x"/>'),
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    const first = await transaction.addRelationship(
      "/ppt/slides/slide2.xml",
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
      "../slideLayouts/slideLayout1.xml",
    );
    assert.equal(first, "rId1", "a new .rels part starts at rId1");
    const change = await transaction.commit();
    assert.deepEqual([...change.addedParts].sort(), [
      "/ppt/notesSlides/notesSlide1.xml",
      "/ppt/slides/_rels/slide2.xml.rels",
      "/ppt/slides/slide2.xml",
    ]);
    assert.deepEqual([...change.changedParts].sort(), [
      "/[Content_Types].xml",
      "/ppt/slides/_rels/slide1.xml.rels",
    ]);
    assert.deepEqual(change.warnings, []);

    const rels = decoder.decode(
      await pkg.part("/ppt/slides/_rels/slide1.xml.rels"),
    );
    assert.ok(
      rels.startsWith(
        SLIDE_RELS.slice(0, SLIDE_RELS.indexOf('<Relationship Id="rId2"')),
      ),
      "the untouched relationships keep their bytes",
    );
    assert.equal(rels.includes('Id="rId2"'), false, "removed");
    assert.ok(
      rels.includes(
        '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>',
      ),
      "appended before the end tag",
    );
    assert.equal(
      rels.includes("rId5"),
      false,
      "added then removed in the same transaction",
    );
    const reread = await pkg.relationships("/ppt/slides/slide1.xml");
    assert.deepEqual(
      reread.items.map((item) => item.id),
      ["rId3", "rId1", "rId4"],
    );

    const types = decoder.decode(await pkg.part("/[Content_Types].xml"));
    assert.ok(
      types.startsWith(
        CONTENT_TYPES.slice(0, CONTENT_TYPES.indexOf("</Types>")),
      ),
      "the existing content types keep their bytes",
    );
    assert.ok(
      types.includes('<Override PartName="/ppt/notesSlides/notesSlide1.xml"'),
    );
    assert.ok(types.includes('<Override PartName="/ppt/slides/slide2.xml"'));
    assert.equal(
      await pkg.contentTypeOf("/ppt/slides/slide2.xml"),
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
    );
    const newRels = decoder.decode(
      await pkg.part("/ppt/slides/_rels/slide2.xml.rels"),
    );
    assert.ok(
      newRels.includes('Id="rId1"') && newRels.includes("slideLayout1.xml"),
    );
    // A part whose extension already has the right Default gets no Override.
    const again = pkg.transaction();
    again.setPart("/ppt/media/image2.png", Uint8Array.of(1), "image/png");
    await again.commit();
    assert.equal(
      decoder
        .decode(await pkg.part("/[Content_Types].xml"))
        .includes("image2.png"),
      false,
    );
  });
});
