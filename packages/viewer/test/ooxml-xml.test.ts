import assert from "node:assert/strict";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { describe, it } from "node:test";
import { pathToFileURL } from "node:url";

import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import {
  decodeEntities,
  decodePart,
  encodePart,
  escapeAttribute,
  escapeText,
  scanXml,
  type XmlElement,
  type XmlPart,
} from "../src/edit/ooxml/xml.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import { minimalPackage } from "./fixtures/zip-builder.js";

/*
 * Task 41 of the OOXML package layer: the offset-preserving XML scanner on
 * hand-written parts with every construct the spec names, and on every XML
 * part of every corpus package.
 */

const PACKAGE_DIR = pathToFileURL(`${process.cwd()}/`);
const CORPUS = new URL("../../.cache/corpus/", PACKAGE_DIR);
const FIXTURES = new URL("../../tests/fixtures/", PACKAGE_DIR);
const limits = defaultResourceLimits;

const SLIDE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <!-- a comment between elements -->
  <p:cSld>
    <p:spTree>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name='Title &amp; "more"'/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>Hello &amp; &lt;world&gt; &#169; &#x263A;</a:t></a:r><a:br/><a:r><a:t><![CDATA[raw <cdata> & text]]></a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="p14"><p:sp/></mc:Choice><mc:Fallback><p:sp/></mc:Fallback></mc:AlternateContent>
      <?custom instruction?>
    </p:spTree>
  </p:cSld>
  <p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="123"/></p:ext></p:extLst>
</p:sld>
`;

function every(node: XmlElement, visit: (node: XmlElement) => void): void {
  visit(node);
  for (const child of node.children) every(child, visit);
}

/** Every element's range reproduces its source, nests inside its parent, and holds its children in order. */
function checkRanges(part: XmlPart): number {
  let count = 0;
  every(part.root, (node) => {
    count += 1;
    const source = part.text.slice(node.start, node.end);
    assert.ok(
      source.startsWith(`<${node.name}`),
      `${part.name}: ${node.name} starts its slice`,
    );
    assert.ok(
      node.selfClosing
        ? source.endsWith("/>")
        : source.endsWith(`</${node.name}>`) || source.endsWith(">"),
      `${part.name}: ${node.name} ends its slice`,
    );
    assert.ok(
      node.start <= node.contentStart &&
        node.contentStart <= node.contentEnd &&
        node.contentEnd <= node.end,
    );
    if (node.parent) {
      assert.ok(
        node.start >= node.parent.contentStart &&
          node.end <= node.parent.contentEnd,
        "inside the parent",
      );
      assert.equal(node.parent.children[node.index], node);
      assert.deepEqual(part.at(node.path), node);
    }
    let cursor = node.contentStart;
    for (const child of node.children) {
      assert.ok(child.start >= cursor, "children in order");
      cursor = child.end;
    }
    for (const attribute of node.attributes) {
      const slice = part.text.slice(attribute.start, attribute.end);
      assert.ok(
        slice.startsWith(attribute.name),
        `attribute ${attribute.name} slice`,
      );
      assert.ok(
        slice.endsWith(`${attribute.rawValue}"`) ||
          slice.endsWith(`${attribute.rawValue}'`),
      );
    }
  });
  return count;
}

describe("xml scanner (ooxml package)", () => {
  it("scans every construct with exact ranges, namespaces and decoded text", () => {
    const part = scanXml("/ppt/slides/slide1.xml", SLIDE);
    assert.equal(part.root.name, "p:sld");
    assert.equal(
      part.root.namespace,
      "http://schemas.openxmlformats.org/presentationml/2006/main",
    );
    assert.deepEqual(part.declaration, {
      kind: "declaration",
      start: 0,
      end: SLIDE.indexOf("?>") + 2,
    });
    assert.deepEqual(
      part.opaque.map((range) => range.kind),
      ["comment", "cdata", "pi"],
    );
    const count = checkRanges(part);
    assert.ok(count > 20, `${count} elements`);
    const title = part.find("cNvPr")!;
    assert.equal(part.attribute(title, "name"), 'Title & "more"');
    assert.equal(title.attributes[1]!.rawValue, 'Title &amp; "more"');
    assert.equal(title.selfClosing, true);
    assert.equal(title.contentStart, title.contentEnd);
    const paragraph = part.find("a:p")!;
    assert.equal(
      part.textOf(paragraph),
      "Hello & <world> © ☺raw <cdata> & text",
    );
    const run = part.findAll("a:t");
    assert.equal(run.length, 2);
    assert.equal(part.textOf(run[1]!), "raw <cdata> & text");
    const choice = part.find("mc:Choice")!;
    assert.equal(
      choice.namespace,
      "http://schemas.openxmlformats.org/markup-compatibility/2006",
    );
    assert.equal(choice.parent?.name, "mc:AlternateContent");
    const creation = part.find("p14:creationId")!;
    assert.equal(
      creation.namespace,
      "http://schemas.microsoft.com/office/powerpoint/2010/main",
    );
    assert.equal(part.find("nope"), undefined);
    assert.equal(part.at([0, 0, 0, 0]), part.find("nvSpPr"));
    assert.equal(part.at([9]), undefined);
    // Local-name search matches any prefix; qualified search is exact.
    assert.equal(part.findAll("sp").length, 3);
    assert.equal(part.findAll("p:sp").length, 3);
    assert.equal(part.findAll("a:sp").length, 0);
  });

  it("decodes entities and escapes text and attributes", () => {
    assert.equal(
      decodeEntities(
        "a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos; &#65; &#x42; &unknown; &#xFFFFFFFF;",
      ),
      "a & b < c > d \"e\" 'f' A B &unknown; &#xFFFFFFFF;",
    );
    assert.equal(
      escapeText('<a href="x">&</a>'),
      '&lt;a href="x"&gt;&amp;&lt;/a&gt;',
    );
    assert.equal(
      escapeAttribute('say "hi" & <bye>\n'),
      '"say &quot;hi&quot; &amp; &lt;bye>&#10;"',
    );
    assert.equal(
      decodeEntities(escapeText("round <&> trip")),
      "round <&> trip",
    );
  });

  it("keeps a byte-order mark and refuses parts that are not UTF-8", () => {
    const withBom = new Uint8Array([
      0xef,
      0xbb,
      0xbf,
      ...new TextEncoder().encode("<a/>"),
    ]);
    const decoded = decodePart("/a.xml", withBom);
    assert.equal(decoded.hasBom, true);
    assert.deepEqual(
      [...encodePart(decoded.text)],
      [...withBom],
      "re-encoding gives the original bytes back",
    );
    const part = scanXml("/a.xml", decoded.text);
    assert.equal(part.hasBom, true);
    assert.equal(part.root.start, 1);
    assert.throws(
      () =>
        decodePart(
          "/latin1.xml",
          Uint8Array.of(0x3c, 0x61, 0x3e, 0xe9, 0x3c, 0x2f, 0x61, 0x3e),
        ),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "unsupported-part",
    );
  });

  it("reports malformed parts with the offset", () => {
    const cases: [string, RegExp][] = [
      ["<a><b></a>", /closes <b>/],
      ["<a>", /never closed/],
      ["<a></a><b/>", /second root/],
      ["text<a/>", /outside the root/],
      ["<a b></a>", /no value/],
      ["<a b=c></a>", /not quoted/],
      ['<a b="<"></a>', /contains "<"/],
      ["<a><!-- open </a>", /unterminated comment/],
      ["<a><![CDATA[x</a>", /unterminated CDATA/],
      ["</a>", /without a start tag/],
      ["<a/><?xml version='1.0'?>", /declaration after/],
      ["", /no root/],
      ["<a / ></a>", /stray/],
    ];
    for (const [text, pattern] of cases)
      assert.throws(
        () => scanXml("/bad.xml", text),
        (error: unknown) =>
          error instanceof ViewerError &&
          error.code === "malformed-xml" &&
          pattern.test(error.message) &&
          typeof error.details?.offset === "number",
        `${JSON.stringify(text)} → ${pattern}`,
      );
    // A DOCTYPE with an internal subset is opaque, not an error.
    const doctyped = scanXml(
      "/d.xml",
      '<!DOCTYPE a [<!ENTITY x "y">]><a>&x;</a>',
    );
    assert.equal(doctyped.opaque[0]!.kind, "doctype");
    assert.equal(
      doctyped.textOf(doctyped.root),
      "&x;",
      "an undeclared entity stays as text",
    );
  });

  it("serves scanned parts from the package and rescans after a change", async () => {
    const bytes = minimalPackage([
      { name: "ppt/slides/slide1.xml", data: SLIDE, method: 8 },
    ]);
    const pkg = await OoxmlPackage.open(bytes, { limits });
    const first = await pkg.xml("/ppt/slides/slide1.xml");
    assert.equal(first.revision, 0);
    assert.equal(await pkg.xml("/ppt/slides/slide1.xml"), first, "cached");
    pkg.applyOverlay({
      set: new Map([
        [
          "/ppt/slides/slide1.xml",
          new TextEncoder().encode("<p:sld xmlns:p='x'><p:cSld/></p:sld>"),
        ],
      ]),
      remove: new Set(),
    });
    const second = await pkg.xml("/ppt/slides/slide1.xml");
    assert.equal(second.revision, 1);
    assert.equal(second.root.children.length, 1);
    await assert.rejects(
      pkg.xml("/missing.xml"),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "invalid-file",
    );
    pkg.applyOverlay({
      set: new Map([
        ["/ppt/slides/slide1.xml", Uint8Array.of(0xff, 0xfe, 0x3c)],
      ]),
      remove: new Set(),
    });
    await assert.rejects(
      pkg.xml("/ppt/slides/slide1.xml"),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "unsupported-part",
    );
  });

  it("scans every XML part of every corpus package with ranges that reproduce the source", async (t) => {
    const files: URL[] = [];
    if (existsSync(CORPUS))
      for (const name of readdirSync(CORPUS))
        if (/\.(pptx|docx)$/i.test(name)) files.push(new URL(name, CORPUS));
    for (const folder of ["pptx/", "docx/"]) {
      const dir = new URL(folder, FIXTURES);
      if (existsSync(dir))
        for (const name of readdirSync(dir))
          if (/\.(pptx|docx|pptm|docm)$/i.test(name))
            files.push(new URL(name, dir));
    }
    if (files.length === 0) {
      t.skip("no packages; run npm run corpus:fetch");
      return;
    }
    for (const file of files) {
      const pkg = await OoxmlPackage.open(new Uint8Array(readFileSync(file)), {
        limits,
      });
      let parts = 0;
      let elements = 0;
      let largest = { name: "", chars: 0, ms: 0 };
      for (const name of pkg.partNames) {
        if (!/\.(xml|rels)$/i.test(name)) continue;
        const bytes = await pkg.part(name);
        const started = performance.now();
        const part = await pkg.xml(name);
        const ms = performance.now() - started;
        assert.deepEqual(
          [...encodePart(part.text)],
          [...bytes],
          `${name}: text re-encodes to the bytes`,
        );
        elements += checkRanges(part);
        parts += 1;
        if (part.text.length > largest.chars)
          largest = { name, chars: part.text.length, ms };
      }
      console.log(
        `scanner ${file.pathname.split("/").pop()}: ${parts} XML parts, ${elements} elements; largest ${largest.name} ${largest.chars} chars in ${largest.ms.toFixed(2)} ms`,
      );
      assert.ok(parts > 10);
    }
  });
});
