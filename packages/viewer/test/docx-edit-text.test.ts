import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { prepareDocxForDisplay } from "../src/adapters/docx-prepass.js";
import { DocxEditEngine } from "../src/edit/docx/engine.js";
import type { DocxOperation } from "../src/edit/docx/types.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { parseZip } from "../src/edit/ooxml/zip.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";
import { docxSession, run } from "./fixtures/docx-session.js";

/*
 * Task 55 of the DOCX module: replaceText, setTextStyle and
 * setParagraphStyle rebuild one paragraph at a time, keep every byte the
 * edit does not touch, write the paragraph's id, split paragraphs on
 * newlines, refuse ranges that cut fields and wrapper edges, and keep
 * element ids stable across edits, shown copies and replays.
 */

const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const W14 = 'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

async function open(bytes: Uint8Array): Promise<DocxEditEngine> {
  return DocxEditEngine.open(bytes, limits, signal);
}

async function documentXml(bytes: Uint8Array): Promise<string> {
  const pkg = await OoxmlPackage.open(bytes, { limits });
  return new TextDecoder().decode(await pkg.part("/word/document.xml"));
}

/** The body of the main part, as one string. */
async function bodyOf(engine: DocxEditEngine): Promise<string> {
  const xml = await documentXml(await engine.materialize("save", {}, signal));
  return xml.slice(xml.indexOf("<w:body>") + 8, xml.indexOf("</w:body>"));
}

function ids(engine: DocxEditEngine): Promise<string[]> {
  return engine
    .getElements({}, signal)
    .then((elements) => elements.map((element) => element.id));
}

async function apply(
  engine: DocxEditEngine,
  operations: readonly DocxOperation[],
) {
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

/** Names of the ZIP entries whose bytes differ between two packages. */
function changedEntries(a: Uint8Array, b: Uint8Array): string[] {
  const left = parseZip(a, limits);
  const right = parseZip(b, limits);
  const changed: string[] = [];
  for (const entry of right.entries) {
    const before = left.entries.find(
      (candidate) => candidate.name === entry.name,
    );
    if (
      !before ||
      before.crc32 !== entry.crc32 ||
      before.uncompressedSize !== entry.uncompressedSize
    )
      changed.push(entry.name);
  }
  return changed;
}

const RUN = (text: string, rPr = ""): string =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

describe("DOCX replaceText (docx-edit T55)", () => {
  it("rebuilds the paragraph it targets, keeps the rest byte for byte and writes the id", async () => {
    const original = buildDocx({
      body:
        `<w:p w:rsidR="00AA"><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr>${RUN("Hello ", "<w:rPr><w:i/></w:rPr>")}${RUN("world")}</w:p>` +
        paragraph("Untouched") +
        sectPr(),
    });
    const engine = await open(original);
    const [first, second] = await ids(engine);
    const change = await apply(engine, [
      { op: "replaceText", target: first!, text: "Bye\tnow" },
    ]);
    assert.deepEqual(change.createdIds, []);
    assert.equal(change.reflowFrom, first!.slice(2));
    const body = await bodyOf(engine);
    assert.equal(
      body,
      `<w:p w:rsidR="00AA" w14:paraId="${first!.slice(2)}"><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>Bye</w:t><w:tab/><w:t>now</w:t></w:r></w:p>` +
        paragraph("Untouched") +
        sectPr(),
    );
    const saved = await engine.materialize("save", {}, signal);
    const xml = await documentXml(saved);
    assert.match(
      xml,
      /xmlns:w14="http:\/\/schemas\.microsoft\.com\/office\/word\/2010\/wordml"/,
    );
    assert.match(xml, /mc:Ignorable="w14"/);
    assert.deepEqual(changedEntries(original, saved), ["word/document.xml"]);
    assert.deepEqual(await ids(engine), [first, second]);
    const [element] = await engine.getElements({}, signal);
    assert.equal(element!.text, "Bye\tnow");
    assert.equal(element!.textStyle!.italic, true);
    await engine.dispose();
  });

  it("replaces a range inside a run, at a caret, at the start and in an empty paragraph", async () => {
    const original = buildDocx({
      body:
        `<w:p>${RUN("abc", "<w:rPr><w:b/></w:rPr>")}${RUN("def", "<w:rPr><w:i/></w:rPr>")}</w:p>` +
        `<w:p><w:pPr><w:rPr><w:u w:val="single"/></w:rPr></w:pPr></w:p>` +
        sectPr(),
    });
    const engine = await open(original);
    const [p1, p2] = await ids(engine);
    const range = (start: number, end: number) => ({
      start: { elementId: p1!, offset: start },
      end: { elementId: p1!, offset: end },
    });
    await apply(engine, [
      { op: "replaceText", target: p1!, text: "XY", range: range(1, 2) },
    ]);
    assert.match(
      await bodyOf(engine),
      /<w:r><w:rPr><w:b\/><\/w:rPr><w:t>a<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>XY<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>c<\/w:t><\/w:r><w:r><w:rPr><w:i\/><\/w:rPr><w:t xml:space="preserve">def<\/w:t><\/w:r>/,
    );
    // A caret after the bold run takes the bold style; at the start, the run after.
    await apply(engine, [
      { op: "replaceText", target: p1!, text: "+", range: range(4, 4) },
      { op: "replaceText", target: p1!, text: "^", range: range(0, 0) },
    ]);
    const text = (await engine.getElement(p1!, signal))!.text;
    assert.equal(text, "^aXYc+def");
    const body = await bodyOf(engine);
    assert.match(
      body,
      /<w:p[^>]*><w:r><w:rPr><w:b\/><\/w:rPr><w:t>\^<\/w:t><\/w:r>/,
    );
    assert.match(
      body,
      /<w:t>c<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>\+<\/w:t><\/w:r><w:r><w:rPr><w:i\/>/,
    );
    // The empty paragraph styles new text like its paragraph mark.
    await apply(engine, [{ op: "replaceText", target: p2!, text: "new" }]);
    assert.match(
      await bodyOf(engine),
      /<w:pPr><w:rPr><w:u w:val="single"\/><\/w:rPr><\/w:pPr><w:r><w:rPr><w:u w:val="single"\/><\/w:rPr><w:t>new<\/w:t><\/w:r><\/w:p>/,
    );
    await engine.dispose();
  });

  it("keeps bookmarks, enters a hyperlink, removes covered fields and pictures, refuses cuts", async () => {
    const original = buildDocx({
      media: [{ name: "word/media/image1.png", data: PNG }],
      body:
        `<w:p><w:bookmarkStart w:id="1" w:name="mark"/>${RUN("one ")}<w:hyperlink r:id="rId9">${RUN("link", '<w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr>')}</w:hyperlink>` +
        `<w:bookmarkEnd w:id="1"/>${RUN(" two ")}` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${RUN("77")}<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `${RUN(" three ")}${inlinePicture(914400, 914400)}${RUN(" four")}</w:p>` +
        sectPr(),
    });
    const engine = await open(original);
    const [p] = await ids(engine);
    const text = (await engine.getElement(p!, signal))!.text;
    assert.equal(text, "one link two 77 three ￼ four");
    const range = (start: number, end: number) => ({
      start: { elementId: p!, offset: start },
      end: { elementId: p!, offset: end },
    });
    // The field is 13..15, the picture 22..23, the link 4..8.
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "x", range: range(12, 14) },
      "invalid-range",
    );
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "x", range: range(2, 6) },
      "invalid-range",
    );
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "a\nb", range: range(5, 6) },
      "invalid-range",
      "/text",
    );
    // Inside the link: the link stays and its text changes.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "LINK", range: range(5, 7) },
    ]);
    let body = await bodyOf(engine);
    assert.match(
      body,
      /<w:hyperlink r:id="rId9"><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>l<\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>LINK<\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="Hyperlink"\/><\/w:rPr><w:t>k<\/w:t><\/w:r><\/w:hyperlink>/,
    );
    assert.ok(body.includes('<w:bookmarkStart w:id="1" w:name="mark"/>'));
    assert.ok(body.includes('<w:bookmarkEnd w:id="1"/>'));
    // Covering the field, the picture and the link whole removes them.
    const now = (await engine.getElement(p!, signal))!.text;
    assert.equal(now, "one lLINKk two 77 three ￼ four");
    const change = await apply(engine, [
      {
        op: "replaceText",
        target: p!,
        text: "-",
        range: range(now.indexOf("7"), now.indexOf("four")),
      },
    ]);
    assert.deepEqual(change.removedIds, [`img:${p!.slice(2)}.0`]);
    body = await bodyOf(engine);
    assert.ok(!body.includes("fldChar"));
    assert.ok(!body.includes("w:drawing"));
    assert.equal(
      (await engine.getElement(p!, signal))!.text,
      "one lLINKk two -four",
    );
    await apply(engine, [
      { op: "replaceText", target: p!, text: "", range: range(3, 10) },
    ]);
    body = await bodyOf(engine);
    assert.ok(!body.includes("w:hyperlink"));
    assert.equal((await engine.getElement(p!, signal))!.text, "one two -four");
    assert.ok(body.includes('<w:bookmarkEnd w:id="1"/>'));
    await engine.dispose();
  });

  it("splits paragraphs on newlines with fresh ids, copied properties and a moved tail", async () => {
    const original = buildDocx({
      rootAttributes: W14,
      body:
        `<w:p w14:paraId="1A00AAAA"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${RUN("first ", "<w:rPr><w:b/></w:rPr>")}${RUN("second")}</w:p>` +
        paragraph("after") +
        sectPr(),
    });
    const engine = await open(original);
    const [p, after] = await ids(engine);
    const change = await apply(engine, [
      {
        op: "replaceText",
        target: p!,
        text: "A\nB\nC",
        range: {
          start: { elementId: p!, offset: 3 },
          end: { elementId: p!, offset: 8 },
        },
      },
    ]);
    assert.equal(change.createdIds.length, 2);
    const [b, c] = change.createdIds;
    assert.deepEqual(await ids(engine), [p, b, c, after]);
    const texts = (await engine.getElements({}, signal)).map((e) => e.text);
    assert.deepEqual(texts, ["firA", "B", "Ccond", "after"]);
    const body = await bodyOf(engine);
    const expectedB = `<w:p w14:paraId="${b!.slice(2)}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>B</w:t></w:r></w:p>`;
    assert.ok(body.includes(expectedB), body);
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${c!.slice(2)}"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>C</w:t></w:r><w:r><w:t>cond</w:t></w:r></w:p>`,
      ),
      body,
    );
    // The same batches replayed from the original issue the same ids.
    const replay: DocxOperation = {
      op: "replaceText",
      target: p!,
      text: "A\nB\nC",
      range: {
        start: { elementId: p!, offset: 3 },
        end: { elementId: p!, offset: 8 },
      },
    };
    await engine.restore([[replay]], signal);
    assert.deepEqual(await ids(engine), [p, b, c, after]);
    await engine.dispose();
  });

  it("keeps the ids of untouched paragraphs stable and stamps them only in the shown copy", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + paragraph("three") + sectPr(),
    });
    const engine = await open(original);
    const before = await ids(engine);
    await apply(engine, [
      { op: "replaceText", target: before[0]!, text: "1\n1b" },
    ]);
    const after = await ids(engine);
    assert.deepEqual(after.slice(2), before.slice(1));
    const saved = await documentXml(
      await engine.materialize("save", {}, signal),
    );
    const shown = await documentXml(
      await engine.materialize("show", {}, signal),
    );
    assert.equal((saved.match(/w14:paraId=/g) ?? []).length, 2);
    assert.equal((shown.match(/w14:paraId=/g) ?? []).length, 4);
    for (const id of after)
      assert.ok(shown.includes(`w14:paraId="${id.slice(2)}"`), id);
    // The pre-pass reads the stamped ids back, so runs will name the engine's paragraphs.
    const display = await prepareDocxForDisplay(
      await engine.materialize("show", {}, signal),
      limits,
    );
    const marked = [
      ...(await documentXml(display.bytes)).matchAll(/_wd([0-9A-F]{8})/g),
    ].map((m) => m[1]);
    assert.deepEqual(
      marked,
      after.map((id) => id.slice(2)),
    );
    // A shown copy opened as a base keeps every id.
    await engine.restore(
      { base: await engine.materialize("show", {}, signal), batches: [] },
      signal,
    );
    assert.deepEqual(await ids(engine), after);
    await engine.dispose();
  });

  it("refuses bad text, unknown and read-only targets and bad ranges", async () => {
    const engine = await open(
      buildDocx({
        body:
          paragraph("ok") +
          `<w:p><w:pPr>${sectPr()}</w:pPr>${RUN("section")}</w:p>` +
          sectPr(),
      }),
    );
    const [p, readOnly] = await ids(engine);
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "a\u0000b" },
      "invalid-text",
      "/text",
    );
    await expectIssue(
      engine,
      { op: "replaceText", target: "p:00000001", text: "x" },
      "unknown-target",
    );
    await expectIssue(
      engine,
      { op: "replaceText", target: readOnly!, text: "x" },
      "invalid-target",
    );
    await expectIssue(
      engine,
      {
        op: "replaceText",
        target: p!,
        text: "x",
        range: {
          start: { elementId: p!, offset: 1 },
          end: { elementId: p!, offset: 9 },
        },
      },
      "invalid-range",
    );
    await expectIssue(
      engine,
      {
        op: "replaceText",
        target: p!,
        text: "x",
        range: {
          start: { elementId: "p:00000001", offset: 0 },
          end: { elementId: p!, offset: 1 },
        },
      },
      "invalid-range",
    );
    await assert.rejects(
      apply(engine, [{ op: "replaceText", target: readOnly!, text: "x" }]),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "invalid-operation",
    );
    assert.equal((await engine.getElement(readOnly!, signal))!.text, "section");
    await engine.dispose();
  });
});

describe("DOCX setTextStyle and setParagraphStyle (docx-edit T55)", () => {
  it("merges run properties in schema order over a range and the paragraph mark", async () => {
    const engine = await open(
      buildDocx({
        theme: { major: "Aptos Display", minor: "Aptos" },
        body:
          `<w:p><w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>${RUN("abcdef", '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:cs="Arial"/><w:i/><w:u w:val="single"/><w:lang w:val="en-US"/></w:rPr>')}</w:p>` +
          sectPr(),
      }),
    );
    const [p] = await ids(engine);
    await apply(engine, [
      {
        op: "setTextStyle",
        target: p!,
        range: {
          start: { elementId: p!, offset: 2 },
          end: { elementId: p!, offset: 4 },
        },
        style: {
          bold: true,
          italic: false,
          fontSize: 14,
          color: { theme: "accent1" },
          highlight: "yellow",
          fontFamily: "Georgia",
        },
      },
    ]);
    const body = await bodyOf(engine);
    const own =
      '<w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:cs="Arial"/><w:i/><w:u w:val="single"/><w:lang w:val="en-US"/></w:rPr>';
    const changed =
      '<w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Arial"/><w:b/><w:bCs/><w:i w:val="0"/><w:iCs w:val="0"/><w:color w:val="4472C4" w:themeColor="accent1"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:highlight w:val="yellow"/><w:u w:val="single"/><w:lang w:val="en-US"/></w:rPr>';
    assert.ok(
      body.includes(
        `<w:r>${own}<w:t>ab</w:t></w:r><w:r>${changed}<w:t>cd</w:t></w:r><w:r>${own}<w:t>ef</w:t></w:r></w:p>`,
      ),
      body,
    );
    assert.ok(
      body.includes('<w:pPr><w:rPr><w:sz w:val="20"/></w:rPr></w:pPr>'),
    );
    // To the end: the paragraph mark takes the change too; "none" drops the highlight.
    await apply(engine, [
      {
        op: "setTextStyle",
        target: p!,
        range: {
          start: { elementId: p!, offset: 4 },
          end: { elementId: p!, offset: 6 },
        },
        style: { underline: false, highlight: "none", color: "auto" },
      },
    ]);
    const next = await bodyOf(engine);
    assert.ok(
      next.includes(
        '<w:pPr><w:rPr><w:color w:val="auto"/><w:sz w:val="20"/><w:u w:val="none"/></w:rPr></w:pPr>',
      ),
      next,
    );
    assert.ok(
      next.includes(
        '<w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:cs="Arial"/><w:i/><w:color w:val="auto"/><w:u w:val="none"/><w:lang w:val="en-US"/></w:rPr><w:t>ef</w:t></w:r>',
      ),
      next,
    );
    const element = (await engine.getElement(p!, signal))!;
    assert.equal(element.textStyle!.underline, true);
    assert.equal(element.textStyle!.fontFamily, "Aptos");
    await engine.dispose();
  });

  it("styles a whole paragraph, a field and a link, and ignores an empty change", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p>${RUN("a")}<w:hyperlink r:id="rId9">${RUN("b")}</w:hyperlink><w:fldSimple w:instr="PAGE">${RUN("3")}</w:fldSimple>` +
          `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${RUN("x")}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
          sectPr(),
      }),
    );
    const original = await engine.materialize("save", {}, signal);
    const [p] = await ids(engine);
    const none = await apply(engine, [
      { op: "setTextStyle", target: p!, style: {} },
    ]);
    assert.deepEqual(none.createdIds, []);
    assert.deepEqual(await engine.materialize("save", {}, signal), original);
    await apply(engine, [
      { op: "setTextStyle", target: p!, style: { bold: true } },
    ]);
    const body = await bodyOf(engine);
    assert.equal((body.match(/<w:b\/>/g) ?? []).length, 8);
    assert.match(
      body,
      /<w:hyperlink r:id="rId9"><w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:t>b<\/w:t><\/w:r><\/w:hyperlink>/,
    );
    assert.match(
      body,
      /<w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:fldChar w:fldCharType="begin"\/><\/w:r>/,
    );
    await expectIssue(
      engine,
      { op: "setTextStyle", target: p!, style: { color: "red" } },
      "invalid-value",
      "/style/color",
    );
    await expectIssue(
      engine,
      {
        op: "setTextStyle",
        target: p!,
        style: { color: { theme: "accent9" } },
      },
      "invalid-value",
      "/style/color",
    );
    await engine.dispose();
  });

  it("writes alignment and spacing into the paragraph properties in order", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p><w:pPr><w:keepNext/><w:spacing w:before="100" w:beforeLines="1"/><w:ind w:left="720"/><w:rPr><w:b/></w:rPr></w:pPr>${RUN("x")}</w:p>` +
          `<w:p>${RUN("y")}</w:p>` +
          `<w:p><w:pPr>${sectPr()}</w:pPr>${RUN("z")}</w:p>` +
          sectPr(),
      }),
    );
    const [p1, p2, readOnly] = await ids(engine);
    await apply(engine, [
      {
        op: "setParagraphStyle",
        target: p1!,
        style: { align: "justify", spacing: { after: 6, line: 1.5 } },
      },
      { op: "setParagraphStyle", target: p2!, style: { align: "center" } },
      { op: "setParagraphStyle", target: p2!, style: {} },
    ]);
    const body = await bodyOf(engine);
    assert.ok(
      body.includes(
        '<w:pPr><w:keepNext/><w:spacing w:before="100" w:beforeLines="1" w:after="120" w:line="360" w:lineRule="auto"/><w:ind w:left="720"/><w:jc w:val="both"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>',
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${p2!.slice(2)}"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">y</w:t></w:r></w:p>`,
      ),
      body,
    );
    const [first, second] = await engine.getElements({}, signal);
    assert.equal(first!.paragraphStyle!.align, "justify");
    assert.deepEqual(first!.paragraphStyle!.spacing, {
      before: 5,
      after: 6,
      line: 1.5,
      lineRule: "auto",
    });
    assert.equal(second!.paragraphStyle!.align, "center");
    await expectIssue(
      engine,
      { op: "setParagraphStyle", target: readOnly!, style: { align: "left" } },
      "invalid-target",
    );
    await engine.dispose();
  });
});

describe("DOCX text edits through the session (docx-edit T55)", () => {
  it("repaints from the paragraph's page to the end, undoes to identical bytes and reports ids", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + paragraph("three") + sectPr(),
    });
    const engine = await open(original);
    const [a, b, c] = (await ids(engine)).map((id) => id.slice(2));
    await engine.dispose();
    const pages = [
      [run(a!, "one", 72, 72)],
      [run(b!, "two", 72, 72)],
      [run(c!, "three", 72, 72)],
    ];
    const { session, end } = await docxSession(original, pages);
    try {
      const receipt = await session.replaceText({
        target: `p:${b}`,
        text: "TWO",
      });
      assert.deepEqual(receipt.changedPages, [1, 2]);
      assert.equal(receipt.revision, 1);
      assert.equal(session.state.dirty, true);
      const element = (await session.getElement(`p:${b}`)).item!;
      assert.equal(element.text, "TWO");
      const saved = await session.save();
      assert.deepEqual(changedEntries(original, saved.bytes), [
        "word/document.xml",
      ]);
      const undone = await session.undo();
      assert.deepEqual(undone.changedPages, [1, 2]);
      assert.deepEqual((await session.save()).bytes, original);
      const redone = await session.redo();
      assert.deepEqual(redone.changedPages, [1, 2]);
      const split = await session.replaceText({
        target: `p:${a}`,
        text: "1\n2",
      });
      assert.equal(split.createdIds.length, 1);
      assert.deepEqual(split.changedPages, [0, 1, 2]);
      const texts = (await session.getElements()).items.map((e) => e.text);
      assert.deepEqual(texts, ["1", "2", "TWO", "three"]);
      await session.reset();
      assert.deepEqual((await session.save()).bytes, original);
    } finally {
      await end();
    }
  });
});
