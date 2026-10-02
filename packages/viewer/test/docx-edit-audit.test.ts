import assert from "node:assert/strict";
import { describe, it } from "node:test";

import type { TextRun } from "../src/contracts.js";
import { DocxEditEngine } from "../src/edit/docx/engine.js";
import { loadDocxEditEngine } from "../src/edit/docx/provider.js";
import { DocxSession } from "../src/edit/docx/session.js";
import type { DocxOperation } from "../src/edit/docx/types.js";
import type { EditSessionAccess } from "../src/edit/engine.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { createOoxmlEditHandler } from "../src/edit/pptx/handler.js";
import {
  EditSessionController,
  type EditSessionHost,
} from "../src/edit/session.js";
import {
  defaultResourceLimits,
  ViewerError,
  type DocxEditSession,
  type ResourceLimits,
} from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";
import { docxSession, run } from "./fixtures/docx-session.js";
import { loopbackWorker } from "./fixtures/loopback-worker.js";

/*
 * Coverage audit of the DOCX editing module (docs/document-editing/todo/
 * 06-docx-edit.md, docs/api/editing.md "DOCX"): behaviour the T54–T58
 * suites do not exercise. Zero-width items (note references, anchored
 * drawings, bookmarks) around replaceText ranges, inline content controls,
 * fields next to hyperlinks and nested fields, tracked changes, text boxes
 * and the unauthored-id bookkeeping, batch rollback, restore from a shown
 * base, dry runs, style resolution edge cases, pictures sharing a
 * relationship, table widths across sections, and spacing merges.
 *
 * The last describe block holds the regressions for the defects the audit
 * found, fixed in the hardening pass that followed it.
 */

const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const PNG = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 1, 2, 3,
]);
const A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";
const WPS_NS =
  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

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

async function ids(engine: DocxEditEngine): Promise<string[]> {
  return (await engine.getElements({}, signal)).map((element) => element.id);
}

async function texts(engine: DocxEditEngine): Promise<(string | undefined)[]> {
  return (await engine.getElements({}, signal)).map((element) => element.text);
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

async function expectNoIssue(
  engine: DocxEditEngine,
  operation: DocxOperation,
): Promise<void> {
  const issues = await engine.validate([operation], signal);
  assert.deepEqual(issues, []);
}

const RUN = (text: string, rPr = ""): string =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

const range = (id: string, start: number, end: number) => ({
  start: { elementId: id, offset: start },
  end: { elementId: id, offset: end },
});

/** A complex field from begin to end with `instr` and the cached `result`. */
const FIELD = (instr: string, result: string): string =>
  `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>${result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;

/** An anchored drawing holding a text box with one paragraph of `text`. */
const TEXT_BOX = (text: string, docPrId = 7): string =>
  `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:docPr id="${docPrId}" name="Text Box ${docPrId}"/><a:graphic xmlns:a="${A_NS}"><a:graphicData uri="${WPS_NS}"><wps:wsp xmlns:wps="${WPS_NS}"><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;

const TABLE = (cells: readonly string[]): string =>
  `<w:tbl><w:tblPr/><w:tblGrid>${cells.map(() => '<w:gridCol w:w="2000"/>').join("")}</w:tblGrid><w:tr>${cells
    .map((cell) => `<w:tc>${paragraph(cell)}</w:tc>`)
    .join("")}</w:tr></w:tbl>`;

const FOOTNOTES =
  '<w:footnote w:id="1"><w:p><w:r><w:t>a note</w:t></w:r></w:p></w:footnote>';

function count(haystack: string, needle: string): number {
  return haystack.split(needle).length - 1;
}

/** Every w14:paraId value of a main part, in document order. */
function paraIds(xml: string): string[] {
  return [...xml.matchAll(/w14:paraId="([0-9A-F]{8})"/g)].map((m) => m[1]!);
}

/**
 * The session fixture of test/fixtures/docx-session.ts with the resource
 * limits chosen by the test, for checkpoint behaviour.
 */
async function sessionWithLimits(
  original: Uint8Array,
  pages: readonly (readonly TextRun[])[],
  own: ResourceLimits,
): Promise<{ session: DocxSession & DocxEditSession; end(): Promise<void> }> {
  const pair = loopbackWorker(createOoxmlEditHandler());
  const engine = await loadDocxEditEngine(
    original,
    { format: "docx", limits: own, signal },
    { createWorker: () => pair.worker },
  );
  const host: EditSessionHost = {
    format: "docx",
    limits: own,
    prepareDocument: async () => ({ pageCount: pages.length }),
    commitDocument: (prepared) => prepared.pageCount,
    discardDocument: () => {},
    emit: () => {},
    pageOf: (paragraphId) => {
      const index = pages.findIndex((runs) =>
        runs.some((item) => item.paragraphId === paragraphId),
      );
      return index < 0 ? undefined : index;
    },
  };
  const core = new EditSessionController(engine, host, original, pages.length);
  const cached = new Set<number>();
  const access: EditSessionAccess = {
    getTextRuns: async (pageIndex) => {
      cached.add(pageIndex);
      return pages[pageIndex] ?? [];
    },
    cachedPages: () => [...cached],
  };
  return { session: new DocxSession(core, access), end: () => core.end() };
}

describe("audit: zero-width items around replaceText and setTextStyle", () => {
  it("keeps a note reference in its own run when the range ends at it or starts after it", async () => {
    const body =
      `<w:p>${RUN("abc")}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>${RUN("def")}</w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, footnotes: FOOTNOTES }));
    const [p] = await ids(engine);
    assert.equal((await texts(engine))[0], "abcdef");
    // The range ends at the reference: the reference is after it.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 0, 3) },
    ]);
    let xml = await bodyOf(engine);
    assert.equal(count(xml, "<w:footnoteReference"), 1);
    assert.match(
      xml,
      /<w:t>X<\/w:t><\/w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"\/><\/w:rPr><w:footnoteReference w:id="1"\/><\/w:r><w:r><w:t xml:space="preserve">def<\/w:t>/,
    );
    assert.equal((await texts(engine))[0], "Xdef");
    // A caret right at the reference inserts after it, once.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "+", range: range(p!, 1, 1) },
    ]);
    xml = await bodyOf(engine);
    assert.equal(count(xml, "<w:footnoteReference"), 1);
    assert.equal((await texts(engine))[0], "X+def");
    await engine.dispose();
  });

  it("keeps a note reference that shares a run with text when the edit does not reach it", async () => {
    const body =
      `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>abc</w:t><w:footnoteReference w:id="1"/><w:t>def</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, footnotes: FOOTNOTES }));
    const [p] = await ids(engine);
    // "c" replaced: the reference starts the kept tail and stays, once.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 2, 3) },
    ]);
    let xml = await bodyOf(engine);
    assert.equal(count(xml, "<w:footnoteReference"), 1, xml);
    assert.match(
      xml,
      /<w:t>X<\/w:t><\/w:r><w:r><w:rPr><w:i\/><\/w:rPr><w:footnoteReference w:id="1"\/><w:t>def<\/w:t>/,
      xml,
    );
    assert.equal((await texts(engine))[0], "abXdef");
    // "d" replaced: the reference ends the kept head and stays, once.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "Y", range: range(p!, 3, 4) },
    ]);
    xml = await bodyOf(engine);
    assert.equal(count(xml, "<w:footnoteReference"), 1, xml);
    assert.equal((await texts(engine))[0], "abXYef");
    await engine.dispose();
  });

  it("cuts a run of several content children in the middle and drops only what the range covers", async () => {
    const body =
      `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t><w:tab/><w:t>cd</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [p] = await ids(engine);
    assert.equal((await texts(engine))[0], "ab\tcd");
    await apply(engine, [
      { op: "replaceText", target: p!, text: "-", range: range(p!, 1, 4) },
    ]);
    assert.equal((await texts(engine))[0], "a-d");
    let xml = await bodyOf(engine);
    assert.ok(!xml.includes("<w:tab/>"), xml);
    assert.match(
      xml,
      /<w:r><w:rPr><w:b\/><\/w:rPr><w:t>a<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>-<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>d<\/w:t><\/w:r>/,
      xml,
    );
    await engine.dispose();
    // Covering only the tab keeps both texts in one rebuilt run each.
    const again = await open(buildDocx({ body }));
    const [q] = await ids(again);
    await apply(again, [
      { op: "replaceText", target: q!, text: " ", range: range(q!, 2, 3) },
    ]);
    assert.equal((await texts(again))[0], "ab cd");
    xml = await bodyOf(again);
    assert.ok(!xml.includes("<w:tab/>"), xml);
    assert.match(xml, /<w:t xml:space="preserve"> <\/w:t>/);
    await again.dispose();
  });

  it("inserts at offset 0 of a paragraph that starts with a bookmark and keeps the pair", async () => {
    const body =
      `<w:p><w:bookmarkStart w:id="3" w:name="start"/>${RUN("abc")}<w:bookmarkEnd w:id="3"/></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [p] = await ids(engine);
    await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 0, 0) },
    ]);
    const xml = await bodyOf(engine);
    assert.equal((await texts(engine))[0], "Xabc");
    assert.equal(count(xml, "<w:bookmarkStart"), 1);
    assert.equal(count(xml, "<w:bookmarkEnd"), 1);
    assert.ok(
      xml.indexOf('<w:bookmarkStart w:id="3" w:name="start"/>') <
        xml.indexOf("<w:t>X</w:t>"),
      xml,
    );
    assert.ok(xml.indexOf("<w:t>X</w:t>") < xml.indexOf("<w:bookmarkEnd"));
    await engine.dispose();
  });

  it("removes an inline content control covered whole, refuses an edge cut and edits inside it", async () => {
    const body =
      `<w:p>${RUN("pre ")}<w:sdt><w:sdtPr><w:alias w:val="x"/></w:sdtPr><w:sdtContent>${RUN("control", "<w:rPr><w:b/></w:rPr>")}</w:sdtContent></w:sdt>${RUN(" post")}</w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [p] = await ids(engine);
    assert.equal((await texts(engine))[0], "pre control post");
    // Cutting an edge: refused.
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "x", range: range(p!, 2, 6) },
      "invalid-range",
      "/range",
    );
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "x", range: range(p!, 6, 13) },
      "invalid-range",
      "/range",
    );
    // Strictly inside: the control stays and its run changes.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "ON", range: range(p!, 5, 7) },
    ]);
    let xml = await bodyOf(engine);
    assert.equal((await texts(engine))[0], "pre cONtrol post");
    assert.match(
      xml,
      /<w:sdt><w:sdtPr><w:alias w:val="x"\/><\/w:sdtPr><w:sdtContent><w:r><w:rPr><w:b\/><\/w:rPr><w:t>c<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>ON<\/w:t><\/w:r><w:r><w:rPr><w:b\/><\/w:rPr><w:t>trol<\/w:t><\/w:r><\/w:sdtContent><\/w:sdt>/,
      xml,
    );
    // A paragraph break inside the control is refused.
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "a\nb", range: range(p!, 5, 6) },
      "invalid-range",
      "/text",
    );
    // Covered whole: gone.
    await apply(engine, [
      { op: "replaceText", target: p!, text: "", range: range(p!, 4, 11) },
    ]);
    xml = await bodyOf(engine);
    assert.ok(!xml.includes("<w:sdt>"), xml);
    assert.equal((await texts(engine))[0], "pre  post");
    await engine.dispose();
  });

  it("keeps an anchored drawing in its own run when the range ends at it", async () => {
    const body =
      `<w:p>${RUN("abc")}${inlinePicture(100, 100, "rId9", true)}${RUN("def")}</w:p>` +
      sectPr();
    const engine = await open(
      buildDocx({
        body,
        media: [{ name: "word/media/image1.png", data: PNG }],
      }),
    );
    const [p, other] = await ids(engine);
    assert.equal(other, `other:${p!.slice(2)}.0`);
    const change = await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 0, 3) },
    ]);
    assert.deepEqual(change.removedIds, []);
    assert.ok((await bodyOf(engine)).includes("<wp:anchor"));
    assert.deepEqual(await ids(engine), [p, other]);
    await engine.dispose();
  });

  it("reports an anchored drawing it drops under a covering range", async () => {
    const body =
      `<w:p>${RUN("abc")}${inlinePicture(100, 100, "rId9", true)}${RUN("def")}</w:p>` +
      sectPr();
    const engine = await open(
      buildDocx({
        body,
        media: [{ name: "word/media/image1.png", data: PNG }],
      }),
    );
    const [p, other] = await ids(engine);
    const change = await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 1, 5) },
    ]);
    assert.deepEqual(change.removedIds, [other]);
    assert.ok(!(await bodyOf(engine)).includes("<wp:anchor"));
    assert.deepEqual(await ids(engine), [p]);
    await engine.dispose();
  });
});

describe("audit: fields next to hyperlinks and nested fields", () => {
  it("edits inside a hyperlink without touching a complex field later in the paragraph", async () => {
    const body =
      `<w:p>${RUN("see ")}<w:hyperlink r:id="rId9">${RUN("link")}</w:hyperlink>${RUN(" page ")}${FIELD("PAGE", "77")}${RUN(" end")}</w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [p] = await ids(engine);
    assert.equal((await texts(engine))[0], "see link page 77 end");
    const before = await bodyOf(engine);
    const field = before.slice(
      before.indexOf('<w:r><w:fldChar w:fldCharType="begin"/>'),
      before.indexOf('<w:fldChar w:fldCharType="end"/></w:r>') +
        '<w:fldChar w:fldCharType="end"/></w:r>'.length,
    );
    await apply(engine, [
      { op: "replaceText", target: p!, text: "IN", range: range(p!, 5, 7) },
    ]);
    const after = await bodyOf(engine);
    assert.equal((await texts(engine))[0], "see lINk page 77 end");
    assert.ok(after.includes(field), after);
    assert.equal(count(after, "<w:fldChar "), 3);
    // A range that starts in the result of the field is refused; one that
    // starts before it and ends right at its end removes it.
    await expectIssue(
      engine,
      { op: "replaceText", target: p!, text: "x", range: range(p!, 15, 17) },
      "invalid-range",
    );
    await apply(engine, [
      { op: "replaceText", target: p!, text: "N", range: range(p!, 13, 16) },
    ]);
    const gone = await bodyOf(engine);
    assert.equal(count(gone, "<w:fldChar "), 0);
    assert.ok(gone.includes('<w:hyperlink r:id="rId9">'));
    assert.equal((await texts(engine))[0], "see lINk pageN end");
    await engine.dispose();
  });

  it("reads a field nested in another field's instruction as one item with the outer result", async () => {
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> IF </w:instrText></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r><w:instrText xml:space="preserve"> = 7 "yes" "no" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>yes</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>${RUN("!")}</w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [element] = await engine.getElements({}, signal);
    assert.equal(element!.text, "yes!");
    const model = await engine.model(signal);
    const items = model.paragraphs[0]!.text.items;
    assert.deepEqual(
      items.map((item) => item.kind),
      ["field", "text"],
    );
    assert.equal(items[0]!.fieldRuns?.length, 11);
    // The text after the field is reachable and the field keeps its runs.
    const [p] = await ids(engine);
    await apply(engine, [
      { op: "replaceText", target: p!, text: "?", range: range(p!, 3, 4) },
    ]);
    assert.equal((await texts(engine))[0], "yes?");
    assert.equal(count(await bodyOf(engine), "<w:fldChar "), 6);
    await engine.dispose();
  });
});

describe("audit: tracked changes", () => {
  it("marks deletion-only and marker-only paragraphs read-only and limits them to insertions", async () => {
    const body =
      `<w:p><w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>` +
      `<w:p><w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins></w:p>` +
      `<w:p><w:r><w:rPr><w:ins w:id="3" w:author="a" w:date="2026-01-01T00:00:00Z"/></w:rPr><w:t>marked</w:t></w:r></w:p>` +
      `<w:p><w:moveFrom w:id="4" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>moved</w:t></w:r></w:moveFrom></w:p>` +
      paragraph("plain") +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const elements = await engine.getElements({}, signal);
    assert.deepEqual(
      elements.map((element) => [element.text, element.readOnlyReason]),
      [
        ["", "tracked-changes"],
        ["added", "tracked-changes"],
        ["marked", "tracked-changes"],
        ["", "tracked-changes"],
        ["plain", undefined],
      ],
    );
    for (const element of elements.slice(0, 4))
      assert.deepEqual(element.operations, [
        "insertParagraph",
        "insertTable",
        "insertImage",
      ]);
    const [del, ins, marked] = elements.map((element) => element.id);
    await expectIssue(
      engine,
      { op: "replaceText", target: del!, text: "x" },
      "invalid-target",
    );
    await expectIssue(
      engine,
      { op: "setTextStyle", target: ins!, style: { bold: true } },
      "invalid-target",
    );
    await expectIssue(
      engine,
      { op: "setParagraphStyle", target: marked!, style: { align: "center" } },
      "invalid-target",
    );
    await expectIssue(
      engine,
      { op: "deleteElement", target: ins! },
      "invalid-target",
    );
    // Siblings may still be placed next to a tracked paragraph.
    const change = await apply(engine, [
      { op: "insertParagraph", after: del!, text: "next" },
    ]);
    assert.equal(change.createdIds.length, 1);
    assert.deepEqual((await texts(engine)).slice(0, 2), ["", "next"]);
    // The tracked paragraph keeps its bytes.
    assert.ok(
      (await bodyOf(engine)).includes(
        '<w:p><w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>',
      ),
    );
    await engine.dispose();
  });
});

describe("audit: text boxes and the unauthored-id bookkeeping", () => {
  const HOST = (text = "box"): string =>
    `<w:p>${RUN("host ")}${TEXT_BOX(text)}${RUN(" end")}</w:p>`;

  it("lists a text box host as a paragraph with one other object and keeps ids through a caret edit", async () => {
    const engine = await open(
      buildDocx({ body: HOST() + paragraph("after") + sectPr() }),
    );
    const listed = await engine.getElements({}, signal);
    assert.deepEqual(
      listed.map((element) => [element.kind, element.text]),
      [
        ["paragraph", "host  end"],
        ["other", undefined],
        ["paragraph", "after"],
      ],
    );
    const [host, other, after] = listed.map((element) => element.id);
    await apply(engine, [
      {
        op: "replaceText",
        target: host!,
        text: "X",
        range: range(host!, 0, 0),
      },
    ]);
    assert.deepEqual(await ids(engine), [host, other, after]);
    assert.equal((await texts(engine))[0], "Xhost  end");
    assert.ok((await bodyOf(engine)).includes("<w:txbxContent>"));
    // The shown copy stamps the host, the box paragraph and the rest with distinct ids.
    const shown = paraIds(
      await documentXml(await engine.materialize("show", {}, signal)),
    );
    assert.equal(shown.length, 3);
    assert.equal(new Set(shown).size, 3);
    assert.equal(shown[0], host!.slice(2));
    assert.equal(shown[2], after!.slice(2));
    await engine.dispose();
  });

  it("moves and deletes a text box host with its nested paragraph accounted for", async () => {
    const engine = await open(
      buildDocx({ body: HOST() + paragraph("after") + sectPr() }),
    );
    const [host, other, after] = await ids(engine);
    await apply(engine, [{ op: "moveElement", target: host!, after: after! }]);
    assert.deepEqual(await ids(engine), [after, host, other]);
    const moved = await documentXml(
      await engine.materialize("show", {}, signal),
    );
    const shown = paraIds(moved);
    assert.equal(shown.length, 3);
    assert.equal(new Set(shown).size, 3);
    const change = await apply(engine, [
      { op: "deleteElement", target: host! },
    ]);
    assert.deepEqual(change.removedIds, [host, other]);
    assert.deepEqual(await ids(engine), [after]);
    assert.ok(!(await bodyOf(engine)).includes("<w:txbxContent>"));
    assert.deepEqual(
      paraIds(await documentXml(await engine.materialize("show", {}, signal))),
      [after!.slice(2)],
    );
    await engine.dispose();
  });

  it("replaces a cell whose other paragraph holds a text box and stays consistent", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${paragraph("first")}${HOST("in cell")}</w:tc></w:tr></w:tbl>` +
          paragraph("after") +
          sectPr(),
      }),
    );
    const [table, first, host, other, after] = await ids(engine);
    const change = await apply(engine, [
      { op: "setTableCell", target: table!, row: 0, column: 0, text: "new" },
    ]);
    assert.deepEqual(change.removedIds, [host, other]);
    assert.deepEqual(await ids(engine), [table, first, after]);
    const shown = paraIds(
      await documentXml(await engine.materialize("show", {}, signal)),
    );
    assert.deepEqual(shown, [first!.slice(2), after!.slice(2)]);
    assert.ok(!(await bodyOf(engine)).includes("<w:txbxContent>"));
    await engine.dispose();
  });
});

describe("audit: batches and rollback", () => {
  it("undoes everything a failed batch did, ids included, in the engine", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + sectPr(),
    });
    const engine = await open(original);
    const before = await ids(engine);
    // The first operation rebuilds (and stamps) a paragraph; the second
    // refers to an operation that created nothing and fails at apply time.
    await assert.rejects(
      apply(engine, [
        { op: "replaceText", target: before[0]!, text: "ONE" },
        { op: "insertParagraph", after: "$0", text: "x" },
      ]),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "invalid-operation",
    );
    assert.deepEqual(await ids(engine), before);
    assert.deepEqual(await texts(engine), ["one", "two"]);
    assert.deepEqual(await engine.materialize("save", {}, signal), original);
    assert.deepEqual(await engine.materialize("show", {}, signal), original);
    // The engine keeps working and stamps the shown copy from a sound list.
    await apply(engine, [{ op: "replaceText", target: before[1]!, text: "2" }]);
    const shown = paraIds(
      await documentXml(await engine.materialize("show", {}, signal)),
    );
    assert.deepEqual(
      shown,
      before.map((id) => id.slice(2)),
    );
    await engine.dispose();
  });

  it("leaves the session untouched when a batch fails after its first operation", async () => {
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
      await assert.rejects(
        session.apply([
          { op: "replaceText", target: `p:${a}`, text: "ONE" },
          { op: "insertParagraph", after: "$0", text: "x" },
        ]),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
      assert.equal(session.state.revision, 0);
      assert.equal(session.state.dirty, false);
      assert.deepEqual((await session.save()).bytes, original);
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.text),
        ["one", "two"],
      );
      const receipt = await session.replaceText({
        target: `p:${a}`,
        text: "ONE",
      });
      assert.equal(receipt.revision, 1);
    } finally {
      await end();
    }
  });

  it("resolves same-batch references through every operation kind", async () => {
    const engine = await open(buildDocx({ body: paragraph("one") + sectPr() }));
    const [one] = await ids(engine);
    const change = await apply(engine, [
      { op: "insertParagraph", after: one!, text: "p" },
      { op: "insertTable", after: "$0", rows: [["c"]] },
      { op: "setTableCell", target: "$1", row: 0, column: 0, text: "C" },
      {
        op: "insertImage",
        after: "$1",
        data: PNG,
        mimeType: "image/png",
        size: { width: 1, height: 1 },
      },
      { op: "moveElement", target: "$0", after: "$1" },
      { op: "setTextStyle", target: "$0", style: { bold: true } },
      { op: "setParagraphStyle", target: "$0", style: { align: "right" } },
      // A range cannot name "$0": only target, before and after resolve.
      { op: "replaceText", target: "$0", text: "P" },
      { op: "deleteElement", target: "$3" },
    ]);
    assert.equal(change.createdIds.length, 1 + 3 + 2);
    const listed = (await engine.getElements({}, signal)).map((element) => [
      element.kind,
      element.text,
    ]);
    assert.deepEqual(listed, [
      ["paragraph", "one"],
      ["table", "C"],
      ["paragraph", "C"],
      ["paragraph", "P"],
      ["paragraph", ""],
    ]);
    const [, , , moved] = await engine.getElements({}, signal);
    assert.equal(moved!.textStyle!.bold, true);
    assert.equal(moved!.paragraphStyle!.align, "right");
    await engine.dispose();
  });
});

describe("audit: restore from a shown base", () => {
  it("issues the same ids after a shown copy as after the live edits", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + paragraph("three") + sectPr(),
    });
    const live = await open(original);
    const [one, , three] = await ids(live);
    const a: DocxOperation[] = [
      { op: "replaceText", target: one!, text: "1\n1b" },
      { op: "deleteElement", target: three! },
    ];
    await live.apply({ stateId: 1, operations: a }, signal);
    const shownAfterA = await live.materialize("show", {}, signal);
    const afterA = await ids(live);
    const b: DocxOperation[] = [
      { op: "insertParagraph", after: afterA[1]!, text: "z\nw" },
      { op: "insertTable", before: "$0", rows: [["t"]] },
      { op: "replaceText", target: afterA[2]!, text: "2\n2b" },
    ];
    const liveChange = await live.apply({ stateId: 2, operations: b }, signal);
    const liveIds = await ids(live);
    const liveTexts = await texts(live);
    await live.dispose();

    const restored = await open(original);
    await restored.restore(
      { base: shownAfterA, batches: [{ stateId: 2, operations: b }] },
      signal,
    );
    assert.deepEqual(await ids(restored), liveIds);
    assert.deepEqual(await texts(restored), liveTexts);
    // A dry run of the same batch on a third engine names the same ids.
    const dry = await open(original);
    await dry.restore({ base: shownAfterA, batches: [] }, signal);
    const dryChange = await dry.apply({ stateId: 2, operations: b }, signal);
    assert.deepEqual(dryChange.createdIds, liveChange.createdIds);
    await restored.dispose();
    await dry.dispose();
  });
});

describe("audit: session dry runs, hit tests and search", () => {
  it("applies a dry run without changing the document and names the ids a real apply gives", async () => {
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
      const dry = await session.apply(
        [
          { op: "replaceText", target: `p:${a}`, text: "1\n1b" },
          { op: "insertTable", after: "$0", rows: [["x"]] },
        ],
        { dryRun: true },
      );
      assert.equal(dry.dryRun, true);
      assert.equal(dry.createdIds.length, 1 + 2);
      assert.equal(session.state.revision, 0);
      assert.equal(session.state.dirty, false);
      assert.deepEqual((await session.save()).bytes, original);
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.text),
        ["one", "two"],
      );
      const real = await session.apply([
        { op: "replaceText", target: `p:${a}`, text: "1\n1b" },
        { op: "insertTable", after: "$0", rows: [["x"]] },
      ]);
      assert.deepEqual(real.createdIds, dry.createdIds);
      assert.equal(real.dryRun, false);
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.text),
        ["1", "1b", "x", "x", "two"],
      );
    } finally {
      await end();
    }
  });

  it("hit-tests a page without runs to nothing and reads only that page", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + sectPr(),
    });
    const probe = await open(original);
    const [a] = (await ids(probe)).map((id) => id.slice(2));
    await probe.dispose();
    const { session, reads, end } = await docxSession(original, [
      [run(a!, "one", 72, 72)],
      [],
    ]);
    try {
      assert.deepEqual(
        (await session.elementsAt(1, { x: 80, y: 80 })).items,
        [],
      );
      assert.deepEqual(reads, [1]);
      assert.deepEqual((await session.getElements({ pageIndex: 1 })).items, []);
    } finally {
      await end();
    }
  });

  it("places matches from cached pages and cuts rectangles to the matched characters", async () => {
    const original = buildDocx({
      body: paragraph("alpha one") + paragraph("alpha two") + sectPr(),
    });
    const probe = await open(original);
    const [a, b] = (await ids(probe)).map((id) => id.slice(2));
    await probe.dispose();
    const { session, reads, end } = await docxSession(
      original,
      [[run(a!, "alpha one", 72, 72)], [run(b!, "alpha two", 72, 72)]],
      { cached: [1] },
    );
    try {
      const hits = (await session.findText("two")).items;
      assert.equal(hits.length, 1);
      // The cached page is read first and holds the match; one more page
      // is read to see whether the paragraph continues, then the scan stops.
      assert.deepEqual(reads, [1, 0]);
      assert.deepEqual(hits[0]!.rects, [
        { x: 72 + 6 * 6, y: 72, width: 18, height: 12 },
      ]);
      assert.equal(hits[0]!.pageIndex, 1);
    } finally {
      await end();
    }
  });

  it("undoes a split and a move with their ids and pages, then redoes them", async () => {
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
      const split = await session.replaceText({
        target: `p:${b}`,
        text: "2\n2b",
      });
      const moved = await session.moveElement({
        target: `p:${b}`,
        before: `p:${a}`,
      });
      assert.deepEqual(moved.changedPages, [0, 1]);
      const undoneMove = await session.undo();
      assert.deepEqual(undoneMove.changedPages, [0, 1]);
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.id),
        [`p:${a}`, `p:${b}`, split.createdIds[0]],
      );
      const undoneSplit = await session.undo();
      assert.deepEqual(undoneSplit.removedIds, split.createdIds);
      assert.deepEqual(undoneSplit.changedPages, [1]);
      assert.deepEqual((await session.save()).bytes, original);
      const redone = await session.redo();
      assert.deepEqual(redone.createdIds, split.createdIds);
      await session.redo();
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.text),
        ["2", "one", "2b"],
      );
    } finally {
      await end();
    }
  });
});

describe("audit: style resolution edge cases", () => {
  it("reads toggle values, stops at a style cycle, defaults the size and ignores highlight none", async () => {
    const styles =
      "<w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>" +
      '<w:style w:type="character" w:styleId="Loop1"><w:basedOn w:val="Loop2"/><w:rPr><w:b/></w:rPr></w:style>' +
      '<w:style w:type="character" w:styleId="Loop2"><w:basedOn w:val="Loop1"/><w:rPr><w:i/></w:rPr></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Self"><w:basedOn w:val="Self"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style>';
    const body =
      `<w:p><w:r><w:rPr><w:b w:val="false"/><w:i w:val="off"/><w:u w:val="none"/></w:rPr><w:t>off</w:t></w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:b w:val="on"/><w:i w:val="true"/><w:highlight w:val="none"/></w:rPr><w:t>on</w:t></w:r></w:p>` +
      `<w:p><w:r><w:rPr><w:b w:val="1"/><w:i w:val="0"/></w:rPr><w:t>digits</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Self"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Loop1"/></w:rPr><w:t>loop</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, styles }));
    const elements = await engine.getElements({}, signal);
    assert.deepEqual(
      elements.map((element) => [
        element.textStyle!.bold,
        element.textStyle!.italic,
        element.textStyle!.underline,
        element.textStyle!.fontSize,
        element.textStyle!.highlight,
      ]),
      [
        [false, false, false, 10, undefined],
        [true, true, false, 10, undefined],
        [true, false, false, 10, undefined],
        [true, true, false, 10, undefined],
      ],
    );
    assert.equal(elements[3]!.paragraphStyle!.align, "right");
    assert.equal(elements[3]!.paragraphStyle!.styleId, "Self");
    assert.equal(elements[0]!.textStyle!.fontFamily, "Calibri");
    await engine.dispose();
  });

  it("writes a theme colour as automatic without a theme part and leaves a run without highlight alone", async () => {
    const engine = await open(
      buildDocx({
        body: `<w:p>${RUN("a", "<w:rPr><w:i/></w:rPr>")}</w:p>` + sectPr(),
      }),
    );
    const [p] = await ids(engine);
    await apply(engine, [
      {
        op: "setTextStyle",
        target: p!,
        style: { color: { theme: "accent1" }, highlight: "none" },
      },
    ]);
    const xml = await bodyOf(engine);
    assert.match(
      xml,
      /<w:rPr><w:i\/><w:color w:val="auto" w:themeColor="accent1"\/><\/w:rPr>/,
      xml,
    );
    assert.ok(!xml.includes("w:highlight"), xml);
    const [element] = await engine.getElements({}, signal);
    assert.deepEqual(element!.textStyle!.color, { theme: "accent1" });
    assert.equal(element!.textStyle!.highlight, undefined);
    // Setting a highlight then removing it restores a run without one.
    await apply(engine, [
      { op: "setTextStyle", target: p!, style: { highlight: "cyan" } },
    ]);
    assert.equal(
      (await engine.getElement(p!, signal))!.textStyle!.highlight,
      "cyan",
    );
    await apply(engine, [
      { op: "setTextStyle", target: p!, style: { highlight: "none" } },
    ]);
    assert.ok(!(await bodyOf(engine)).includes("w:highlight"));
    await engine.dispose();
  });

  it("reports a theme colour of a styled run from a theme without that slot as the theme name", async () => {
    const engine = await open(
      buildDocx({
        theme: { major: "M", minor: "m" },
        body:
          `<w:p>${RUN("a", '<w:rPr><w:color w:val="1F3864" w:themeColor="accent5"/></w:rPr>')}</w:p>` +
          sectPr(),
      }),
    );
    const [element] = await engine.getElements({}, signal);
    assert.deepEqual(element!.textStyle!.color, { theme: "accent5" });
    const [p] = await ids(engine);
    // accent5 is absent from the fixture theme: the value falls back to auto.
    await apply(engine, [
      {
        op: "setTextStyle",
        target: p!,
        style: { color: { theme: "accent5" } },
      },
    ]);
    assert.match(
      await bodyOf(engine),
      /<w:color w:val="auto" w:themeColor="accent5"\/>/,
    );
    await engine.dispose();
  });
});

describe("audit: pictures through assets and shared relationships", () => {
  it("inserts a picture from a session asset, stores identical bytes once and replays through undo and redo", async () => {
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
      const asset = await session.addAsset(PNG, { mimeType: "image/png" });
      assert.match(asset, /^asset:[0-9a-f]{64}$/);
      const first = await session.insertImage({
        after: `p:${a}`,
        data: asset,
        mimeType: "image/png",
        size: { width: 10, height: 10 },
      });
      const second = await session.insertImage({
        after: `p:${b}`,
        data: PNG,
        mimeType: "image/png",
        size: { width: 20, height: 20 },
      });
      assert.equal(first.createdIds.length, 2);
      assert.equal(second.createdIds.length, 2);
      const pkg = await OoxmlPackage.open((await session.save()).bytes, {
        limits,
      });
      assert.deepEqual(
        pkg.currentPartNames.filter((name) => name.startsWith("/word/media/")),
        ["/word/media/image1.png"],
      );
      const xml = new TextDecoder().decode(
        await pkg.part("/word/document.xml"),
      );
      assert.equal(count(xml, "<wp:inline"), 2);
      assert.match(xml, /<wp:docPr id="1" /);
      assert.match(xml, /<wp:docPr id="2" /);
      await session.undo();
      await session.undo();
      assert.deepEqual((await session.save()).bytes, original);
      await session.redo();
      await session.redo();
      const items = (await session.getElements()).items;
      assert.deepEqual(
        items.map((item) => item.kind),
        ["paragraph", "paragraph", "image", "paragraph", "paragraph", "image"],
      );
      await assert.rejects(
        session.insertImage({
          after: `p:${a}`,
          data: `asset:${"0".repeat(64)}`,
          mimeType: "image/png",
          size: { width: 1, height: 1 },
        }),
        (error: unknown) =>
          error instanceof ViewerError && error.code === "invalid-operation",
      );
    } finally {
      await end();
    }
  });

  it("keeps a relationship two pictures share until the last of them goes", async () => {
    const engine = await open(
      buildDocx({
        media: [{ name: "word/media/image1.png", data: PNG }],
        body:
          `<w:p>${inlinePicture(100, 100)}</w:p>` +
          `<w:p>${inlinePicture(200, 200)}</w:p>` +
          paragraph("text") +
          sectPr(),
      }),
    );
    const [, img1, , img2] = await ids(engine);
    const rels = async () =>
      (
        await OoxmlPackage.open(await engine.materialize("save", {}, signal), {
          limits,
        })
      ).relationships("/word/document.xml");
    await apply(engine, [{ op: "deleteElement", target: img1! }]);
    assert.ok((await rels()).byId("rId9"), "the shared relationship stays");
    assert.ok(
      (await bodyOf(engine)).includes('<wp:extent cx="200" cy="200"/>'),
    );
    await apply(engine, [{ op: "deleteElement", target: img2! }]);
    assert.equal((await rels()).byId("rId9"), undefined);
    assert.ok(!(await bodyOf(engine)).includes("w:drawing"));
    // The media part stays, as in PPTX.
    const pkg = await OoxmlPackage.open(
      await engine.materialize("save", {}, signal),
      { limits },
    );
    assert.ok(pkg.has("/word/media/image1.png"));
    await engine.dispose();
  });
});

describe("audit: tables across sections and uneven weights", () => {
  const SECTIONED = buildDocx({
    body:
      paragraph("one") +
      `<w:p><w:pPr>${sectPr({ width: 8000, margin: 500 })}</w:pPr><w:r><w:t>ends section one</w:t></w:r></w:p>` +
      paragraph("two") +
      sectPr(),
  });

  it("accepts a section-break paragraph as a reference and sizes a table before it from its section", async () => {
    const engine = await open(SECTIONED);
    const [, section] = await ids(engine);
    assert.equal(
      (await engine.getElement(section!, signal))!.readOnlyReason,
      "section-break",
    );
    await expectNoIssue(engine, {
      op: "insertTable",
      before: section!,
      rows: [["a", "b", "c"]],
    });
    const change = await apply(engine, [
      { op: "insertTable", before: section!, rows: [["a", "b", "c"]] },
    ]);
    assert.equal(change.createdIds.length, 4);
    const xml = await bodyOf(engine);
    // 7000 twips split three ways: the remainder goes to the last column.
    assert.ok(
      xml.includes(
        '<w:tblGrid><w:gridCol w:w="2333"/><w:gridCol w:w="2333"/><w:gridCol w:w="2334"/></w:tblGrid>',
      ),
      xml,
    );
    assert.equal(count(xml, '<w:tcW w:w="2333" w:type="dxa"/>'), 2);
    assert.equal(count(xml, '<w:tcW w:w="2334" w:type="dxa"/>'), 1);
    // The section paragraph still ends its section, after the table.
    assert.ok(xml.indexOf("</w:tbl>") < xml.indexOf("ends section one"));
    assert.equal(count(xml, "<w:sectPr>"), 2);
    await engine.dispose();
  });

  it("splits weights that do not divide evenly so the grid sums to the content width", async () => {
    const engine = await open(buildDocx({ body: paragraph("x") + sectPr() }));
    const [x] = await ids(engine);
    await apply(engine, [
      {
        op: "insertTable",
        after: x!,
        rows: [["a", "b", "c", "d", "e", "f", "g"]],
        columnWidths: [1, 1, 1, 1, 1, 1, 1],
      },
      {
        op: "insertTable",
        after: x!,
        rows: [["a", "b", "c"]],
        columnWidths: [0.5, 0.25, 0.25],
      },
    ]);
    const xml = await bodyOf(engine);
    const grids = [...xml.matchAll(/<w:tblGrid>(.*?)<\/w:tblGrid>/g)].map((m) =>
      [...m[1]!.matchAll(/w:w="(\d+)"/g)].map((w) => Number(w[1])),
    );
    assert.equal(grids.length, 2);
    for (const grid of grids)
      assert.equal(
        grid.reduce((a, b) => a + b, 0),
        9360,
        JSON.stringify(grid),
      );
    assert.deepEqual(grids[0], [4680, 2340, 2340]);
    assert.equal(grids[1]!.length, 7);
    assert.ok(grids[1]!.slice(0, 6).every((w) => w === 1337));
    assert.equal(grids[1]![6], 9360 - 6 * 1337);
    await engine.dispose();
  });
});

describe("audit: paragraph spacing merges", () => {
  it("rewrites an exact line rule as auto when the line changes and keeps it otherwise", async () => {
    const body =
      `<w:p><w:pPr><w:spacing w:before="100" w:line="480" w:lineRule="exact"/><w:jc w:val="right"/></w:pPr>${RUN("a")}</w:p>` +
      `<w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>${RUN("b")}</w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [p1, p2] = await ids(engine);
    await apply(engine, [
      {
        op: "setParagraphStyle",
        target: p1!,
        style: { spacing: { line: 1.5 } },
      },
      {
        op: "setParagraphStyle",
        target: p2!,
        style: { spacing: { before: 6 }, align: "center" },
      },
    ]);
    const xml = await bodyOf(engine);
    assert.ok(
      xml.includes(
        '<w:pPr><w:spacing w:before="100" w:line="360" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr>',
      ),
      xml,
    );
    assert.ok(
      xml.includes(
        '<w:pPr><w:spacing w:line="480" w:lineRule="exact" w:before="120"/><w:jc w:val="center"/></w:pPr>',
      ),
      xml,
    );
    const [first, second] = await engine.getElements({}, signal);
    assert.deepEqual(first!.paragraphStyle!.spacing, {
      before: 5,
      line: 1.5,
      lineRule: "auto",
    });
    assert.deepEqual(second!.paragraphStyle!.spacing, {
      before: 6,
      line: 24,
      lineRule: "exact",
    });
    await engine.dispose();
  });
});

describe("regressions for the defects the audit found", () => {
  it("a note reference inside a run survives an edit of the text around it", async () => {
    const body =
      `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>abc</w:t><w:footnoteReference w:id="1"/><w:t>def</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, footnotes: FOOTNOTES }));
    const [p] = await ids(engine);
    // "b" replaced: the reference lies inside the kept tail "c[note]def".
    await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 1, 2) },
    ]);
    const xml = await bodyOf(engine);
    assert.equal((await texts(engine))[0], "aXcdef");
    assert.equal(
      count(xml, "<w:footnoteReference"),
      1,
      `the footnote reference was dropped: ${xml}`,
    );
    assert.match(
      xml,
      /<w:t>c<\/w:t><w:footnoteReference w:id="1"\/><w:t>def<\/w:t>/,
      xml,
    );
    await engine.dispose();
  });

  it("a caret exactly at a note reference inside a run writes the reference once", async () => {
    const body =
      `<w:p><w:r><w:t>abc</w:t><w:footnoteReference w:id="1"/><w:t>def</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, footnotes: FOOTNOTES }));
    const [p] = await ids(engine);
    await apply(engine, [
      { op: "replaceText", target: p!, text: "+", range: range(p!, 3, 3) },
    ]);
    const xml = await bodyOf(engine);
    assert.equal((await texts(engine))[0], "abc+def");
    assert.equal(
      count(xml, "<w:footnoteReference"),
      1,
      `the footnote reference was duplicated: ${xml}`,
    );
    await engine.dispose();
  });

  it("setTextStyle keeps a note reference that sits between the texts of a run", async () => {
    const body =
      `<w:p><w:r><w:t>abc</w:t><w:footnoteReference w:id="1"/><w:t>def</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, footnotes: FOOTNOTES }));
    const [p] = await ids(engine);
    await apply(engine, [
      { op: "setTextStyle", target: p!, style: { bold: true } },
    ]);
    const xml = await bodyOf(engine);
    assert.equal((await texts(engine))[0], "abcdef");
    assert.equal(
      count(xml, "<w:footnoteReference"),
      1,
      `the footnote reference was dropped by a whole-paragraph restyle: ${xml}`,
    );
    assert.match(
      xml,
      /<w:r><w:rPr><w:b\/><w:bCs\/><\/w:rPr><w:t>abc<\/w:t><w:footnoteReference w:id="1"\/><w:t>def<\/w:t><\/w:r>/,
      xml,
    );
    await engine.dispose();
  });

  it("findText with a pageRange returns only matches on those pages", async () => {
    const original = buildDocx({
      body:
        paragraph("alpha one") +
        paragraph("alpha two") +
        paragraph("alpha three") +
        sectPr(),
    });
    const probe = await open(original);
    const [a, b, c] = (await ids(probe)).map((id) => id.slice(2));
    await probe.dispose();
    const { session, end } = await docxSession(original, [
      [run(a!, "alpha one", 72, 72)],
      [run(b!, "alpha two", 72, 72)],
      [run(c!, "alpha three", 72, 72)],
    ]);
    try {
      const hits = (
        await session.findText("alpha", { maxResults: 2, pageRange: [1, 2] })
      ).items;
      assert.equal(hits.length, 2);
      // As in the PPTX engine, the page range bounds the matches, so the
      // two results are the ones on pages 1 and 2, each placed.
      assert.deepEqual(
        hits.map((hit) => [hit.elementIds[0], hit.pageIndex]),
        [
          [`p:${b}`, 1],
          [`p:${c}`, 2],
        ],
      );
    } finally {
      await end();
    }
  });

  it("setTextStyle keeps w:rPrChange last in the run properties", async () => {
    const body =
      `<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:rPr/></w:rPrChange></w:rPr><w:t>abc</w:t></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [p] = await ids(engine);
    await apply(engine, [
      { op: "setTextStyle", target: p!, style: { underline: true } },
    ]);
    const xml = await bodyOf(engine);
    assert.match(
      xml,
      /<w:rPr><w:b\/><w:u w:val="single"\/><w:rPrChange [^>]*><w:rPr\/><\/w:rPrChange><\/w:rPr>/,
      xml,
    );
    await engine.dispose();
  });

  it("a whole replacement of a text box host keeps the id list and the engine usable", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p>${RUN("host ")}${TEXT_BOX("box")}${RUN(" end")}</w:p>` +
          paragraph("after") +
          sectPr(),
      }),
    );
    const [host, other, after] = await ids(engine);
    const change = await apply(engine, [
      { op: "replaceText", target: host!, text: "replaced" },
    ]);
    // The anchored drawing is covered whole and goes, with its id reported.
    assert.deepEqual(change.removedIds, [other]);
    assert.deepEqual(await ids(engine), [host, after]);
    assert.deepEqual(await texts(engine), ["replaced", "after"]);
    const shown = paraIds(
      await documentXml(await engine.materialize("show", {}, signal)),
    );
    assert.deepEqual(shown, [host!.slice(2), after!.slice(2)]);
    await engine.dispose();
  });

  it("a split of a text box host keeps the id list and the engine usable", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p>${RUN("host ")}${TEXT_BOX("box")}${RUN(" end")}</w:p>` +
          paragraph("after") +
          sectPr(),
      }),
    );
    const [host, , after] = await ids(engine);
    const change = await apply(engine, [
      { op: "replaceText", target: host!, text: "a\nb" },
    ]);
    assert.equal(change.createdIds.length, 1);
    assert.deepEqual(await ids(engine), [host, change.createdIds[0], after]);
    await engine.dispose();
  });

  it("the session recovers from a failed text box host replacement", async () => {
    const original = buildDocx({
      body:
        `<w:p>${RUN("host ")}${TEXT_BOX("box")}${RUN(" end")}</w:p>` +
        paragraph("after") +
        sectPr(),
    });
    const probe = await open(original);
    const [host, , after] = (await ids(probe)).map((id) => id.slice(2));
    await probe.dispose();
    const { session, end } = await docxSession(original, [
      [run(host!, "host  end", 72, 72), run(after!, "after", 72, 90)],
    ]);
    try {
      // Whatever the outcome, the receipt or the error must leave a usable session.
      const receipt = await session
        .replaceText({ target: `p:${host}`, text: "replaced" })
        .catch((error: unknown) => error);
      assert.ok(
        !(receipt instanceof Error),
        `replaceText on a text box host failed: ${String(receipt)}`,
      );
      assert.deepEqual(
        (await session.getElements()).items.map((item) => item.text),
        ["replaced", "after"],
      );
    } finally {
      await end();
    }
  });

  it("a range starting at an anchored drawing's run keeps it like a range ending there", async () => {
    const body =
      `<w:p>${RUN("abc")}${inlinePicture(100, 100, "rId9", true)}${RUN("def")}</w:p>` +
      sectPr();
    const engine = await open(
      buildDocx({
        body,
        media: [{ name: "word/media/image1.png", data: PNG }],
      }),
    );
    const [p, other] = await ids(engine);
    const change = await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 3, 6) },
    ]);
    const xml = await bodyOf(engine);
    // Either the drawing stays, or its removal is reported: never a silent drop.
    if (xml.includes("<wp:anchor")) {
      assert.deepEqual(change.removedIds, []);
      assert.deepEqual(await ids(engine), [p, other]);
    } else assert.deepEqual(change.removedIds, [other]);
    await engine.dispose();
  });

  it("a range starting right after a note reference in its own run keeps the reference", async () => {
    const body =
      `<w:p>${RUN("abc")}<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r>${RUN("def")}</w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body, footnotes: FOOTNOTES }));
    const [p] = await ids(engine);
    await apply(engine, [
      { op: "replaceText", target: p!, text: "X", range: range(p!, 3, 6) },
    ]);
    const xml = await bodyOf(engine);
    assert.equal(count(xml, "<w:footnoteReference"), 1, xml);
    assert.equal((await texts(engine))[0], "abcX");
    await engine.dispose();
  });

  it("a field nested in another field's result reads as the outer result", async () => {
    const body =
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> QUOTE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>7</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>` +
      `<w:r><w:t xml:space="preserve"> pages</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [element] = await engine.getElements({}, signal);
    assert.equal(element!.text, "7 pages");
    await engine.dispose();
  });

  it("a table placed after a section's last paragraph takes the next section's width", async () => {
    const engine = await open(
      buildDocx({
        body:
          paragraph("one") +
          `<w:p><w:pPr>${sectPr({ width: 8000, margin: 500 })}</w:pPr><w:r><w:t>ends section one</w:t></w:r></w:p>` +
          paragraph("two") +
          sectPr(),
      }),
    );
    const [, section] = await ids(engine);
    await apply(engine, [
      { op: "insertTable", after: section!, rows: [["a", "b"]] },
    ]);
    const xml = await bodyOf(engine);
    // The body's section is 9360 twips wide; the table sits in it.
    assert.ok(
      xml.includes(
        '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>',
      ),
      xml.slice(xml.indexOf("<w:tblGrid>"), xml.indexOf("</w:tblGrid>") + 12),
    );
    await engine.dispose();
  });

  it("deleting the only table of a body leaves a paragraph behind", async () => {
    const engine = await open(buildDocx({ body: TABLE(["only"]) + sectPr() }));
    const [table] = await ids(engine);
    const change = await apply(engine, [
      { op: "deleteElement", target: table! },
    ]);
    assert.equal(change.createdIds.length, 1);
    const xml = await bodyOf(engine);
    assert.match(xml, /<w:p[ >]/, xml);
    assert.equal((await ids(engine)).length, 1);
    await engine.dispose();
  });

  it("a read-only paragraph cannot be moved, as its operations list says", async () => {
    const body =
      `<w:p><w:ins w:id="2" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins></w:p>` +
      paragraph("plain") +
      sectPr();
    const engine = await open(buildDocx({ body }));
    const [tracked, plain] = await ids(engine);
    assert.ok(
      !(await engine.getElement(tracked!, signal))!.operations.includes(
        "moveElement",
      ),
    );
    await expectIssue(
      engine,
      { op: "moveElement", target: tracked!, after: plain! },
      "invalid-target",
      "/target",
    );
    await engine.dispose();
  });

  it("a save after an undo to a checkpointed state keeps untouched paragraphs' bytes", async () => {
    const original = buildDocx({
      body: paragraph("one") + paragraph("two") + paragraph("three") + sectPr(),
    });
    const probe = await open(original);
    const [a, b, c] = (await ids(probe)).map((id) => id.slice(2));
    await probe.dispose();
    // A history limit of 4 checkpoints every state.
    const { session, end } = await sessionWithLimits(
      original,
      [
        [
          run(a!, "one", 72, 72),
          run(b!, "two", 72, 90),
          run(c!, "three", 72, 108),
        ],
      ],
      { ...limits, maxEditHistory: 4 },
    );
    try {
      await session.replaceText({ target: `p:${a}`, text: "ONE" });
      const afterOne = (await session.save()).bytes;
      await session.replaceText({ target: `p:${b}`, text: "TWO" });
      await session.undo();
      const saved = (await session.save()).bytes;
      const xml = await documentXml(saved);
      // Only the rebuilt paragraph carries an id; "two" and "three" keep their bytes.
      assert.ok(xml.includes(paragraph("two")), xml);
      assert.ok(xml.includes(paragraph("three")), xml);
      assert.deepEqual(paraIds(xml), [a]);
      assert.deepEqual(saved, afterOne);
    } finally {
      await end();
    }
  });
});
