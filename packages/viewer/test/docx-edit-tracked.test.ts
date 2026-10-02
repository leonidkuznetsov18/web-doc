import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { prepareDocxForDisplay } from "../src/adapters/docx-prepass.js";
import { DocxEditEngine } from "../src/edit/docx/engine.js";
import type { DocxOperation } from "../src/edit/docx/types.js";
import type { BatchMode } from "../src/edit/engine.js";
import { OoxmlPackage } from "../src/edit/ooxml/package.js";
import { defaultResourceLimits, ViewerError } from "../src/index.js";
import {
  buildDocx,
  inlinePicture,
  paragraph,
  sectPr,
} from "./fixtures/docx-builder.js";
import { docxSession } from "./fixtures/docx-session.js";

/*
 * Task 63 of the ai-edit module: with `changeMode: "tracked"` the DOCX
 * engine writes a batch as Word revisions (`w:ins`, `w:del`, `w:delText`,
 * `w:rPrChange`, `w:pPrChange`) with the batch's author and date and
 * fresh ids, leaves untouched paragraphs byte for byte, lists the accepted
 * text, locks a paragraph that holds a revision, refuses what has no
 * tracked form, and replays the same bytes from the history.
 */

const limits = defaultResourceLimits;
const signal = new AbortController().signal;
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 1, 2, 3]);
const TRACKED: BatchMode = {
  changeMode: "tracked",
  author: "Agent <AI>",
  timestamp: "2026-10-02T10:00:00Z",
};
const REV = (id: number, date = true): string =>
  ` w:id="${id}" w:author="Agent &lt;AI>"${date ? ' w:date="2026-10-02T10:00:00Z"' : ""}`;

const RUN = (text: string, rPr = ""): string =>
  `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r>`;

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

function ids(engine: DocxEditEngine): Promise<string[]> {
  return engine
    .getElements({}, signal)
    .then((elements) => elements.map((element) => element.id));
}

function tracked(
  engine: DocxEditEngine,
  operations: readonly DocxOperation[],
  stateId = 1,
  mode: BatchMode = TRACKED,
) {
  return engine.apply({ stateId, operations, ...mode }, signal);
}

async function expectIssue(
  engine: DocxEditEngine,
  operation: DocxOperation,
  code: string,
  path?: string,
  mode: BatchMode = TRACKED,
): Promise<void> {
  const issues = await engine.validate([operation], signal, mode);
  assert.ok(
    issues.some(
      (issue) => issue.code === code && (!path || issue.path === path),
    ),
    `expected ${code} at ${path ?? "any"}, got ${JSON.stringify(issues)}`,
  );
}

describe("DOCX tracked changes (ai-edit T63)", () => {
  it("writes replaceText as a deletion and an insertion, keeps the rest and locks the paragraph", async () => {
    const original = buildDocx({
      body:
        `<w:p><w:bookmarkStart w:id="4" w:name="m"/>${RUN("Hello ", "<w:rPr><w:i/></w:rPr>")}${RUN("world")}<w:bookmarkEnd w:id="4"/></w:p>` +
        paragraph("Untouched") +
        sectPr(),
    });
    const engine = await open(original);
    const [first, second] = await ids(engine);
    const change = await tracked(engine, [
      { op: "replaceText", target: first!, text: "Bye" },
    ]);
    assert.deepEqual(change.createdIds, []);
    const body = await bodyOf(engine);
    assert.equal(
      body,
      `<w:p w14:paraId="${first!.slice(2)}"><w:bookmarkStart w:id="4" w:name="m"/>` +
        `<w:del${REV(5)}><w:r><w:rPr><w:i/></w:rPr><w:delText xml:space="preserve">Hello </w:delText></w:r><w:r><w:delText>world</w:delText></w:r></w:del>` +
        `<w:ins${REV(6)}><w:r><w:rPr><w:i/></w:rPr><w:t>Bye</w:t></w:r></w:ins>` +
        `<w:bookmarkEnd w:id="4"/></w:p>` +
        paragraph("Untouched") +
        sectPr(),
    );
    // The accepted text is what the element reads; the paragraph is locked.
    const [element] = await engine.getElements({}, signal);
    assert.equal(element!.text, "Bye");
    assert.equal(element!.readOnlyReason, "tracked-changes");
    assert.deepEqual(await ids(engine), [first, second]);
    await expectIssue(
      engine,
      { op: "replaceText", target: first!, text: "again" },
      "invalid-target",
      "/target",
    );
    assert.deepEqual(await engine.revisions(first!, signal), [
      {
        kind: "del",
        scope: "runs",
        id: 5,
        author: "Agent <AI>",
        date: "2026-10-02T10:00:00Z",
        text: "Hello world",
      },
      {
        kind: "ins",
        scope: "runs",
        id: 6,
        author: "Agent <AI>",
        date: "2026-10-02T10:00:00Z",
        text: "Bye",
      },
    ]);
    assert.deepEqual(await engine.revisions(second!, signal), []);
    assert.deepEqual(await engine.revisions("tbl:nope", signal), []);
    // The display copy shows the accepted text.
    const shown = await prepareDocxForDisplay(
      await engine.materialize("show", {}, signal),
      limits,
    );
    const pkg = await OoxmlPackage.open(shown.bytes, { limits });
    const xml = new TextDecoder().decode(await pkg.part("/word/document.xml"));
    assert.ok(xml.includes("<w:t>Bye</w:t>"), xml);
    await engine.dispose();
  });

  it("splits a cut run, keeps its style on the insertion, and handles a caret and a pure deletion", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p>${RUN("abcdef", "<w:rPr><w:b/></w:rPr>")}</w:p>` +
          `<w:p>${RUN("one two")}</w:p>` +
          `<w:p>${RUN("gone")}</w:p>` +
          sectPr(),
      }),
    );
    const [a, b, c] = await ids(engine);
    const range = (id: string, start: number, end: number) => ({
      start: { elementId: id, offset: start },
      end: { elementId: id, offset: end },
    });
    await tracked(engine, [
      { op: "replaceText", target: a!, text: "XY", range: range(a!, 2, 4) },
      { op: "replaceText", target: b!, text: " and", range: range(b!, 3, 3) },
      { op: "replaceText", target: c!, text: "", range: range(c!, 0, 4) },
    ]);
    const body = await bodyOf(engine);
    assert.ok(
      body.includes(
        `<w:r><w:rPr><w:b/></w:rPr><w:t>ab</w:t></w:r><w:del${REV(1)}><w:r><w:rPr><w:b/></w:rPr><w:delText>cd</w:delText></w:r></w:del><w:ins${REV(2)}><w:r><w:rPr><w:b/></w:rPr><w:t>XY</w:t></w:r></w:ins><w:r><w:rPr><w:b/></w:rPr><w:t>ef</w:t></w:r>`,
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:r><w:t>one</w:t></w:r><w:ins${REV(3)}><w:r><w:t xml:space="preserve"> and</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> two</w:t></w:r>`,
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:del${REV(4)}><w:r><w:delText>gone</w:delText></w:r></w:del></w:p>`,
      ),
      body,
    );
    assert.ok(!body.includes(`<w:ins${REV(5)}`));
    const texts = (await engine.getElements({}, signal)).map((e) => e.text);
    assert.deepEqual(texts, ["abXYef", "one and two", ""]);
    await engine.dispose();
  });

  it("splits a paragraph with inserted paragraph marks and leaves the original mark last", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr>${RUN("head tail")}</w:p>` +
          sectPr(),
      }),
    );
    const [p] = await ids(engine);
    const change = await tracked(engine, [
      {
        op: "replaceText",
        target: p!,
        text: "A\nB\nC",
        range: {
          start: { elementId: p!, offset: 4 },
          end: { elementId: p!, offset: 5 },
        },
      },
    ]);
    assert.equal(change.createdIds.length, 2);
    const [b, c] = change.createdIds;
    const body = await bodyOf(engine);
    assert.equal(
      body,
      `<w:p w14:paraId="${p!.slice(2)}"><w:pPr><w:jc w:val="center"/><w:rPr><w:ins${REV(2)}/><w:b/></w:rPr></w:pPr>` +
        `<w:r><w:t>head</w:t></w:r><w:del${REV(1)}><w:r><w:delText xml:space="preserve"> </w:delText></w:r></w:del><w:ins${REV(3)}><w:r><w:t>A</w:t></w:r></w:ins></w:p>` +
        `<w:p w14:paraId="${b!.slice(2)}"><w:pPr><w:jc w:val="center"/><w:rPr><w:ins${REV(4)}/><w:b/></w:rPr></w:pPr><w:ins${REV(5)}><w:r><w:t>B</w:t></w:r></w:ins></w:p>` +
        `<w:p w14:paraId="${c!.slice(2)}"><w:pPr><w:jc w:val="center"/><w:rPr><w:b/></w:rPr></w:pPr><w:ins${REV(6)}><w:r><w:t>C</w:t></w:r></w:ins><w:r><w:t>tail</w:t></w:r></w:p>` +
        sectPr(),
    );
    assert.deepEqual(
      (await engine.getElements({}, signal)).map((e) => [
        e.text,
        e.readOnlyReason,
      ]),
      [
        ["headA", "tracked-changes"],
        ["B", "tracked-changes"],
        ["Ctail", "tracked-changes"],
      ],
    );
    const marks = await engine.revisions(p!, signal);
    assert.deepEqual(
      marks.map((revision) => [revision.kind, revision.scope, revision.text]),
      [
        ["ins", "mark", "\n"],
        ["del", "runs", " "],
        ["ins", "runs", "A"],
      ],
    );
    await engine.dispose();
  });

  it("refuses what has no tracked form and an author the file cannot carry", async () => {
    const engine = await open(
      buildDocx({
        media: [{ name: "word/media/image1.png", data: PNG }],
        body:
          `<w:p>${RUN("a ")}<w:hyperlink r:id="rId9">${RUN("link")}</w:hyperlink>${RUN(" z")}</w:p>` +
          paragraph("plain") +
          `<w:p>${RUN("pic ")}${inlinePicture(914400, 457200)}</w:p>` +
          `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${paragraph("cell")}</w:tc></w:tr></w:tbl>` +
          paragraph("last") +
          sectPr(),
      }),
    );
    const [linked, plain, pictured, picture, table] = await ids(engine);
    const range = (id: string, start: number, end: number) => ({
      start: { elementId: id, offset: start },
      end: { elementId: id, offset: end },
    });
    await expectIssue(
      engine,
      {
        op: "replaceText",
        target: linked!,
        text: "x",
        range: range(linked!, 2, 6),
      },
      "unsupported-change-mode",
      "/range",
    );
    // Next to the hyperlink, not inside it: allowed.
    const beside: DocxOperation = {
      op: "replaceText",
      target: linked!,
      text: "A",
      range: range(linked!, 0, 1),
    };
    assert.deepEqual(await engine.validate([beside], signal, TRACKED), []);
    await expectIssue(
      engine,
      { op: "deleteElement", target: linked! },
      "unsupported-change-mode",
      "/target",
    );
    await expectIssue(
      engine,
      { op: "deleteElement", target: table! },
      "unsupported-change-mode",
      "/target",
    );
    await expectIssue(
      engine,
      { op: "deleteElement", target: picture! },
      "unsupported-change-mode",
      "/target",
    );
    await expectIssue(
      engine,
      { op: "moveElement", target: plain!, after: pictured! },
      "unsupported-change-mode",
      "",
    );
    await expectIssue(
      engine,
      { op: "insertTable", after: plain!, rows: [["a"]] },
      "unsupported-change-mode",
      "",
    );
    await expectIssue(
      engine,
      {
        op: "insertImage",
        after: plain!,
        data: PNG,
        mimeType: "image/png",
        size: { width: 10, height: 10 },
      },
      "unsupported-change-mode",
      "",
    );
    await expectIssue(
      engine,
      { op: "replaceText", target: plain!, text: "x" },
      "invalid-value",
      "/author",
      { changeMode: "tracked", author: "bad\u0000name" },
    );
    // The same operations apply directly.
    const direct: DocxOperation = { op: "deleteElement", target: table! };
    assert.deepEqual(await engine.validate([direct], signal), []);
    await engine.dispose();
  });

  it("inserts paragraphs, deletes a paragraph and sets a cell under revision", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p><w:pPr><w:pStyle w:val="Quote"/></w:pPr>${RUN("ref", "<w:rPr><w:i/></w:rPr>")}</w:p>` +
          `<w:p>${RUN("doomed")}<w:bookmarkStart w:id="9" w:name="b"/>${RUN(" end")}<w:bookmarkEnd w:id="9"/></w:p>` +
          `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc>${paragraph("c1")}${paragraph("c2")}</w:tc></w:tr></w:tbl>` +
          paragraph("last") +
          sectPr(),
      }),
    );
    const [ref, doomed, table] = await ids(engine);
    const change = await tracked(engine, [
      {
        op: "insertParagraph",
        after: ref!,
        text: "new\nalso",
        style: { bold: true },
      },
      { op: "deleteElement", target: doomed! },
      { op: "setTableCell", target: table!, row: 0, column: 0, text: "C" },
    ]);
    assert.equal(change.createdIds.length, 2);
    assert.deepEqual(change.removedIds, []);
    const [n1, n2] = change.createdIds;
    const body = await bodyOf(engine);
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${n1!.slice(2)}"><w:pPr><w:pStyle w:val="Quote"/><w:rPr><w:ins${REV(10)}/></w:rPr></w:pPr><w:ins${REV(11)}><w:r><w:rPr><w:b/><w:bCs/><w:i/></w:rPr><w:t>new</w:t></w:r></w:ins></w:p>` +
          `<w:p w14:paraId="${n2!.slice(2)}"><w:pPr><w:pStyle w:val="Quote"/><w:rPr><w:ins${REV(12)}/></w:rPr></w:pPr><w:ins${REV(13)}><w:r><w:rPr><w:b/><w:bCs/><w:i/></w:rPr><w:t>also</w:t></w:r></w:ins></w:p>`,
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:p w14:paraId="${doomed!.slice(2)}"><w:pPr><w:rPr><w:del${REV(14)}/></w:rPr></w:pPr><w:del${REV(15)}><w:r><w:delText xml:space="preserve">doomed</w:delText></w:r></w:del><w:bookmarkStart w:id="9" w:name="b"/><w:del${REV(16)}><w:r><w:delText xml:space="preserve"> end</w:delText></w:r></w:del><w:bookmarkEnd w:id="9"/></w:p>`,
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:del${REV(17)}><w:r><w:delText>c1</w:delText></w:r></w:del><w:ins${REV(18)}><w:r><w:t>C</w:t></w:r></w:ins></w:p>`,
      ),
      body,
    );
    assert.ok(
      body.includes(
        `<w:pPr><w:rPr><w:del${REV(19)}/></w:rPr></w:pPr><w:del${REV(20)}><w:r><w:delText>c2</w:delText></w:r></w:del></w:p>`,
      ),
      body,
    );
    const elements = await engine.getElements({}, signal);
    assert.deepEqual(
      elements.map((e) => [e.kind, e.text ?? "", e.readOnlyReason ?? ""]),
      [
        ["paragraph", "ref", ""],
        ["paragraph", "new", "tracked-changes"],
        ["paragraph", "also", "tracked-changes"],
        ["paragraph", "", "tracked-changes"],
        ["table", "C\n", ""],
        ["paragraph", "C", "tracked-changes"],
        ["paragraph", "", "tracked-changes"],
        ["paragraph", "last", ""],
      ],
    );
    await engine.dispose();
  });

  it("records previous properties in w:rPrChange and w:pPrChange, leaving the paragraph editable", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="120"/><w:rPr><w:i/></w:rPr></w:pPr>${RUN("ab", "<w:rPr><w:b/></w:rPr>")}${RUN("cd")}</w:p>` +
          sectPr(),
      }),
    );
    const [p] = await ids(engine);
    await tracked(engine, [
      { op: "setTextStyle", target: p!, style: { underline: true } },
      { op: "setParagraphStyle", target: p!, style: { align: "center" } },
    ]);
    const body = await bodyOf(engine);
    assert.equal(
      body,
      `<w:p w14:paraId="${p!.slice(2)}"><w:pPr><w:jc w:val="center"/><w:spacing w:after="120"/><w:rPr><w:i/><w:u w:val="single"/><w:rPrChange${REV(3)}><w:rPr><w:i/></w:rPr></w:rPrChange></w:rPr><w:pPrChange${REV(4)}><w:pPr><w:jc w:val="left"/><w:spacing w:after="120"/></w:pPr></w:pPrChange></w:pPr>` +
        `<w:r><w:rPr><w:b/><w:u w:val="single"/><w:rPrChange${REV(1)}><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>ab</w:t></w:r>` +
        `<w:r><w:rPr><w:u w:val="single"/><w:rPrChange${REV(2)}><w:rPr/></w:rPrChange></w:rPr><w:t>cd</w:t></w:r></w:p>` +
        sectPr(),
    );
    const [element] = await engine.getElements({}, signal);
    assert.equal(element!.readOnlyReason, undefined);
    assert.equal(element!.textStyle!.underline, true);
    assert.deepEqual(
      (await engine.revisions(p!, signal)).map((r) => [
        r.kind,
        r.scope,
        r.id,
        r.text,
      ]),
      [
        ["rPrChange", "mark", 3, undefined],
        ["pPrChange", "paragraph", 4, undefined],
        ["rPrChange", "runs", 1, "ab"],
        ["rPrChange", "runs", 2, "cd"],
      ],
    );
    await engine.dispose();
  });

  it("leaves the date out without a timestamp and starts ids after the file's own", async () => {
    const engine = await open(
      buildDocx({
        body:
          `<w:p><w:bookmarkStart w:id="40" w:name="b"/>${RUN("x")}<w:bookmarkEnd w:id="40"/></w:p>` +
          sectPr(),
      }),
    );
    const [p] = await ids(engine);
    await tracked(engine, [{ op: "replaceText", target: p!, text: "y" }], 1, {
      changeMode: "tracked",
      author: "Agent <AI>",
    });
    const body = await bodyOf(engine);
    assert.ok(body.includes(`<w:del${REV(41, false)}>`), body);
    assert.ok(body.includes(`<w:ins${REV(42, false)}>`), body);
    assert.ok(!body.includes("w:date="));
    await engine.dispose();
  });

  it("replays tracked batches from the history through the session and reads revisions", async () => {
    const bytes = buildDocx({
      body: paragraph("Hello") + paragraph("World") + sectPr(),
    });
    const { session, end } = await docxSession(bytes, [[]]);
    try {
      const [hello] = (await session.getElements()).items.map((e) => e.id);
      await assert.rejects(
        session.replaceText(
          { target: hello!, text: "Bye" },
          { changeMode: "tracked" },
        ),
        (error: unknown) =>
          error instanceof ViewerError &&
          error.code === "invalid-operation" &&
          (error.details?.issues as { path: string }[])[0]!.path === "/author",
      );
      const receipt = await session.replaceText(
        { target: hello!, text: "Bye" },
        {
          changeMode: "tracked",
          author: "Agent",
          timestamp: "2026-10-02T10:00:00Z",
        },
      );
      assert.equal(receipt.revision, 1);
      const first = await session.save();
      const revisions = await session.getRevisions(hello!);
      assert.equal(revisions.revision, 1);
      assert.deepEqual(
        revisions.items.map((r) => [r.kind, r.author, r.text]),
        [
          ["del", "Agent", "Hello"],
          ["ins", "Agent", "Bye"],
        ],
      );
      const element = await session.getElement(hello!);
      assert.equal(element.item?.text, "Bye");
      assert.equal(element.item?.readOnlyReason, "tracked-changes");
      await session.undo();
      assert.deepEqual((await session.getRevisions(hello!)).items, []);
      await session.redo();
      const again = await session.save();
      assert.deepEqual(again.bytes, first.bytes);
      // The tool path carries the mode from its options.
      await session.undo();
      const tool = await session.callTool(
        {
          name: "document_apply",
          arguments: {
            operations: [{ op: "replaceText", target: hello!, text: "Bye" }],
          },
        },
        { changeMode: "tracked", author: "Agent" },
      );
      assert.equal(tool.ok, true, tool.text);
      assert.equal((await session.getRevisions(hello!)).items.length, 2);
      const outline = await session.getOutline();
      assert.equal(outline.items[0]!.readOnlyReason, "tracked-changes");
    } finally {
      await end();
    }
  });
});
