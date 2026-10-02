import assert from "node:assert/strict";
import { describe, it } from "node:test";

import { buildToolSet } from "../src/edit/ai/tools.js";
import { docxOperationSchemas } from "../src/edit/docx/schemas.js";
import { pdfOperationSchemas } from "../src/edit/pdf/schemas.js";
import { pptxOperationSchemas } from "../src/edit/pptx/schemas.js";
import { assertSupportedSchema } from "../src/edit/schema.js";
import { EditSessionController } from "../src/edit/session.js";
import {
  describeReceipt,
  ViewerError,
  type EditOperation,
  type EditReceipt,
  type JsonSchema,
  type ToolResult,
} from "../src/index.js";
import { buildDocx, paragraph, sectPr } from "./fixtures/docx-builder.js";
import { docxSession } from "./fixtures/docx-session.js";
import {
  encodePages,
  FakeEditEngine,
  FakeHost,
  fakeSchemas,
} from "./fixtures/fake-edit-engine.js";
import { buildPdf } from "./fixtures/pdf-builder.js";
import { pdfSession } from "./fixtures/pdf-session.js";

/*
 * Task 62 of the ai-edit module: the tool set is the editing API as a
 * model provider lists tools; `callTool()` validates a call, runs it and
 * answers with JSON and a sentence, refusing a model's mistakes without
 * throwing; `describeReceipt()` puts a receipt into words.
 */

const TOOL_NAMES = [
  "document_describe",
  "document_outline",
  "document_find",
  "document_inspect",
  "document_preview",
  "document_apply",
  "document_undo",
  "document_checkpoint",
];

function receipt(fields: Partial<EditReceipt> = {}): EditReceipt {
  return {
    sessionId: "s",
    revision: 3,
    dryRun: false,
    operationCount: 1,
    createdIds: [],
    removedIds: [],
    changedPages: [0],
    pageCount: 2,
    warnings: [],
    ...fields,
  };
}

function session() {
  const original = encodePages(["one two", "three"]);
  const engine = new FakeEditEngine(original);
  const host = new FakeHost();
  const core = new EditSessionController(engine, host, original, 2);
  return { core, engine, host };
}

function issueCodes(result: ToolResult): string[] {
  return (result.issues ?? []).map((issue) => `${issue.code}@${issue.path}`);
}

describe("tool set", () => {
  it("lists the same tools for every format with valid schemas", () => {
    for (const schemas of [
      pdfOperationSchemas,
      pptxOperationSchemas,
      docxOperationSchemas,
      fakeSchemas,
    ]) {
      const tools = buildToolSet(schemas.format, schemas);
      assert.equal(tools.version, 1);
      assert.equal(tools.format, schemas.format);
      assert.deepEqual(
        tools.definitions.map((tool) => tool.name),
        TOOL_NAMES,
      );
      assert.equal(Object.isFrozen(tools.definitions), true);
      for (const tool of tools.definitions) {
        assert.ok(tool.description.length > 20, tool.name);
        assertSupportedSchema(tool.inputSchema);
      }
      // The batch schema offers every operation of the format as one oneOf,
      // with the shared definitions hoisted to the root.
      const apply = tools.definitions.find(
        (tool) => tool.name === "document_apply",
      )!;
      const operations = (
        apply.inputSchema.properties as Record<string, JsonSchema>
      ).operations!;
      const branches = (operations.items as JsonSchema).oneOf as JsonSchema[];
      assert.deepEqual(
        branches.map(
          (branch) =>
            ((branch.properties as Record<string, JsonSchema>).op as JsonSchema)
              .const,
        ),
        Object.keys(schemas.operations),
      );
      for (const branch of branches) assert.equal("$defs" in branch, false);
      if (schemas !== fakeSchemas) assert.ok(apply.inputSchema.$defs);
      const preview = tools.definitions.find(
        (tool) => tool.name === "document_preview",
      )!;
      assert.equal(
        "label" in (preview.inputSchema.properties as object),
        false,
      );
    }
  });

  it("narrows binary fields to asset references and says which mode the format has", () => {
    const pdf = buildToolSet("pdf", pdfOperationSchemas);
    const apply = pdf.definitions.find(
      (tool) => tool.name === "document_apply",
    )!;
    const branches = (
      (apply.inputSchema.properties as Record<string, JsonSchema>).operations!
        .items as JsonSchema
    ).oneOf as JsonSchema[];
    const image = branches.find(
      (branch) =>
        ((branch.properties as Record<string, JsonSchema>).op as JsonSchema)
          .const === "insertImage",
    )!;
    const data = (image.properties as Record<string, JsonSchema>).data!;
    assert.equal(data.pattern, "^asset:[0-9a-f]{64}$");
    assert.equal("x-binary" in data, false);
    assert.match(apply.description, /PDF has no tracked changes/);
    const docx = buildToolSet("docx", docxOperationSchemas);
    assert.match(
      docx.definitions.find((tool) => tool.name === "document_apply")!
        .description,
      /tracked/,
    );
  });
});

describe("callTool", () => {
  it("refuses unknown tools and bad arguments without throwing", async () => {
    const { core } = session();
    try {
      const unknown = await core.callTool({
        name: "document_delete",
        arguments: {},
      });
      assert.equal(unknown.ok, false);
      assert.deepEqual(issueCodes(unknown), ["unknown-tool@/name"]);
      assert.match(unknown.text, /document_describe, document_outline/);
      assert.deepEqual(unknown.content, { issues: unknown.issues });
      const bad = await core.callTool({
        name: "document_find",
        arguments: { text: "x", maxResults: 0, extra: true },
      });
      assert.equal(bad.ok, false);
      assert.deepEqual(
        issueCodes(bad).sort(),
        ["additional-property@/extra", "minimum@/maxResults"].sort(),
      );
      assert.match(bad.text, /^Invalid arguments for document_find/);
      const noArgs = await core.callTool({
        name: "document_inspect",
        arguments: undefined,
      });
      assert.deepEqual(issueCodes(noArgs), ["required@/id"]);
    } finally {
      await core.end();
    }
  });

  it("reads through the tools with text for the chat", async () => {
    const { core } = session();
    try {
      const described = await core.callTool({
        name: "document_describe",
        arguments: {},
      });
      assert.equal(described.ok, true);
      const description = described.content as {
        revision: number;
        text: string;
        elementCount: number;
      };
      assert.equal(description.revision, 0);
      assert.equal(description.text.split("\n")[0], "pdf: 2 pages, 3 elements");
      assert.equal(
        described.text,
        "Described 3 elements on 2 pages at revision 0.",
      );
      const outlined = await core.callTool({
        name: "document_outline",
        arguments: { pageRange: [1, 1] },
      });
      assert.equal((outlined.content as { nodes: unknown[] }).nodes.length, 1);
      assert.equal(outlined.text, "Outlined 1 element at revision 0.");
      const found = await core.callTool({
        name: "document_find",
        arguments: { text: "three" },
      });
      assert.equal(found.ok, true);
      assert.equal(
        (found.content as { candidates: unknown[] }).candidates.length,
        1,
      );
      assert.equal(
        found.text,
        "Found 1 candidate; best: exact match in p1w0 on page 2 (score 1).",
      );
      const missed = await core.callTool({
        name: "document_find",
        arguments: { text: "nothing like this" },
      });
      assert.equal(missed.ok, true);
      assert.match(missed.text, /^Nothing matched/);
      const inspected = await core.callTool({
        name: "document_inspect",
        arguments: { id: "p0w1" },
      });
      assert.equal(inspected.ok, true);
      const element = inspected.content as Record<string, unknown>;
      assert.equal(element.kind, "word");
      assert.equal(element.text, "two");
      assert.equal("bounds" in element, false);
      assert.equal(
        inspected.text,
        "word p0w1 on page 1: two. Operations: setText.",
      );
      const gone = await core.callTool({
        name: "document_inspect",
        arguments: { id: "p9w9" },
      });
      assert.equal(gone.ok, false);
      assert.deepEqual(issueCodes(gone), ["unknown-target@/id"]);
    } finally {
      await core.end();
    }
  });

  it("previews, applies, undoes and refuses with the batch's issues", async () => {
    const { core, host } = session();
    try {
      const op = { op: "setText", pageIndex: 1, text: "THREE" };
      const preview = await core.callTool({
        name: "document_preview",
        arguments: { operations: [op] },
      });
      assert.equal(preview.ok, true);
      assert.equal((preview.content as EditReceipt).dryRun, true);
      assert.equal(
        preview.text,
        "Preview, nothing changed: Applied setText. Would repaint page 2; 2 pages after.",
      );
      assert.equal(host.shown.length, 0);
      assert.equal(core.state.revision, 0);
      const stale = await core.callTool(
        { name: "document_apply", arguments: { operations: [op] } },
        { expectedRevision: 4 },
      );
      assert.equal(stale.ok, false);
      assert.deepEqual(issueCodes(stale), ["edit-conflict@"]);
      assert.match(stale.text, /now revision 0/);
      const applied = await core.callTool(
        {
          name: "document_apply",
          arguments: { operations: [op], label: "shout" },
        },
        { expectedRevision: 0 },
      );
      assert.equal(applied.ok, true);
      assert.equal((applied.content as EditReceipt).revision, 1);
      assert.equal(
        applied.text,
        "Applied setText. Revision 1; changed page 2; 2 pages.",
      );
      assert.deepEqual(host.current, ["one two", "THREE"]);
      const invalid = await core.callTool({
        name: "document_apply",
        arguments: {
          operations: [
            { op: "setText", pageIndex: 7, text: "x" },
            { op: "nope" },
          ],
        },
      });
      assert.equal(invalid.ok, false);
      assert.deepEqual(issueCodes(invalid), ["unknown-operation@/op"]);
      assert.match(
        invalid.text,
        /^Refused: operation 1 at \/op: Unknown pdf operation nope/,
      );
      const empty = await core.callTool({
        name: "document_apply",
        arguments: { operations: [] },
      });
      assert.deepEqual(issueCodes(empty), ["min-items@/operations"]);
      const undone = await core.callTool({
        name: "document_undo",
        arguments: {},
      });
      assert.equal(undone.ok, true);
      assert.equal(undone.text, "Undid the last change; revision 2.");
      assert.deepEqual(host.current, ["one two", "three"]);
      const nothing = await core.callTool({
        name: "document_undo",
        arguments: {},
      });
      assert.equal(nothing.ok, true);
      assert.equal(nothing.text, "Nothing to undo.");
      const tracked = await core.callTool(
        { name: "document_apply", arguments: { operations: [op] } },
        { changeMode: "tracked", author: "Agent" },
      );
      assert.equal(tracked.ok, false);
      assert.deepEqual(issueCodes(tracked), ["unsupported-change-mode@"]);
      assert.match(tracked.text, /PDF has no tracked changes/);
    } finally {
      await core.end();
    }
  });

  it("takes binary payloads by asset reference only", async () => {
    const { core, host } = session();
    try {
      const raw = await core.callTool({
        name: "document_apply",
        arguments: {
          operations: [{ op: "stamp", pageIndex: 0, data: btoa("abc") }],
        },
      });
      assert.equal(raw.ok, false);
      assert.deepEqual(issueCodes(raw), ["unknown-asset@/data"]);
      assert.match(raw.text, /addAsset\(\)/);
      const reference = await core.addAsset(new Uint8Array([1, 2, 3]));
      const stamped = await core.callTool({
        name: "document_apply",
        arguments: {
          operations: [{ op: "stamp", pageIndex: 0, data: reference }],
        },
      });
      assert.equal(stamped.ok, true);
      assert.deepEqual(host.current, ["one two+3", "three"]);
      const unknown = await core.callTool({
        name: "document_apply",
        arguments: {
          operations: [
            { op: "stamp", pageIndex: 0, data: `asset:${"0".repeat(64)}` },
          ],
        },
      });
      assert.equal(unknown.ok, false);
      assert.deepEqual(issueCodes(unknown), ["unknown-asset@/data"]);
    } finally {
      await core.end();
    }
  });

  it("manages checkpoints through one tool", async () => {
    const { core, host } = session();
    try {
      const none = await core.callTool({
        name: "document_checkpoint",
        arguments: { action: "list" },
      });
      assert.equal(none.text, "No checkpoints.");
      const created = await core.callTool({
        name: "document_checkpoint",
        arguments: { action: "create", label: "start" },
      });
      assert.equal(created.ok, true);
      const { id } = created.content as { id: string };
      assert.equal(
        created.text,
        `Checkpoint ${id} created at revision 0 ("start").`,
      );
      await core.callTool({
        name: "document_apply",
        arguments: {
          operations: [{ op: "insertPage", index: 0, text: "front" }],
        },
      });
      const listed = await core.callTool({
        name: "document_checkpoint",
        arguments: { action: "list" },
      });
      assert.equal(listed.text, `1 checkpoint: ${id} (revision 0, "start").`);
      const noId = await core.callTool({
        name: "document_checkpoint",
        arguments: { action: "restore" },
      });
      assert.deepEqual(issueCodes(noId), ["required@/id"]);
      const missing = await core.callTool({
        name: "document_checkpoint",
        arguments: { action: "restore", id: "missing" },
      });
      assert.equal(missing.ok, false);
      assert.deepEqual(issueCodes(missing), ["unknown-checkpoint@/id"]);
      const restored = await core.callTool(
        { name: "document_checkpoint", arguments: { action: "restore", id } },
        { expectedRevision: 1 },
      );
      assert.equal(restored.ok, true);
      assert.equal(
        restored.text,
        `Restored checkpoint ${id}, removing 1 element; revision 2.`,
      );
      assert.deepEqual(host.current, ["one two", "three"]);
      const again = await core.callTool({
        name: "document_checkpoint",
        arguments: { action: "restore", id },
      });
      assert.equal(again.text, `Already at checkpoint ${id}; revision 2.`);
    } finally {
      await core.end();
    }
  });

  it("throws only for the session's own errors", async () => {
    const { core } = session();
    const controller = new AbortController();
    controller.abort();
    await assert.rejects(
      core.callTool(
        { name: "document_describe", arguments: {} },
        { signal: controller.signal },
      ),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "aborted",
    );
    await core.end();
    await assert.rejects(
      core.callTool({ name: "document_describe", arguments: {} }),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "lifecycle-error",
    );
  });
});

describe("describeReceipt", () => {
  it("names what each operation did, with ids when the receipt has one per creation", () => {
    const text = describeReceipt(
      "pdf",
      [
        { op: "insertTextBox", pageIndex: 0, rect: {}, text: "x" },
        { op: "deleteElement", target: "p0:o3" },
        {
          op: "insertTable",
          pageIndex: 1,
          rows: [
            ["a", "b"],
            ["c", "d"],
          ],
        },
        { op: "replaceText", target: "$0", text: "y", range: {} },
      ] as unknown as EditOperation[],
      receipt({
        createdIds: ["p0:n1", "p1:n2"],
        removedIds: ["p0:o3"],
        changedPages: [0, 1],
        warnings: [
          {
            code: "font-substitution",
            message: "Helvetica stands in for Arial",
          },
        ],
      }),
    );
    assert.equal(
      text,
      "Added a text box on page 1 (p0:n1). Deleted p0:o3. Added a 2×2 table on page 2 (p1:n2). Replaced a range of the text of the element created by operation 1. Removed p0:o3. Warning (font-substitution): Helvetica stands in for Arial. Revision 3; changed pages 1, 2; 2 pages.",
    );
  });

  it("speaks of slides, paragraphs and previews", () => {
    assert.equal(
      describeReceipt(
        "pptx",
        [
          { op: "deleteSlide", pageIndex: 2 },
          { op: "moveSlide", from: 0, to: 3 },
        ] as unknown as EditOperation[],
        receipt({
          removedIds: ["sld3:2", "sld3:3"],
          changedPages: [0, 1, 2, 3],
          pageCount: 4,
        }),
      ),
      "Deleted slide 3. Moved slide 1 to 4. Removed sld3:2, sld3:3. Revision 3; changed slides 1, 2, 3, 4; 4 slides.",
    );
    assert.equal(
      describeReceipt(
        "docx",
        [
          { op: "insertParagraph", after: "p:1A000000", text: "New" },
        ] as unknown as EditOperation[],
        receipt({
          dryRun: true,
          createdIds: ["p:2B000000"],
          changedPages: [0, 1, 2, 3, 4, 5, 6, 7],
          pageCount: 8,
        }),
      ),
      "Preview, nothing changed: Added a paragraph after p:1A000000 (p:2B000000). Would repaint 8 pages (1–8); 8 pages after.",
    );
    assert.equal(
      describeReceipt(
        "docx",
        [
          { op: "setTableCell", target: "tbl:1", row: 0, column: 1, text: "x" },
        ] as unknown as EditOperation[],
        receipt({
          changedPages: [],
          removedIds: ["a", "b", "c", "d", "e", "f"],
        }),
      ),
      "Set cell 0,1 of tbl:1. Removed 6 elements. Revision 3; no page changed; 2 pages.",
    );
  });
});

describe("tools on format sessions", () => {
  it("runs an agent turn on a PDF through the tools only", async () => {
    const { session: edit, end } = await pdfSession(
      await buildPdf(["Quarterly review of the north"]),
    );
    try {
      const described = await edit.callTool({
        name: "document_describe",
        arguments: {},
      });
      const revision = (described.content as { revision: number }).revision;
      const found = await edit.callTool({
        name: "document_find",
        arguments: { text: "north" },
      });
      const [best] = (
        found.content as { candidates: { elementId: string; range: unknown }[] }
      ).candidates;
      assert.ok(best);
      const checkpoint = await edit.callTool({
        name: "document_checkpoint",
        arguments: { action: "create", label: "turn" },
      });
      const { id } = checkpoint.content as { id: string };
      const applied = await edit.callTool(
        {
          name: "document_apply",
          arguments: {
            operations: [
              {
                op: "replaceText",
                target: best.elementId,
                range: best.range,
                text: "south",
              },
            ],
          },
        },
        { expectedRevision: revision },
      );
      assert.equal(applied.ok, true, applied.text);
      assert.equal(
        applied.text,
        `Replaced a range of the text of ${best.elementId}. Revision 1; changed page 1; 1 page.`,
      );
      const after = await edit.callTool({
        name: "document_find",
        arguments: { text: "south" },
      });
      assert.equal(
        (after.content as { candidates: unknown[] }).candidates.length,
        1,
      );
      const restored = await edit.callTool({
        name: "document_checkpoint",
        arguments: { action: "restore", id },
      });
      assert.equal(restored.ok, true);
      const back = await edit.callTool({
        name: "document_find",
        arguments: { text: "north" },
      });
      assert.equal(
        (back.content as { candidates: unknown[] }).candidates.length,
        1,
      );
    } finally {
      await end();
    }
  });

  it("refuses a tracked DOCX batch without an author", async () => {
    const bytes = buildDocx({ body: paragraph("Hello") + sectPr() });
    const { session: edit, end } = await docxSession(bytes, [[]]);
    try {
      const outline = await edit.getOutline();
      const target = outline.items[0]!.id;
      const refused = await edit.callTool(
        {
          name: "document_apply",
          arguments: {
            operations: [{ op: "replaceText", target, text: "Bye" }],
          },
        },
        { changeMode: "tracked" },
      );
      assert.equal(refused.ok, false);
      assert.deepEqual(issueCodes(refused), ["required@/author"]);
      assert.match(refused.text, /ApplyOptions\.author/);
      assert.equal(edit.state.revision, 0);
    } finally {
      await end();
    }
  });
});
