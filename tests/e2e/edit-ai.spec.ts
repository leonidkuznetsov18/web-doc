import { expect, test, type Page } from "@playwright/test";

import {
  buildDocx,
  paragraph,
  sectPr,
  syntheticDocument,
} from "../../packages/viewer/test/fixtures/docx-builder.js";
import { buildPdf } from "../../packages/viewer/test/fixtures/pdf-builder.js";
import {
  buildDeck,
  syntheticDeck,
  textShape,
} from "../../packages/viewer/test/fixtures/pptx-builder.js";

/*
 * The AI tooling of module 07 against the real adapters: a scripted agent
 * turn over each format through `callTool()` only (describe, find,
 * preview, apply, checkpoint, restore), a tracked DOCX edit that survives
 * save and reload with its accepted text, and the latency of `describe()`
 * on 500 pages and 500 slides.
 */

interface ToolResult {
  ok: boolean;
  content: any;
  text: string;
  issues?: { operationIndex: number; path: string; code: string }[];
}

async function loadDocument(
  page: Page,
  bytes: Uint8Array,
  fileName: string,
  limits: Record<string, number> = {},
): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ({ data, fileName, limits }) => {
      const { ViewerClient } = (await import("/main.js")) as {
        ViewerClient: { create(config: unknown): { createViewer(): unknown } };
      };
      const client = ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
        fontPolicy: { mode: "offline" },
        limits,
      });
      const viewer = client.createViewer() as {
        load(bytes: Uint8Array, options: unknown): Promise<void>;
      };
      await viewer.load(new Uint8Array(data), { fileName });
      (window as unknown as { __viewer: unknown }).__viewer = viewer;
    },
    { data: Array.from(bytes), fileName, limits },
  );
}

/**
 * One agent turn through the tools: checkpoint, describe, find `quote`,
 * preview and apply a replacement of it by `replacement`, read the text
 * back, then restore the checkpoint. Every step goes through `callTool`.
 */
async function agentTurn(
  page: Page,
  quote: string,
  replacement: string,
  options: { changeMode?: "direct" | "tracked"; author?: string } = {},
) {
  return page.evaluate(
    async ({ quote, replacement, options }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const call = (
        name: string,
        args: unknown,
        extra: Record<string, unknown> = {},
      ): Promise<ToolResult> =>
        session.callTool({ name, arguments: args }, { ...options, ...extra });
      const tools = session.tools.definitions.map((tool: any) => tool.name);
      const checkpoint = await call("document_checkpoint", {
        action: "create",
        label: "turn",
      });
      const described = await call("document_describe", {});
      const revision = described.content.revision as number;
      const found = await call("document_find", { text: quote });
      const best = found.content.candidates[0];
      const operations = [
        {
          op: "replaceText",
          target: best.elementId,
          ...(best.range ? { range: best.range } : {}),
          text: replacement,
        },
      ];
      const preview = await call("document_preview", { operations });
      const stale = await call(
        "document_apply",
        { operations },
        { expectedRevision: revision + 7 },
      );
      const applied = await call(
        "document_apply",
        { operations, label: "agent" },
        { expectedRevision: revision },
      );
      const after = await call("document_describe", {});
      const inspected = await call("document_inspect", { id: best.elementId });
      const pageText = await viewer.getPageText(Math.max(0, best.pageIndex));
      const restored = await call("document_checkpoint", {
        action: "restore",
        id: checkpoint.content.id,
      });
      const back = await call("document_find", { text: quote });
      return {
        tools,
        format: session.format,
        checkpoint: checkpoint.text,
        description: described.content.text as string,
        describedText: described.text,
        best,
        foundText: found.text,
        preview,
        stale,
        applied,
        afterText: after.content.text as string,
        inspected,
        pageText,
        restored,
        backCount: back.content.candidates.length as number,
        revision: session.state.revision as number,
      };
    },
    { quote, replacement, options },
  );
}

const TOOLS = [
  "document_describe",
  "document_outline",
  "document_find",
  "document_inspect",
  "document_preview",
  "document_apply",
  "document_undo",
  "document_checkpoint",
];

test("runs an agent turn on a PDF through the tools only", async ({ page }) => {
  await loadDocument(
    page,
    await buildPdf(["Quarterly review of the northern region", "Appendix"]),
    "turn.pdf",
  );
  const turn = await agentTurn(page, "northern region", "southern region");
  expect(turn.tools).toEqual(TOOLS);
  expect(turn.format).toBe("pdf");
  expect(turn.checkpoint).toMatch(
    /^Checkpoint [A-Za-z0-9_-]{22} created at revision 0 \("turn"\)\.$/,
  );
  expect(turn.description.split("\n")[0]).toBe("pdf: 2 pages, 2 elements");
  expect(turn.description).toContain(
    "page 1 text: Quarterly review of the northern region",
  );
  expect(turn.best.reason).toBe("exact");
  expect(turn.best.score).toBe(1);
  expect(turn.best.range).toBeTruthy();
  expect(turn.preview.ok).toBe(true);
  expect(turn.preview.content.dryRun).toBe(true);
  expect(turn.preview.text).toMatch(
    /^Preview, nothing changed: Replaced a range of the text of /,
  );
  expect(turn.stale.ok).toBe(false);
  expect(turn.stale.issues![0]!.code).toBe("edit-conflict");
  expect(turn.applied.ok).toBe(true);
  expect(turn.applied.content.revision).toBe(1);
  expect(turn.afterText).toContain("southern region");
  expect(turn.pageText).toContain("southern region");
  expect(turn.inspected.ok).toBe(true);
  expect(turn.inspected.content.text).toContain("southern region");
  expect(turn.inspected.content.bounds).toBeUndefined();
  expect(turn.restored.ok).toBe(true);
  expect(turn.restored.text).toMatch(/^Restored checkpoint .*; revision 2\.$/);
  expect(turn.backCount).toBe(1);
  expect(turn.revision).toBe(2);
});

test("runs an agent turn on a deck through the tools only", async ({
  page,
}) => {
  const deck = buildDeck({
    slides: [
      {
        shapes: [
          textShape({
            id: 2,
            name: "Title 1",
            x: 914400,
            y: 457200,
            cx: 6400800,
            cy: 914400,
            paragraphs: [["Agenda for the quarter"]],
          }),
        ],
      },
      {
        shapes: [
          textShape({
            id: 2,
            name: "Title 1",
            x: 914400,
            y: 457200,
            cx: 6400800,
            cy: 914400,
            paragraphs: [["Results by region"]],
          }),
        ],
      },
    ],
  });
  await loadDocument(page, deck, "turn.pptx");
  const turn = await agentTurn(page, "Results by region", "Results by country");
  expect(turn.format).toBe("pptx");
  expect(turn.description.split("\n")[0]).toBe("pptx: 2 slides, 2 elements");
  expect(turn.description).toContain(
    '[sld2:2] slide 2 shape "Title 1": Results by region',
  );
  expect(turn.best.elementId).toBe("sld2:2");
  expect(turn.best.range).toBeUndefined();
  expect(turn.applied.ok).toBe(true);
  expect(turn.applied.text).toBe(
    "Replaced the text of sld2:2. Revision 1; changed slide 2; 2 slides.",
  );
  expect(turn.afterText).toContain("Results by country");
  expect(turn.pageText).toContain("Results by country");
  expect(turn.restored.ok).toBe(true);
  expect(turn.backCount).toBe(1);
});

test("runs an agent turn on a Word document and keeps a tracked edit's accepted text through save and reload", async ({
  page,
}) => {
  const bytes = buildDocx({
    body:
      paragraph("First paragraph of the report") +
      paragraph("The review covers three regions") +
      paragraph("Closing remarks") +
      sectPr(),
  });
  await loadDocument(page, bytes, "turn.docx");
  const turn = await agentTurn(page, "three regions", "four regions");
  expect(turn.format).toBe("docx");
  expect(turn.description.split("\n")[0]).toBe("docx: 1 page, 3 elements");
  // Before the viewer lays a page out, DOCX nodes carry no page.
  expect(turn.description).toMatch(
    /\[p:[0-9A-F]{8}\] paragraph: The review covers three regions/,
  );
  expect(turn.best.reason).toBe("exact");
  expect(turn.applied.ok).toBe(true);
  expect(turn.afterText).toContain("four regions");
  expect(turn.pageText).toContain("four regions");
  expect(turn.restored.ok).toBe(true);
  expect(turn.backCount).toBe(1);

  // The same edit as a suggestion: refused without an author, then
  // written as revisions the reads and the renderer show accepted.
  const tracked = await page.evaluate(async () => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    const session = await viewer.edit();
    const found = await session.callTool({
      name: "document_find",
      arguments: { text: "three regions" },
    });
    const best = found.content.candidates[0];
    const operations = [
      {
        op: "replaceText",
        target: best.elementId,
        range: best.range,
        text: "four regions",
      },
    ];
    const anonymous = await session.callTool(
      { name: "document_apply", arguments: { operations } },
      { changeMode: "tracked" },
    );
    const applied = await session.callTool(
      { name: "document_apply", arguments: { operations } },
      { changeMode: "tracked", author: "Writer agent" },
    );
    const revisions = (await session.getRevisions(best.elementId)).items;
    const element = (await session.getElement(best.elementId)).item;
    const outline = await session.describe();
    const pageText = await viewer.getPageText(0);
    const saved = await session.save();
    const client = (await import("/main.js")) as any;
    const fresh = client.ViewerClient.create({
      assetBaseUrl: new URL("/", location.href),
      fontPolicy: { mode: "offline" },
    }).createViewer();
    await fresh.load(saved.bytes, { fileName: "suggested.docx" });
    const reloaded = await fresh.getPageText(0);
    const reopened = await fresh.edit();
    const again = (await reopened.getElement(best.elementId)).item;
    const locked = await reopened.callTool(
      {
        name: "document_apply",
        arguments: {
          operations: [
            { op: "replaceText", target: best.elementId, text: "x" },
          ],
        },
      },
      {},
    );
    return {
      anonymous,
      applied,
      revisions: revisions.map((r: any) => [r.kind, r.author, r.text]),
      element,
      descriptionLine: (outline.item.text as string).split("\n")[2],
      pageText,
      reloaded,
      again,
      locked,
    };
  });
  expect(tracked.anonymous.ok).toBe(false);
  expect(tracked.anonymous.issues![0]).toMatchObject({
    code: "required",
    path: "/author",
  });
  expect(tracked.applied.ok).toBe(true);
  expect(tracked.revisions).toEqual([
    ["del", "Writer agent", "three regions"],
    ["ins", "Writer agent", "four regions"],
  ]);
  expect(tracked.element.text).toBe("The review covers four regions");
  expect(tracked.element.readOnlyReason).toBe("tracked-changes");
  expect(tracked.descriptionLine).toMatch(
    /paragraph \(read-only: tracked-changes\): The review covers four regions$/,
  );
  expect(tracked.pageText).toContain("four regions");
  expect(tracked.pageText).not.toContain("three regions");
  expect(tracked.reloaded).toContain("four regions");
  expect(tracked.reloaded).not.toContain("three regions");
  expect(tracked.again.text).toBe("The review covers four regions");
  expect(tracked.again.readOnlyReason).toBe("tracked-changes");
  expect(tracked.locked.ok).toBe(false);
  expect(tracked.locked.issues![0]!.code).toBe("invalid-target");
});

test(
  "describes 500 pages and 500 slides within the operation budget and records the latency",
  { tag: "@performance" },
  async ({ page }) => {
    test.setTimeout(300_000);
    const timings: Record<string, number> = {};
    const documents: [string, Uint8Array][] = [
      ["docx", syntheticDocument(500)],
      ["pptx", syntheticDeck(500)],
      [
        "pdf",
        await buildPdf(Array.from({ length: 500 }, (_, i) => `Page ${i + 1}`)),
      ],
    ];
    for (const [format, bytes] of documents) {
      // Opening 500 pages for editing is the engine's work, not what this
      // measures: a loaded CI runner may need more than the default budget
      // for it, while `describe()` itself is still held to that budget.
      await loadDocument(page, bytes, `large.${format}`, {
        maxOperationMs: 180_000,
      });
      const result = await page.evaluate(async () => {
        const viewer = (window as unknown as { __viewer: any }).__viewer;
        const session = await viewer.edit();
        const started = performance.now();
        const description = await session.describe();
        const elapsed = performance.now() - started;
        const again = performance.now();
        await session.describe();
        const warm = performance.now() - again;
        return {
          elapsed,
          warm,
          pageCount: description.item.pageCount,
          elementCount: description.item.elementCount,
          chars: description.item.text.length,
          truncated: description.item.truncated,
          head: description.item.text.split("\n").slice(0, 2),
        };
      });
      timings[format] = result.elapsed;
      console.log(
        `describe ${format} 500: ${result.elapsed.toFixed(0)} ms cold, ${result.warm.toFixed(0)} ms warm; ${result.elementCount} elements, ${result.chars} chars${result.truncated ? " (truncated)" : ""}`,
      );
      expect(result.pageCount).toBe(500);
      expect(result.elementCount).toBeGreaterThanOrEqual(500);
      // The 500-page PDF fits the default budget; the deck and the Word
      // document do not.
      expect(result.truncated).toBe(result.chars >= 49_000);
      expect(result.head[0]).toMatch(
        new RegExp(`^${format}: 500 (pages|slides), \\d+ elements$`),
      );
      expect(result.elapsed).toBeLessThan(30_000);
    }
  },
);
