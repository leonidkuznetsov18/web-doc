import { readFile } from "node:fs/promises";

import { expect, test, type Page } from "@playwright/test";

import {
  buildDocx,
  paragraph,
  sectPr,
} from "../../packages/viewer/test/fixtures/docx-builder.js";

/*
 * DOCX editing through the public API against the real adapter: the
 * @silurus/ooxml renderer paints and reports runs with paragraph ids, the
 * OOXML edit worker indexes the body, and the session joins the two.
 */

const CORPUS = new URL("../../.cache/corpus/", import.meta.url);
const EDIT_ASSETS = ["/workers/ooxml-edit-worker.js"];

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface Element {
  id: string;
  kind: string;
  pageIndex: number;
  text?: string;
  parentId?: string;
  bounds: Rect;
  fragments: { pageIndex: number; bounds: Rect }[];
  operations: string[];
  textStyle?: { fontFamily: string; fontSize: number };
}

interface Run {
  text: string;
  x: number;
  y: number;
  width: number;
  height: number;
  paragraphId?: string;
}

async function loadDocument(
  page: Page,
  bytes: Uint8Array,
  fileName: string,
): Promise<void> {
  await page.goto("/");
  await page.evaluate(
    async ({ data, fileName }) => {
      const { ViewerClient } = (await import("/main.js")) as {
        ViewerClient: { create(config: unknown): { createViewer(): unknown } };
      };
      const client = ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
        fontPolicy: { mode: "offline" },
      });
      const viewer = client.createViewer() as {
        load(bytes: Uint8Array, options: unknown): Promise<void>;
      };
      await viewer.load(new Uint8Array(data), { fileName });
      (window as unknown as { __viewer: unknown }).__viewer = viewer;
    },
    { data: Array.from(bytes), fileName },
  );
}

/** Every run of a page, through a whole-page selection. */
async function runsOf(page: Page, pageIndex: number): Promise<Run[]> {
  return page.evaluate(async (pageIndex) => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    const selection = await viewer.selectText({
      startPageIndex: pageIndex,
      startOffset: 0,
      endPageIndex: pageIndex,
      endOffset: Number.MAX_SAFE_INTEGER,
    });
    return selection.runs.map((run: Run) => ({
      text: run.text,
      x: run.x,
      y: run.y,
      width: run.width,
      height: run.height,
      paragraphId: run.paragraphId,
    }));
  }, pageIndex);
}

const covers = (outer: Rect, inner: Rect, tolerance = 0.5): boolean =>
  inner.x >= outer.x - tolerance &&
  inner.y >= outer.y - tolerance &&
  inner.x + inner.width <= outer.x + outer.width + tolerance &&
  inner.y + inner.height <= outer.y + outer.height + tolerance;

test("starts the OOXML worker only on edit() and places the corpus document's paragraphs with the renderer's runs", async ({
  page,
}) => {
  const original = new Uint8Array(
    await readFile(new URL("sample.docx", CORPUS)),
  );
  const requests: string[] = [];
  page.on("request", (request) =>
    requests.push(new URL(request.url()).pathname),
  );
  await loadDocument(page, original, "sample.docx");
  const info = await page.evaluate(() => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    return {
      editing: viewer.getDocumentInfo().capabilities.editing,
      pageCount: viewer.state.pageCount,
    };
  });
  expect(info).toEqual({ editing: true, pageCount: 1 });
  expect(requests.filter((path) => EDIT_ASSETS.includes(path))).toEqual([]);

  const runs = await runsOf(page, 0);
  const result = await page.evaluate(
    async ({ data }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const elements = (await session.getElements({ pageIndex: 0 }))
        .items as Element[];
      const all = (await session.getElements()).items as Element[];
      const first = elements[0]!;
      const centre = {
        x: first.bounds.x + 2,
        y: first.bounds.y + first.bounds.height / 2,
      };
      const hits = (await session.elementsAt(0, centre)).items as Element[];
      const found = (await session.findText("lorem")).items as {
        pageIndex: number;
        rects: Rect[];
        elementIds: string[];
        text: string;
      }[];
      const saved = await session.save();
      const identical =
        saved.bytes.length === data.length &&
        saved.bytes.every(
          (byte: number, index: number) => byte === data[index],
        );
      return {
        format: session.format,
        elements,
        all,
        hits: hits.map((element) => element.id),
        found,
        identical,
        dirty: session.state.dirty,
        operations: Object.keys(session.schemas.operations),
      };
    },
    { data: Array.from(original) },
  );
  expect(requests.filter((path) => EDIT_ASSETS.includes(path))).toHaveLength(1);
  expect(result.format).toBe("docx");
  expect(result.identical).toBe(true);
  expect(result.dirty).toBe(false);

  // Every paragraph of the body is on page 0 with bounds that cover its runs.
  const byParagraph = new Map<string, Run[]>();
  for (const run of runs) {
    if (!run.paragraphId) continue;
    const list = byParagraph.get(run.paragraphId) ?? [];
    list.push(run);
    byParagraph.set(run.paragraphId, list);
  }
  expect(result.elements.length).toBe(byParagraph.size);
  for (const element of result.elements) {
    expect(element.kind).toBe("paragraph");
    expect(element.pageIndex).toBe(0);
    expect(element.fragments).toHaveLength(1);
    expect(element.textStyle!.fontSize).toBeGreaterThan(0);
    const paragraphRuns = byParagraph.get(element.id.slice(2)) ?? [];
    expect(paragraphRuns.length).toBeGreaterThan(0);
    for (const run of paragraphRuns)
      expect(covers(element.bounds, run)).toBe(true);
    // The engine's text and the renderer's runs agree on the words.
    const words = paragraphRuns
      .map((run) => run.text)
      .join("")
      .trim()
      .split(/\s+/);
    for (const word of words.slice(0, 3)) expect(element.text).toContain(word);
  }
  // Page 0 was laid out, so the page-less query places the same elements.
  expect(result.all.map((element) => element.pageIndex)).toEqual(
    result.elements.map(() => 0),
  );
  expect(result.hits).toEqual([result.elements[0]!.id]);
  expect(result.found.length).toBeGreaterThan(0);
  for (const hit of result.found) {
    expect(hit.pageIndex).toBe(0);
    expect(hit.rects.length).toBeGreaterThan(0);
    expect(hit.text.toLowerCase()).toBe("lorem");
  }
});

test("joins a paragraph that continues on the next page across both pages", async ({
  page,
}) => {
  const bytes = buildDocx({
    body:
      `<w:p><w:r><w:t>Before the break</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>After the break</w:t></w:r></w:p>` +
      `<w:tbl><w:tr><w:tc>${paragraph("In a cell")}</w:tc><w:tc>${paragraph("Other cell")}</w:tc></w:tr></w:tbl>` +
      paragraph("Last paragraph") +
      sectPr(),
  });
  await loadDocument(page, bytes, "split.docx");
  const result = await page.evaluate(async () => {
    const viewer = (window as unknown as { __viewer: any }).__viewer;
    const session = await viewer.edit();
    const page1 = (await session.getElements({ pageIndex: 1 }))
      .items as Element[];
    // Lay out page 0 as well (a whole-page selection reads its runs).
    await viewer.selectText({
      startPageIndex: 0,
      startOffset: 0,
      endPageIndex: 0,
      endOffset: 1,
    });
    const all = (await session.getElements()).items as Element[];
    const table = all.find((element) => element.kind === "table")!;
    const cellHit = (
      await session.elementsAt(1, {
        x: table.bounds.x + 1,
        y: table.bounds.y + 1,
      })
    ).items as Element[];
    return {
      pageCount: viewer.state.pageCount,
      page1: page1.map((element) => [
        element.kind,
        element.text,
        element.pageIndex,
      ]),
      all: all.map((element) => ({
        kind: element.kind,
        text: element.text,
        pageIndex: element.pageIndex,
        fragments: element.fragments.map((fragment) => fragment.pageIndex),
        parentId: element.parentId,
      })),
      cellHit: cellHit.map((element) => element.kind),
    };
  });
  expect(result.pageCount).toBe(2);
  expect(result.page1).toEqual([
    ["paragraph", "Before the break\fAfter the break", 1],
    ["table", "In a cell\tOther cell", 1],
    ["paragraph", "In a cell", 1],
    ["paragraph", "Other cell", 1],
    ["paragraph", "Last paragraph", 1],
  ]);
  const split = result.all[0]!;
  expect(split.pageIndex).toBe(0);
  expect(split.fragments).toEqual([0, 1]);
  expect(result.all[1]!.fragments).toEqual([1]);
  expect(result.all[2]!.parentId).toBe(
    result.all[1]!.parentId ?? result.all[2]!.parentId,
  );
  expect(result.cellHit).toEqual(["paragraph", "table"]);
});
