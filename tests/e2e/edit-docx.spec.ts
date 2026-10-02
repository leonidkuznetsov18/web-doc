import { readFile } from "node:fs/promises";

import { expect, test, type Page } from "@playwright/test";

import {
  localRecordOf,
  parseZip,
} from "../../packages/viewer/src/edit/ooxml/zip.js";
import { defaultResourceLimits } from "../../packages/viewer/src/limits.js";
import {
  buildDocx,
  paragraph,
  sectPr,
  syntheticDocument,
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

test("replaces and restyles paragraph text so the renderer shows it, repaints from its page and survives save and reload", async ({
  page,
}) => {
  const bytes = buildDocx({
    body:
      paragraph("First paragraph of the document") +
      paragraph("Second paragraph to edit") +
      paragraph("Third paragraph stays") +
      sectPr(),
  });
  await loadDocument(page, bytes, "edit.docx");
  const result = await page.evaluate(
    async ({ data }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const changes: number[][] = [];
      viewer.on("documentchange", (event: { changedPages: number[] }) =>
        changes.push([...event.changedPages]),
      );
      const session = await viewer.edit();
      const [first, second, third] = (
        await session.getElements({ pageIndex: 0 })
      ).items as Element[];
      const before = await viewer.getPageText(0);
      const replaced = await session.replaceText({
        target: second!.id,
        text: "Rewritten\tparagraph",
      });
      const afterReplace = await viewer.getPageText(0);
      const styled = await session.setTextStyle({
        target: second!.id,
        style: { bold: true, fontSize: 20, color: "#FF0000" },
      });
      await session.setParagraphStyle({
        target: second!.id,
        style: { align: "center" },
      });
      const element = (await session.getElement(second!.id)).item as Element & {
        textStyle: { bold: boolean; fontSize: number; color: string };
        paragraphStyle: { align: string };
      };
      const split = await session.replaceText({
        target: first!.id,
        text: "Alpha\nBeta",
      });
      const texts = (await session.getElements({ pageIndex: 0 })).items.map(
        (item: Element) => item.text,
      );
      const saved = await session.save();
      const client = (await import("/main.js")) as any;
      const fresh = client.ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
        fontPolicy: { mode: "offline" },
      }).createViewer();
      await fresh.load(saved.bytes, { fileName: "edited.docx" });
      const reloaded = await fresh.getPageText(0);
      await fresh.destroy();
      await session.undo();
      await session.undo();
      await session.undo();
      await session.undo();
      const restored = await session.save();
      const identical =
        restored.bytes.length === data.length &&
        restored.bytes.every(
          (byte: number, index: number) => byte === data[index],
        );
      return {
        before,
        afterReplace,
        changedPages: [
          replaced.changedPages,
          styled.changedPages,
          split.changedPages,
        ],
        events: changes,
        element,
        texts,
        reloaded,
        identical,
        third: third!.id,
        thirdAfter: (
          (await session.getElements({ pageIndex: 0 })).items as Element[]
        ).at(-1)!.id,
      };
    },
    { data: Array.from(bytes) },
  );
  expect(result.before).toContain("Second paragraph to edit");
  expect(result.afterReplace).not.toContain("Second paragraph to edit");
  expect(result.afterReplace).toContain("Rewritten");
  expect(result.afterReplace).toContain("paragraph");
  expect(result.changedPages).toEqual([[0], [0], [0]]);
  expect(result.events.slice(0, 3)).toEqual([[0], [0], [0]]);
  expect(result.element.textStyle).toMatchObject({
    bold: true,
    fontSize: 20,
    color: "#FF0000",
  });
  expect(result.element.paragraphStyle.align).toBe("center");
  expect(result.texts).toEqual([
    "Alpha",
    "Beta",
    "Rewritten\tparagraph",
    "Third paragraph stays",
  ]);
  expect(result.reloaded).toContain("Alpha");
  expect(result.reloaded).toContain("Beta");
  expect(result.reloaded).toContain("Rewritten");
  expect(result.identical).toBe(true);
  // The untouched paragraph kept its id through every edit and undo.
  expect(result.thirdAfter).toBe(result.third);
});

test("inserts, moves and deletes paragraphs and pictures that the renderer draws in the new order", async ({
  page,
}) => {
  const bytes = buildDocx({
    body:
      paragraph("Alpha paragraph") +
      paragraph("Beta paragraph") +
      paragraph("Gamma paragraph") +
      sectPr(),
  });
  // A 1×1 PNG, so the renderer has real pixels to decode.
  const png = Uint8Array.from(
    atob(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
    ),
    (c) => c.charCodeAt(0),
  );
  await loadDocument(page, bytes, "structure.docx");
  const result = await page.evaluate(
    async ({ data, png }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const [alpha, beta, gamma] = (await session.getElements({ pageIndex: 0 }))
        .items as Element[];
      const inserted = await session.insertParagraph({
        after: alpha!.id,
        text: "Inserted after alpha",
        style: { italic: true },
      });
      const moved = await session.moveElement({
        target: gamma!.id,
        before: alpha!.id,
      });
      const picture = await session.insertImage({
        after: beta!.id,
        data: new Uint8Array(png),
        mimeType: "image/png",
        size: { width: 48, height: 48 },
      });
      const afterInserts = await viewer.getPageText(0);
      const placed = (await session.getElements({ pageIndex: 0 }))
        .items as Element[];
      const deleted = await session.deleteElement({ target: beta!.id });
      const afterDelete = await viewer.getPageText(0);
      const saved = await session.save();
      const client = (await import("/main.js")) as any;
      const fresh = client.ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
        fontPolicy: { mode: "offline" },
      }).createViewer();
      await fresh.load(saved.bytes, { fileName: "structured.docx" });
      const reloaded = await fresh.getPageText(0);
      await fresh.destroy();
      await session.reset();
      const restored = await session.save();
      const identical =
        restored.bytes.length === data.length &&
        restored.bytes.every(
          (byte: number, index: number) => byte === data[index],
        );
      return {
        inserted: inserted.createdIds,
        movedPages: moved.changedPages,
        picture: picture.createdIds,
        afterInserts,
        order: placed.map((element) => [element.kind, element.text]),
        imagePlaced: placed.find((element) => element.kind === "image")!,
        deleted: deleted.removedIds,
        afterDelete,
        reloaded,
        identical,
      };
    },
    { data: Array.from(bytes), png: Array.from(png) },
  );
  expect(result.inserted).toHaveLength(1);
  expect(result.movedPages).toEqual([0]);
  expect(result.picture).toHaveLength(2);
  const alphaAt = result.afterInserts.indexOf("Alpha");
  const gammaAt = result.afterInserts.indexOf("Gamma");
  const insertedAt = result.afterInserts.indexOf("Inserted after alpha");
  expect(gammaAt).toBeGreaterThanOrEqual(0);
  expect(gammaAt).toBeLessThan(alphaAt);
  expect(insertedAt).toBeGreaterThan(alphaAt);
  expect(result.order).toEqual([
    ["paragraph", "Gamma paragraph"],
    ["paragraph", "Alpha paragraph"],
    ["paragraph", "Inserted after alpha"],
    ["paragraph", "Beta paragraph"],
    ["paragraph", "\ufffc"],
    ["image", undefined],
  ]);
  // The picture's paragraph is placed from the runs; the picture takes it.
  expect(result.imagePlaced.pageIndex).toBe(0);
  expect(result.deleted).toEqual([expect.stringMatching(/^p:/)]);
  expect(result.afterDelete).not.toContain("Beta paragraph");
  expect(result.reloaded).toContain("Inserted after alpha");
  expect(result.reloaded).not.toContain("Beta paragraph");
  expect(result.identical).toBe(true);
});

test("inserts a table the renderer draws and edits a cell that the page text shows", async ({
  page,
}) => {
  const bytes = buildDocx({
    body:
      paragraph("Before the table") + paragraph("After the table") + sectPr(),
  });
  await loadDocument(page, bytes, "table.docx");
  const result = await page.evaluate(
    async ({ data }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const [before] = (await session.getElements({ pageIndex: 0 }))
        .items as Element[];
      const inserted = await session.insertTable({
        after: before!.id,
        rows: [
          ["Name", "Value"],
          ["alpha", "one"],
        ],
        columnWidths: [1, 2],
      });
      const table = inserted.createdIds[0]!;
      const afterInsert = await viewer.getPageText(0);
      const edited = await session.setTableCell({
        target: table,
        row: 1,
        column: 1,
        text: "uno",
      });
      const afterEdit = await viewer.getPageText(0);
      const placed = (await session.getElements({ pageIndex: 0 }))
        .items as (Element & {
        table?: { rows: string[][] };
      })[];
      const tableElement = placed.find((element) => element.id === table)!;
      const hit = (
        await session.elementsAt(0, {
          x: tableElement.bounds.x + 2,
          y: tableElement.bounds.y + 2,
        })
      ).items as Element[];
      const saved = await session.save();
      const client = (await import("/main.js")) as any;
      const fresh = client.ViewerClient.create({
        assetBaseUrl: new URL("/", location.href),
        fontPolicy: { mode: "offline" },
      }).createViewer();
      await fresh.load(saved.bytes, { fileName: "tabled.docx" });
      const reloaded = await fresh.getPageText(0);
      await fresh.destroy();
      await session.undo();
      await session.undo();
      const restored = await session.save();
      const identical =
        restored.bytes.length === data.length &&
        restored.bytes.every(
          (byte: number, index: number) => byte === data[index],
        );
      return {
        created: inserted.createdIds.length,
        insertPages: inserted.changedPages,
        editPages: edited.changedPages,
        afterInsert,
        afterEdit,
        tableRows: tableElement.table,
        tableBounds: tableElement.bounds,
        hit: hit.map((element) => element.kind),
        reloaded,
        identical,
      };
    },
    { data: Array.from(bytes) },
  );
  expect(result.created).toBe(5);
  expect(result.insertPages).toEqual([0]);
  expect(result.editPages).toEqual([0]);
  expect(result.afterInsert).toContain("Name");
  expect(result.afterInsert).toContain("alpha");
  expect(result.afterInsert).toContain("one");
  expect(result.afterEdit).toContain("uno");
  expect(result.afterEdit).not.toContain("one");
  expect(result.tableRows).toEqual({
    rows: [
      ["Name", "Value"],
      ["alpha", "uno"],
    ],
  });
  expect(result.tableBounds.width).toBeGreaterThan(0);
  expect(result.hit).toEqual(["paragraph", "table"]);
  expect(result.reloaded).toContain("uno");
  expect(result.identical).toBe(true);
});

/** Names of the ZIP entries whose local records differ between two packages. */
function changedEntries(a: Uint8Array, b: Uint8Array): string[] {
  const left = parseZip(a, defaultResourceLimits);
  const right = parseZip(b, defaultResourceLimits);
  const changed: string[] = [];
  for (const entry of right.entries) {
    const before = left.entries.find(
      (candidate) => candidate.name === entry.name,
    );
    if (!before) {
      changed.push(entry.name);
      continue;
    }
    const x = localRecordOf(left, before);
    const y = localRecordOf(right, entry);
    const bytesBefore = a.subarray(x.headerOffset, x.recordEnd);
    const bytesAfter = b.subarray(y.headerOffset, y.recordEnd);
    if (
      bytesBefore.length !== bytesAfter.length ||
      bytesBefore.some((byte, index) => byte !== bytesAfter[index])
    )
      changed.push(entry.name);
  }
  for (const entry of left.entries)
    if (!right.entries.some((candidate) => candidate.name === entry.name))
      changed.push(`-${entry.name}`);
  return changed.sort();
}

test("changes only the entries an edit touches: the body part, plus media and relationships for a picture", async ({
  page,
}) => {
  const bytes = buildDocx({
    body: paragraph("One") + paragraph("Two") + paragraph("Three") + sectPr(),
  });
  const png = Uint8Array.from(
    atob(
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==",
    ),
    (c) => c.charCodeAt(0),
  );
  await loadDocument(page, bytes, "entries.docx");
  const saved = await page.evaluate(
    async ({ png }) => {
      const viewer = (window as unknown as { __viewer: any }).__viewer;
      const session = await viewer.edit();
      const [one, two] = (await session.getElements({ pageIndex: 0 }))
        .items as Element[];
      await session.replaceText({ target: two!.id, text: "Two, edited" });
      await session.setTextStyle({ target: one!.id, style: { bold: true } });
      await session.setParagraphStyle({
        target: one!.id,
        style: { align: "right" },
      });
      await session.insertTable({ after: one!.id, rows: [["a", "b"]] });
      const text = Array.from((await session.save()).bytes as Uint8Array);
      await session.insertImage({
        after: two!.id,
        data: new Uint8Array(png),
        mimeType: "image/png",
        size: { width: 10, height: 10 },
      });
      const picture = Array.from((await session.save()).bytes as Uint8Array);
      return { text, picture };
    },
    { png: Array.from(png) },
  );
  expect(changedEntries(bytes, new Uint8Array(saved.text))).toEqual([
    "word/document.xml",
  ]);
  expect(changedEntries(bytes, new Uint8Array(saved.picture))).toEqual([
    "[Content_Types].xml",
    "word/_rels/document.xml.rels",
    "word/document.xml",
    "word/media/image1.png",
  ]);
});

test(
  "applies an edit within the budget on 10- and 100-page documents and records 500",
  { tag: "@performance" },
  async ({ page }) => {
    test.setTimeout(300_000);
    const timings: Record<string, number[]> = {};
    for (const count of [10, 100, 500]) {
      await loadDocument(page, syntheticDocument(count), `pages-${count}.docx`);
      timings[count] = await page.evaluate(async () => {
        const viewer = (window as unknown as { __viewer: any }).__viewer;
        const session = await viewer.edit();
        const elements = (await session.getElements()).items as Element[];
        const first = elements[0]!;
        const last = elements.at(-1)!;
        const out: number[] = [];
        for (const text of ["First edit", "Second edit"]) {
          const started = performance.now();
          await session.replaceText({ target: first.id, text });
          out.push(performance.now() - started);
        }
        const started = performance.now();
        await session.insertParagraph({ after: last.id, text: "Last page" });
        out.push(performance.now() - started);
        return out;
      });
      console.log(
        `docx apply ${count} pages: replaceText ${timings[count]![0]!.toFixed(0)} ms then ${timings[count]![1]!.toFixed(0)} ms, insertParagraph on the last page ${timings[count]![2]!.toFixed(0)} ms`,
      );
      if (count < 500)
        for (const value of timings[count]!) expect(value).toBeLessThan(3000);
    }
  },
);
