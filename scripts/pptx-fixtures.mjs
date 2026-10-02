import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

/*
 * Writes one edited copy of the corpus deck per PPTX operation into
 * artifacts/pptx-fixtures/, for the manual check that PowerPoint and Keynote
 * open every result without a repair prompt. Needs the viewer built
 * (npm run build --workspace web-doc) and the corpus fetched.
 */

const root = resolve(import.meta.dirname, "..");
const { PptxEditEngine } =
  await import("../packages/viewer/dist/edit/pptx/engine.js");
const { defaultResourceLimits } =
  await import("../packages/viewer/dist/limits.js");

const source = resolve(root, ".cache/corpus/sample.pptx");
const outdir = resolve(root, "artifacts/pptx-fixtures");
const original = new Uint8Array(await readFile(source));
const png = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFElEQVR4nGP4z8DwHwyBFJAAgAAA//8R7AP8Ky7YKAAAAABJRU5ErkJggg==",
  "base64",
);

const cases = [
  [
    "replaceText",
    [
      {
        op: "replaceText",
        target: "sld1:2",
        text: "Replaced title\nSecond line",
      },
    ],
  ],
  [
    "replaceText-range",
    [
      {
        op: "replaceText",
        target: "sld1:3",
        text: "patched",
        range: {
          start: { elementId: "sld1:3", offset: 0 },
          end: { elementId: "sld1:3", offset: 8 },
        },
      },
    ],
  ],
  [
    "setTextStyle",
    [
      {
        op: "setTextStyle",
        target: "sld2:3",
        style: {
          bold: true,
          italic: true,
          underline: true,
          fontSize: 20,
          fontFamily: "Arial",
          color: { theme: "accent2" },
          align: "right",
        },
      },
    ],
  ],
  [
    "setShapeStyle",
    [
      {
        op: "setShapeStyle",
        target: "sld1:2",
        fill: "#FFE699",
        line: { color: { theme: "accent1" }, width: 2 },
      },
    ],
  ],
  [
    "moveElement",
    [{ op: "moveElement", target: "sld1:3", by: { dx: 40, dy: -120 } }],
  ],
  [
    "resizeElement",
    [
      {
        op: "resizeElement",
        target: "sld1:2",
        rect: { x: 60, y: 40, width: 600, height: 120 },
      },
    ],
  ],
  ["deleteElement", [{ op: "deleteElement", target: "sld2:3" }]],
  [
    "insertTextBox",
    [
      {
        op: "insertTextBox",
        pageIndex: 1,
        rect: { x: 80, y: 420, width: 500, height: 80 },
        text: "Inserted text box\u000bwith a line break",
        style: { fontSize: 24, color: "#C00000" },
      },
    ],
  ],
  [
    "insertImage",
    [
      {
        op: "insertImage",
        pageIndex: 0,
        rect: { x: 600, y: 60, width: 120, height: 120 },
        data: new Uint8Array(png),
        mimeType: "image/png",
      },
    ],
  ],
  [
    "insertTable",
    [
      {
        op: "insertTable",
        pageIndex: 1,
        rect: { x: 60, y: 420, width: 560, height: 100 },
        rows: [
          ["Metric", "Value"],
          ["Edits", "two"],
        ],
        columnWidths: [2, 1],
      },
    ],
  ],
  [
    "setTableCell",
    [
      {
        op: "insertTable",
        pageIndex: 1,
        rect: { x: 60, y: 420, width: 560, height: 100 },
        rows: [
          ["a", "b"],
          ["c", "d"],
        ],
      },
      { op: "setTableCell", target: "$0", row: 1, column: 1, text: "changed" },
    ],
  ],
  [
    "insertSlide",
    [
      { op: "insertSlide", index: 1 },
      { op: "replaceText", target: "$0", text: "New slide from the layout" },
    ],
  ],
  ["duplicateSlide", [{ op: "duplicateSlide", pageIndex: 0 }]],
  ["deleteSlide", [{ op: "deleteSlide", pageIndex: 1 }]],
  ["moveSlide", [{ op: "moveSlide", from: 0, to: 1 }]],
  [
    "everything",
    [
      { op: "insertSlide", index: 2 },
      { op: "replaceText", target: "$0", text: "Everything at once" },
      {
        op: "insertTextBox",
        pageIndex: 2,
        rect: { x: 80, y: 300, width: 500, height: 60 },
        text: "A box",
        style: { bold: true },
      },
      {
        op: "setShapeStyle",
        target: "$2",
        fill: "#DDEBF7",
        line: { color: "#2F5597", width: 1.5 },
      },
      {
        op: "insertImage",
        pageIndex: 2,
        rect: { x: 600, y: 300, width: 100, height: 100 },
        data: new Uint8Array(png),
        mimeType: "image/png",
      },
      { op: "duplicateSlide", pageIndex: 2 },
      { op: "moveSlide", from: 3, to: 0 },
    ],
  ],
];

await mkdir(outdir, { recursive: true });
const signal = new AbortController().signal;
for (const [name, operations] of cases) {
  const engine = await PptxEditEngine.open(
    original,
    defaultResourceLimits,
    signal,
  );
  const issues = await engine.validate(operations, signal);
  if (issues.length > 0)
    throw new Error(`${name}: ${JSON.stringify(issues, null, 2)}`);
  await engine.apply({ stateId: 1, operations }, signal);
  const bytes = await engine.materialize("save", {}, signal);
  await engine.dispose();
  const file = resolve(outdir, `${name}.pptx`);
  await writeFile(file, bytes);
  console.log(`${name}: ${bytes.length} B → ${file}`);
}
