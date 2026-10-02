import { mkdir, readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";

/*
 * Writes one edited copy of the corpus document per DOCX operation into
 * artifacts/docx-fixtures/, for the manual check that Word and Pages open
 * every result without a repair prompt. Needs the viewer built
 * (npm run build --workspace web-doc) and the corpus fetched.
 */

const root = resolve(import.meta.dirname, "..");
const { DocxEditEngine } =
  await import("../packages/viewer/dist/edit/docx/engine.js");
const { defaultResourceLimits } =
  await import("../packages/viewer/dist/limits.js");

const source = resolve(root, ".cache/corpus/sample.docx");
const outdir = resolve(root, "artifacts/docx-fixtures");
const original = new Uint8Array(await readFile(source));
const png = Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAFElEQVR4nGP4z8DwHwyBFJAAgAAA//8R7AP8Ky7YKAAAAABJRU5ErkJggg==",
  "base64",
);
const signal = new AbortController().signal;

// Element ids come from the document itself: the body paragraphs in order.
const probe = await DocxEditEngine.open(
  original,
  defaultResourceLimits,
  signal,
);
const paragraphs = (
  await probe.getElements({ kinds: ["paragraph"] }, signal)
).map((element) => element.id);
await probe.dispose();
const [first, second, third] = paragraphs;
if (!first || !second || !third)
  throw new Error("The corpus document needs three body paragraphs");

const cases = [
  [
    "replaceText",
    [
      {
        op: "replaceText",
        target: first,
        text: "Replaced first paragraph.\nA second paragraph split off by the edit.",
      },
    ],
  ],
  [
    "replaceText-range",
    [
      {
        op: "replaceText",
        target: second,
        text: "patched",
        range: {
          start: { elementId: second, offset: 0 },
          end: { elementId: second, offset: 5 },
        },
      },
    ],
  ],
  [
    "setTextStyle",
    [
      {
        op: "setTextStyle",
        target: third,
        style: {
          bold: true,
          italic: true,
          underline: true,
          fontSize: 14,
          fontFamily: "Arial",
          color: { theme: "accent2" },
          highlight: "yellow",
        },
      },
    ],
  ],
  [
    "setTextStyle-range",
    [
      {
        op: "setTextStyle",
        target: first,
        range: {
          start: { elementId: first, offset: 6 },
          end: { elementId: first, offset: 11 },
        },
        style: { color: "#C00000", bold: true },
      },
    ],
  ],
  [
    "setParagraphStyle",
    [
      {
        op: "setParagraphStyle",
        target: second,
        style: {
          align: "center",
          spacing: { before: 12, after: 12, line: 1.5 },
        },
      },
    ],
  ],
  [
    "insertParagraph",
    [
      {
        op: "insertParagraph",
        after: first,
        text: "Inserted after the first paragraph.\tWith a tab.",
        style: { italic: true },
      },
      {
        op: "insertParagraph",
        before: first,
        text: "Inserted before everything.",
      },
    ],
  ],
  ["deleteElement", [{ op: "deleteElement", target: second }]],
  ["moveElement", [{ op: "moveElement", target: third, before: first }]],
  [
    "insertImage",
    [
      {
        op: "insertImage",
        after: second,
        data: new Uint8Array(png),
        mimeType: "image/png",
        size: { width: 72, height: 72 },
      },
    ],
  ],
  [
    "insertTable",
    [
      {
        op: "insertTable",
        after: second,
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
        after: third,
        rows: [
          ["a", "b"],
          ["c", "d"],
        ],
      },
      {
        op: "setTableCell",
        target: "$0",
        row: 1,
        column: 1,
        text: "changed\nover two lines",
      },
    ],
  ],
  [
    "everything",
    [
      { op: "replaceText", target: first, text: "Everything at once." },
      {
        op: "setTextStyle",
        target: first,
        style: { bold: true, color: "#2F5597" },
      },
      { op: "setParagraphStyle", target: first, style: { align: "right" } },
      { op: "insertParagraph", after: first, text: "A new paragraph." },
      {
        op: "insertTable",
        after: "$3",
        rows: [["x", "y"]],
      },
      { op: "setTableCell", target: "$4", row: 0, column: 0, text: "X" },
      {
        op: "insertImage",
        after: "$4",
        data: new Uint8Array(png),
        mimeType: "image/png",
        size: { width: 36, height: 36 },
      },
      { op: "moveElement", target: third, before: second },
      { op: "deleteElement", target: second },
    ],
  ],
];

/** Written as Word tracked changes (module 07): suggestions to accept or reject. */
const tracked = {
  changeMode: "tracked",
  author: "web-doc agent",
  timestamp: new Date().toISOString(),
};
cases.push([
  "tracked-changes",
  [
    {
      op: "setTextStyle",
      target: first,
      style: { bold: true, color: "#C00000" },
    },
    { op: "setParagraphStyle", target: first, style: { align: "center" } },
    {
      op: "replaceText",
      target: first,
      text: "A tracked replacement of the first paragraph.",
    },
    {
      op: "replaceText",
      target: second,
      text: "tracked",
      range: {
        start: { elementId: second, offset: 0 },
        end: { elementId: second, offset: 5 },
      },
    },
    {
      op: "insertParagraph",
      after: third,
      text: "A tracked insertion.\nAnd its second paragraph.",
    },
    { op: "deleteElement", target: third },
  ],
  tracked,
]);

await mkdir(outdir, { recursive: true });
const written = [];
for (const [name, operations, mode] of cases) {
  const engine = await DocxEditEngine.open(
    original,
    defaultResourceLimits,
    signal,
  );
  const issues = await engine.validate(operations, signal, mode);
  if (issues.length > 0)
    throw new Error(`${name}: ${JSON.stringify(issues, null, 2)}`);
  await engine.apply({ stateId: 1, operations, ...mode }, signal);
  const bytes = await engine.materialize("save", {}, signal);
  await engine.dispose();
  const file = resolve(outdir, `${name}.docx`);
  await writeFile(file, bytes);
  written.push(`${name}.docx (${bytes.byteLength} bytes)`);
}
await writeFile(
  resolve(outdir, "README.md"),
  [
    "# DOCX edit fixtures",
    "",
    "One edited copy of `.cache/corpus/sample.docx` per DOCX operation, written by",
    "`npm run fixtures:docx`. Open each file in Word and in Pages and record whether",
    "it opens without a repair prompt in `docs/document-editing/todo/06-docx-edit.md`.",
    "`tracked-changes.docx` holds the same edits as Word revisions: check that they",
    "show as suggestions by `web-doc agent`, that accepting them gives the direct",
    "edit's text and rejecting them the original (`docs/document-editing/todo/07-ai-edit.md`).",
    "",
    ...written.map((entry) => `- ${entry}`),
    "",
  ].join("\n"),
);
console.log(`Wrote ${written.length} DOCX fixtures to ${outdir}`);
