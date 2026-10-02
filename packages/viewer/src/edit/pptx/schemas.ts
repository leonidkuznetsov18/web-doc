import type { JsonSchema, OperationSchemaSet } from "../types.js";

/*
 * JSON Schemas of the PPTX operations, shared by the worker (engine
 * validation), the main thread (shape checks) and clients (`session.schemas`).
 * Only keywords the in-repo validator implements may appear here. The
 * exported set lists the operations that have an engine handler.
 */

const definitions: Readonly<Record<string, JsonSchema>> = {
  // DrawingML coordinates are bounded (ST_Coordinate: ±27,273,042,316,900 EMU).
  coordinate: {
    type: "number",
    minimum: -2_800_000_000,
    maximum: 2_800_000_000,
  },
  extent: { type: "number", exclusiveMinimum: 0, maximum: 2_800_000_000 },
  rect: {
    type: "object",
    required: ["x", "y", "width", "height"],
    additionalProperties: false,
    properties: {
      x: { $ref: "#/$defs/coordinate" },
      y: { $ref: "#/$defs/coordinate" },
      width: { $ref: "#/$defs/extent" },
      height: { $ref: "#/$defs/extent" },
    },
  },
  point: {
    type: "object",
    required: ["x", "y"],
    additionalProperties: false,
    properties: {
      x: { $ref: "#/$defs/coordinate" },
      y: { $ref: "#/$defs/coordinate" },
    },
  },
  offset: {
    type: "object",
    required: ["dx", "dy"],
    additionalProperties: false,
    properties: {
      dx: { $ref: "#/$defs/coordinate" },
      dy: { $ref: "#/$defs/coordinate" },
    },
  },
  target: { type: "string", minLength: 1, maxLength: 200 },
  textPosition: {
    type: "object",
    required: ["elementId", "offset"],
    additionalProperties: false,
    properties: {
      elementId: { $ref: "#/$defs/target" },
      offset: { type: "integer", minimum: 0 },
    },
  },
  textRange: {
    type: "object",
    required: ["start", "end"],
    additionalProperties: false,
    properties: {
      start: { $ref: "#/$defs/textPosition" },
      end: { $ref: "#/$defs/textPosition" },
    },
  },
  text: { type: "string", maxLength: 100_000 },
  rgb: { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" },
  themeColor: {
    type: "object",
    required: ["theme"],
    additionalProperties: false,
    properties: {
      theme: { type: "string", minLength: 1, maxLength: 20 },
      mods: { type: "object" },
    },
  },
  color: { oneOf: [{ $ref: "#/$defs/rgb" }, { $ref: "#/$defs/themeColor" }] },
  colorOrNone: {
    oneOf: [
      { $ref: "#/$defs/rgb" },
      { $ref: "#/$defs/themeColor" },
      { const: "none" },
    ],
  },
  textStyle: {
    type: "object",
    additionalProperties: false,
    properties: {
      fontFamily: { type: "string", minLength: 1, maxLength: 100 },
      fontSize: { type: "number", minimum: 1, maximum: 400 },
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      underline: { type: "boolean" },
      color: { $ref: "#/$defs/color" },
      align: { enum: ["left", "center", "right", "justify"] },
    },
  },
  line: {
    type: "object",
    required: ["color"],
    additionalProperties: false,
    properties: {
      color: { $ref: "#/$defs/colorOrNone" },
      width: { type: "number", minimum: 0, maximum: 100 },
    },
  },
  tableStyle: {
    type: "object",
    additionalProperties: false,
    properties: {
      firstRow: { type: "boolean" },
      bandRow: { type: "boolean" },
    },
  },
  tableRows: {
    type: "array",
    minItems: 1,
    maxItems: 100,
    items: {
      type: "array",
      minItems: 1,
      maxItems: 20,
      items: { type: "string", maxLength: 10_000 },
    },
  },
  index: { type: "integer", minimum: 0, maximum: 100_000 },
};

function operation(
  name: string,
  properties: Readonly<Record<string, JsonSchema>>,
  required: readonly string[],
): JsonSchema {
  return {
    type: "object",
    required: ["op", ...required],
    additionalProperties: false,
    properties: { op: { const: name }, ...properties },
    $defs: definitions,
  };
}

const all: Readonly<Record<string, JsonSchema>> = {
  replaceText: operation(
    "replaceText",
    {
      target: { $ref: "#/$defs/target" },
      text: { $ref: "#/$defs/text" },
      range: { $ref: "#/$defs/textRange" },
    },
    ["target", "text"],
  ),
  setTextStyle: operation(
    "setTextStyle",
    {
      target: { $ref: "#/$defs/target" },
      range: { $ref: "#/$defs/textRange" },
      style: { $ref: "#/$defs/textStyle" },
    },
    ["target", "style"],
  ),
  setShapeStyle: operation(
    "setShapeStyle",
    {
      target: { $ref: "#/$defs/target" },
      fill: { oneOf: [{ $ref: "#/$defs/colorOrNone" }, { type: "null" }] },
      line: {
        oneOf: [{ $ref: "#/$defs/line" }, { const: "none" }, { type: "null" }],
      },
    },
    ["target"],
  ),
  moveElement: operation(
    "moveElement",
    {
      target: { $ref: "#/$defs/target" },
      to: { $ref: "#/$defs/point" },
      by: { $ref: "#/$defs/offset" },
    },
    ["target"],
  ),
  resizeElement: operation(
    "resizeElement",
    { target: { $ref: "#/$defs/target" }, rect: { $ref: "#/$defs/rect" } },
    ["target", "rect"],
  ),
  deleteElement: operation(
    "deleteElement",
    { target: { $ref: "#/$defs/target" } },
    ["target"],
  ),
  insertTextBox: operation(
    "insertTextBox",
    {
      pageIndex: { $ref: "#/$defs/index" },
      rect: { $ref: "#/$defs/rect" },
      text: { $ref: "#/$defs/text" },
      style: { $ref: "#/$defs/textStyle" },
    },
    ["pageIndex", "rect", "text"],
  ),
  insertImage: operation(
    "insertImage",
    {
      pageIndex: { $ref: "#/$defs/index" },
      rect: { $ref: "#/$defs/rect" },
      data: { "x-binary": true, type: "string", contentEncoding: "base64" },
      mimeType: { enum: ["image/png", "image/jpeg"] },
    },
    ["pageIndex", "rect", "data", "mimeType"],
  ),
  insertTable: operation(
    "insertTable",
    {
      pageIndex: { $ref: "#/$defs/index" },
      rect: { $ref: "#/$defs/rect" },
      rows: { $ref: "#/$defs/tableRows" },
      columnWidths: {
        type: "array",
        minItems: 1,
        maxItems: 20,
        items: { type: "number", exclusiveMinimum: 0 },
      },
      style: { $ref: "#/$defs/tableStyle" },
    },
    ["pageIndex", "rect", "rows"],
  ),
  setTableCell: operation(
    "setTableCell",
    {
      target: { $ref: "#/$defs/target" },
      row: { $ref: "#/$defs/index" },
      column: { $ref: "#/$defs/index" },
      text: { $ref: "#/$defs/text" },
    },
    ["target", "row", "column", "text"],
  ),
  insertSlide: operation(
    "insertSlide",
    {
      index: { $ref: "#/$defs/index" },
      layout: { type: "string", minLength: 1, maxLength: 50 },
    },
    ["index"],
  ),
  duplicateSlide: operation(
    "duplicateSlide",
    { pageIndex: { $ref: "#/$defs/index" }, index: { $ref: "#/$defs/index" } },
    ["pageIndex"],
  ),
  deleteSlide: operation(
    "deleteSlide",
    { pageIndex: { $ref: "#/$defs/index" } },
    ["pageIndex"],
  ),
  moveSlide: operation(
    "moveSlide",
    { from: { $ref: "#/$defs/index" }, to: { $ref: "#/$defs/index" } },
    ["from", "to"],
  ),
};

/** Operations with an engine handler; grows as the tasks land. */
export const IMPLEMENTED_OPERATIONS: readonly string[] = [
  "replaceText",
  "setTextStyle",
  "setShapeStyle",
  "moveElement",
  "resizeElement",
  "deleteElement",
  "insertTextBox",
  "insertImage",
  "insertTable",
  "setTableCell",
  "insertSlide",
  "duplicateSlide",
  "deleteSlide",
  "moveSlide",
];

export const pptxOperationSchemas: OperationSchemaSet = Object.freeze({
  format: "pptx",
  version: 1,
  operations: Object.freeze(
    Object.fromEntries(
      IMPLEMENTED_OPERATIONS.map((name) => [name, all[name]!]),
    ),
  ),
});

/** Every schema, implemented or not, for the conformance test. */
export const pptxOperationSchemaDrafts = all;
