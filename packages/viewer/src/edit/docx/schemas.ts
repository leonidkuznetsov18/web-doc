import type { JsonSchema, OperationSchemaSet } from "../types.js";

/*
 * JSON Schemas of the DOCX operations, shared by the worker (engine
 * validation), the main thread (shape checks) and clients
 * (`session.schemas`). Only keywords the in-repo validator implements may
 * appear here. The exported set lists the operations that have an engine
 * handler.
 */

const HIGHLIGHTS = [
  "yellow",
  "green",
  "cyan",
  "magenta",
  "blue",
  "red",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "black",
  "white",
];

const definitions: Readonly<Record<string, JsonSchema>> = {
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
  color: {
    oneOf: [
      { $ref: "#/$defs/rgb" },
      { $ref: "#/$defs/themeColor" },
      { const: "auto" },
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
      highlight: { enum: [...HIGHLIGHTS, "none"] },
    },
  },
  paragraphStyle: {
    type: "object",
    additionalProperties: false,
    properties: {
      align: { enum: ["left", "center", "right", "justify"] },
      spacing: {
        type: "object",
        additionalProperties: false,
        properties: {
          before: { type: "number", minimum: 0, maximum: 1584 },
          after: { type: "number", minimum: 0, maximum: 1584 },
          line: { type: "number", minimum: 0.06, maximum: 132 },
        },
      },
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
  size: {
    type: "object",
    required: ["width", "height"],
    additionalProperties: false,
    properties: {
      width: { type: "number", exclusiveMinimum: 0, maximum: 1584 },
      height: { type: "number", exclusiveMinimum: 0, maximum: 1584 },
    },
  },
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

/** `before` or `after`: the element the operation places next to. */
const placement = {
  before: { $ref: "#/$defs/target" },
  after: { $ref: "#/$defs/target" },
};

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
  setParagraphStyle: operation(
    "setParagraphStyle",
    {
      target: { $ref: "#/$defs/target" },
      style: { $ref: "#/$defs/paragraphStyle" },
    },
    ["target", "style"],
  ),
  insertParagraph: operation(
    "insertParagraph",
    {
      ...placement,
      text: { $ref: "#/$defs/text" },
      style: { $ref: "#/$defs/textStyle" },
    },
    ["text"],
  ),
  deleteElement: operation(
    "deleteElement",
    { target: { $ref: "#/$defs/target" } },
    ["target"],
  ),
  moveElement: operation(
    "moveElement",
    { target: { $ref: "#/$defs/target" }, ...placement },
    ["target"],
  ),
  insertTable: operation(
    "insertTable",
    {
      ...placement,
      rows: { $ref: "#/$defs/tableRows" },
      columnWidths: {
        type: "array",
        minItems: 1,
        maxItems: 20,
        items: { type: "number", exclusiveMinimum: 0 },
      },
    },
    ["rows"],
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
  insertImage: operation(
    "insertImage",
    {
      ...placement,
      data: { "x-binary": true, type: "string", contentEncoding: "base64" },
      mimeType: { enum: ["image/png", "image/jpeg"] },
      size: { $ref: "#/$defs/size" },
    },
    ["data", "mimeType", "size"],
  ),
};

/** Operations with an engine handler; grows as the tasks land. */
export const IMPLEMENTED_OPERATIONS: readonly string[] = [
  "replaceText",
  "setTextStyle",
  "setParagraphStyle",
  "insertParagraph",
  "deleteElement",
  "moveElement",
  "insertImage",
  "insertTable",
  "setTableCell",
];

export const docxOperationSchemas: OperationSchemaSet = Object.freeze({
  format: "docx",
  version: 1,
  operations: Object.freeze(
    Object.fromEntries(
      IMPLEMENTED_OPERATIONS.map((name) => [name, all[name]!]),
    ),
  ),
});

/** Every schema, implemented or not, for the conformance test. */
export const docxOperationSchemaDrafts = all;
