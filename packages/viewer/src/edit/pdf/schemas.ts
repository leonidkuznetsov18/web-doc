import type { JsonSchema, OperationSchemaSet } from "../types.js";

/*
 * JSON Schemas of the PDF operations, shared by the worker (engine
 * validation), the main thread (shape checks) and clients (`session.schemas`).
 * Only keywords the in-repo validator implements may appear here.
 */

const definitions: Readonly<Record<string, JsonSchema>> = {
  rect: {
    type: "object",
    required: ["x", "y", "width", "height"],
    additionalProperties: false,
    properties: {
      x: { type: "number" },
      y: { type: "number" },
      width: { type: "number", exclusiveMinimum: 0 },
      height: { type: "number", exclusiveMinimum: 0 },
    },
  },
  color: { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" },
  textBoxStyle: {
    type: "object",
    additionalProperties: false,
    properties: {
      fontFamily: { type: "string", minLength: 1, maxLength: 100 },
      fontSize: { type: "number", minimum: 1, maximum: 500 },
      bold: { type: "boolean" },
      italic: { type: "boolean" },
      color: { $ref: "#/$defs/color" },
      align: { enum: ["left", "center", "right"] },
      lineHeight: { type: "number", minimum: 0.5, maximum: 5 },
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

export const pdfOperationSchemas: OperationSchemaSet = Object.freeze({
  format: "pdf",
  version: 1,
  operations: Object.freeze({
    insertTextBox: operation(
      "insertTextBox",
      {
        pageIndex: { type: "integer", minimum: 0 },
        rect: { $ref: "#/$defs/rect" },
        text: { type: "string", minLength: 1, maxLength: 20000 },
        style: { $ref: "#/$defs/textBoxStyle" },
      },
      ["pageIndex", "rect", "text"],
    ),
  }),
});
