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
  target: { type: "string", minLength: 1, maxLength: 200 },
  point: {
    type: "object",
    required: ["x", "y"],
    additionalProperties: false,
    properties: { x: { type: "number" }, y: { type: "number" } },
  },
  offset: {
    type: "object",
    required: ["dx", "dy"],
    additionalProperties: false,
    properties: { dx: { type: "number" }, dy: { type: "number" } },
  },
  stroke: {
    type: "object",
    required: ["color", "width"],
    additionalProperties: false,
    properties: {
      color: { $ref: "#/$defs/color" },
      width: { type: "number", minimum: 0, maximum: 100 },
    },
  },
  fill: {
    type: "object",
    required: ["color"],
    additionalProperties: false,
    properties: { color: { $ref: "#/$defs/color" } },
  },
  strokeOrNull: {
    type: ["object", "null"],
    required: ["color", "width"],
    additionalProperties: false,
    properties: {
      color: { $ref: "#/$defs/color" },
      width: { type: "number", minimum: 0, maximum: 100 },
    },
  },
  fillOrNull: {
    type: ["object", "null"],
    required: ["color"],
    additionalProperties: false,
    properties: { color: { $ref: "#/$defs/color" } },
  },
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
  tableStyle: {
    type: "object",
    additionalProperties: false,
    properties: {
      fontFamily: { type: "string", minLength: 1, maxLength: 100 },
      fontSize: { type: "number", minimum: 1, maximum: 500 },
      color: { $ref: "#/$defs/color" },
      borderColor: { $ref: "#/$defs/color" },
      borderWidth: { type: "number", minimum: 0, maximum: 20 },
      cellPadding: { type: "number", minimum: 0, maximum: 100 },
      headerFill: { $ref: "#/$defs/color" },
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
      items: { type: "string", maxLength: 2000 },
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
    replaceText: operation(
      "replaceText",
      {
        target: { $ref: "#/$defs/target" },
        text: { type: "string", minLength: 1, maxLength: 20000 },
      },
      ["target", "text"],
    ),
    setTextStyle: operation(
      "setTextStyle",
      {
        target: { $ref: "#/$defs/target" },
        style: { $ref: "#/$defs/textBoxStyle" },
      },
      ["target", "style"],
    ),
    resizeElement: operation(
      "resizeElement",
      {
        target: { $ref: "#/$defs/target" },
        rect: { $ref: "#/$defs/rect" },
      },
      ["target", "rect"],
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
    deleteElement: operation(
      "deleteElement",
      { target: { $ref: "#/$defs/target" } },
      ["target"],
    ),
    insertPage: operation(
      "insertPage",
      {
        index: { type: "integer", minimum: 0 },
        size: {
          type: "object",
          required: ["width", "height"],
          additionalProperties: false,
          properties: {
            width: { type: "number", minimum: 3, maximum: 14400 },
            height: { type: "number", minimum: 3, maximum: 14400 },
          },
        },
      },
      ["index"],
    ),
    deletePage: operation(
      "deletePage",
      { pageIndex: { type: "integer", minimum: 0 } },
      ["pageIndex"],
    ),
    movePage: operation(
      "movePage",
      {
        from: { type: "integer", minimum: 0 },
        to: { type: "integer", minimum: 0 },
      },
      ["from", "to"],
    ),
    rotatePage: operation(
      "rotatePage",
      {
        pageIndex: { type: "integer", minimum: 0 },
        rotation: { enum: [0, 90, 180, 270] },
      },
      ["pageIndex", "rotation"],
    ),
    insertShape: operation(
      "insertShape",
      {
        pageIndex: { type: "integer", minimum: 0 },
        shape: { enum: ["rectangle", "ellipse", "line"] },
        rect: { $ref: "#/$defs/rect" },
        from: { $ref: "#/$defs/point" },
        to: { $ref: "#/$defs/point" },
        stroke: { $ref: "#/$defs/stroke" },
        fill: { $ref: "#/$defs/fill" },
      },
      ["pageIndex", "shape"],
    ),
    setShapeStyle: operation(
      "setShapeStyle",
      {
        target: { $ref: "#/$defs/target" },
        stroke: { $ref: "#/$defs/strokeOrNull" },
        fill: { $ref: "#/$defs/fillOrNull" },
      },
      ["target"],
    ),
    insertImage: operation(
      "insertImage",
      {
        pageIndex: { type: "integer", minimum: 0 },
        rect: { $ref: "#/$defs/rect" },
        data: { "x-binary": true, type: "string", contentEncoding: "base64" },
        mimeType: { enum: ["image/png", "image/jpeg"] },
      },
      ["pageIndex", "rect", "data", "mimeType"],
    ),
    insertTable: operation(
      "insertTable",
      {
        pageIndex: { type: "integer", minimum: 0 },
        at: { $ref: "#/$defs/point" },
        width: { type: "number", exclusiveMinimum: 0 },
        rows: { $ref: "#/$defs/tableRows" },
        columnWidths: {
          type: "array",
          minItems: 1,
          maxItems: 20,
          items: { type: "number", exclusiveMinimum: 0 },
        },
        style: { $ref: "#/$defs/tableStyle" },
      },
      ["pageIndex", "at", "width", "rows"],
    ),
    setTableCell: operation(
      "setTableCell",
      {
        target: { $ref: "#/$defs/target" },
        row: { type: "integer", minimum: 0 },
        column: { type: "integer", minimum: 0 },
        text: { type: "string", maxLength: 2000 },
      },
      ["target", "row", "column", "text"],
    ),
  }),
});

/**
 * What a `WebDoc` mark found in a file must look like before its objects are
 * treated as one of web-doc's own elements. Anything else stays plain objects.
 */
/** The head mark of a table, on its path objects, holding every input. */
export const tableMarkSchema: JsonSchema = {
  type: "object",
  required: ["kind", "id", "at", "rows", "columnWidths", "style"],
  additionalProperties: false,
  properties: {
    kind: { const: "table" },
    id: { $ref: "#/$defs/target" },
    at: { $ref: "#/$defs/point" },
    rows: { $ref: "#/$defs/tableRows" },
    columnWidths: {
      type: "array",
      minItems: 1,
      maxItems: 20,
      items: { type: "number", exclusiveMinimum: 0 },
    },
    style: {
      type: "object",
      required: [
        "fontFamily",
        "fontSize",
        "color",
        "borderColor",
        "borderWidth",
        "cellPadding",
        "headerFill",
      ],
      additionalProperties: false,
      properties: {
        ...(definitions.tableStyle!.properties as Record<string, JsonSchema>),
        headerFill: { type: ["string", "null"], pattern: "^#[0-9A-Fa-f]{6}$" },
      },
    },
  },
  $defs: definitions,
};

/** The mark on a table's text objects: the id alone. */
export const tableMemberMarkSchema: JsonSchema = {
  type: "object",
  required: ["kind", "id"],
  additionalProperties: false,
  properties: { kind: { const: "table" }, id: { $ref: "#/$defs/target" } },
  $defs: definitions,
};

export const textBoxMarkSchema: JsonSchema = {
  type: "object",
  required: ["kind", "id", "rect", "text", "style"],
  additionalProperties: false,
  properties: {
    kind: { const: "textBox" },
    id: { $ref: "#/$defs/target" },
    rect: { $ref: "#/$defs/rect" },
    text: { type: "string", minLength: 1, maxLength: 20000 },
    style: {
      type: "object",
      required: [
        "fontFamily",
        "fontSize",
        "bold",
        "italic",
        "color",
        "align",
        "lineHeight",
      ],
      additionalProperties: false,
      properties: definitions.textBoxStyle!.properties as JsonSchema,
    },
  },
  $defs: definitions,
};
