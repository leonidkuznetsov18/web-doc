import assert from "node:assert/strict";
import { describe, it } from "node:test";

import {
  assertBatchSize,
  checkOperations,
  decodeBinary,
  freezeOperations,
  invalidOperationError,
} from "../src/edit/operations.js";
import { assertSupportedSchema, validateSchema } from "../src/edit/schema.js";
import type { OperationIssue, OperationSchemaSet } from "../src/index.js";
import { ViewerError } from "../src/index.js";

const color = { type: "string", pattern: "^#[0-9A-Fa-f]{6}$" } as const;

const schemas: OperationSchemaSet = {
  format: "pdf",
  version: 1,
  operations: {
    paint: {
      type: "object",
      required: ["op", "target", "color"],
      additionalProperties: false,
      properties: {
        op: { const: "paint" },
        target: { type: "string", minLength: 1 },
        color,
        opacity: { type: "number", minimum: 0, maximum: 1 },
        mode: { enum: ["fill", "stroke"] },
        points: {
          type: "array",
          maxItems: 3,
          items: { $ref: "#/$defs/point" },
        },
        "a/b~c": { type: "integer" },
      },
      $defs: {
        point: {
          type: "object",
          required: ["x", "y"],
          additionalProperties: false,
          properties: { x: { type: "number" }, y: { type: "number" } },
        },
      },
    },
    attach: {
      type: "object",
      required: ["op", "data"],
      additionalProperties: false,
      properties: {
        op: { const: "attach" },
        data: { "x-binary": true, type: "string", contentEncoding: "base64" },
      },
    },
    shape: {
      type: "object",
      required: ["op", "geometry"],
      properties: {
        op: { const: "shape" },
        geometry: {
          oneOf: [
            {
              type: "object",
              required: ["kind", "width"],
              additionalProperties: false,
              properties: {
                kind: { const: "box" },
                width: { type: "number", exclusiveMinimum: 0 },
              },
            },
            {
              type: "object",
              required: ["kind", "to"],
              additionalProperties: false,
              properties: {
                kind: { const: "line" },
                to: { type: "number" },
              },
            },
          ],
        },
      },
    },
  },
};

function codes(issues: readonly OperationIssue[]): string[] {
  return issues.map(
    (issue) => `${issue.operationIndex}${issue.path}:${issue.code}`,
  );
}

describe("operation shape checks and schema validation", () => {
  it("accepts operations that match their schema", () => {
    assert.deepEqual(
      checkOperations(
        [
          {
            op: "paint",
            target: "e1",
            color: "#00aaFF",
            opacity: 0.5,
            mode: "fill",
            points: [{ x: 1, y: 2 }],
            "a/b~c": 3,
          },
          { op: "attach", data: new Uint8Array([1, 2]) },
          { op: "attach", data: "AQID" },
          { op: "shape", geometry: { kind: "line", to: 4 } },
        ],
        schemas,
      ),
      [],
    );
  });

  it("collects every issue with JSON pointers and stable codes", () => {
    const issues = checkOperations(
      [
        {
          op: "paint",
          color: "red",
          opacity: 2,
          mode: "glow",
          extra: true,
          points: [{ x: 1 }, { x: 1, y: "2" }, { x: 0, y: 0 }, { x: 0, y: 0 }],
          "a/b~c": 1.5,
        },
      ],
      schemas,
    );
    assert.deepEqual(codes(issues).sort(), [
      "0/a~1b~0c:type",
      "0/color:pattern",
      "0/extra:additional-property",
      "0/mode:enum",
      "0/opacity:maximum",
      "0/points/0/y:required",
      "0/points/1/y:type",
      "0/points:max-items",
      "0/target:required",
    ]);
  });

  it("rejects values that are not JSON", () => {
    const cyclic: Record<string, unknown> = { op: "paint" };
    cyclic.self = cyclic;
    const issues = checkOperations(
      [
        { op: "paint", target: () => "e1", color: "#000000" },
        { op: "paint", target: "e1", color: "#000000", opacity: Number.NaN },
        { op: "paint", target: new Date(0), color: "#000000" },
        { op: "paint", target: "e1", color: "#000000", points: [undefined] },
        cyclic,
        "paint",
      ],
      schemas,
    );
    assert.deepEqual(codes(issues), [
      "0/target:not-json",
      "1/opacity:not-json",
      "2/target:not-json",
      "3/points/0:not-json",
      "4/self:not-json",
      "5/op:required",
    ]);
  });

  it("checks binary payloads and unknown operations", () => {
    const issues = checkOperations(
      [
        { op: "attach", data: "not base64!" },
        { op: "paint", target: new Uint8Array(1), color: "#000000" },
        { op: "erase" },
      ],
      schemas,
    );
    assert.deepEqual(codes(issues), [
      "0/data:binary",
      "1/target:type",
      "2/op:unknown-operation",
    ]);
    assert.deepEqual(decodeBinary("AQID"), Uint8Array.of(1, 2, 3));
  });

  it("explains oneOf failures through the branch that matches the discriminator", () => {
    const issues = checkOperations(
      [
        { op: "shape", geometry: { kind: "box", width: 0 } },
        { op: "shape", geometry: { kind: "circle" } },
      ],
      schemas,
    );
    assert.deepEqual(codes(issues), [
      "0/geometry/width:minimum",
      "1/geometry:one-of",
    ]);
  });

  it("refuses schemas that use keywords the validator ignores", () => {
    for (const schema of Object.values(schemas.operations))
      assertSupportedSchema(schema);
    assert.throws(
      () => assertSupportedSchema({ type: "string", format: "email" }),
      /format/,
    );
    assert.throws(
      () =>
        assertSupportedSchema({
          properties: { nested: { anyOf: [{ type: "string" }] } },
        }),
      /anyOf at #\/properties\/nested/,
    );
    assert.deepEqual(validateSchema("x", { type: ["string", "null"] }, 0), []);
  });

  it("checks the batch size before anything else", () => {
    assert.throws(
      () => assertBatchSize([], 10),
      (error: unknown) => {
        assert.ok(error instanceof ViewerError);
        assert.equal(error.code, "invalid-operation");
        const issues = error.details?.issues as readonly OperationIssue[];
        assert.equal(issues[0]?.code, "empty-batch");
        return true;
      },
    );
    assert.throws(
      () => assertBatchSize([{}, {}, {}], 2),
      (error: unknown) =>
        error instanceof ViewerError && error.code === "resource-limit",
    );
    assertBatchSize([{}], 1);
  });

  it("freezes errors and stored batches against later mutation", () => {
    const error = invalidOperationError([
      { operationIndex: 0, path: "/color", code: "pattern", message: "Bad" },
    ]);
    assert.equal(error.message, "Invalid operation 0 at /color: Bad");
    assert.equal(Object.isFrozen(error.details?.issues), true);

    const data = Uint8Array.of(9);
    const source = [{ op: "attach", data, note: undefined, tags: ["a"] }];
    const stored = freezeOperations(source);
    source[0]!.tags.push("b");
    data[0] = 0;
    assert.deepEqual(stored, [
      { op: "attach", data: Uint8Array.of(9), tags: ["a"] },
    ]);
    assert.equal(Object.isFrozen(stored[0]), true);
    assert.equal("note" in stored[0]!, false);
  });
});
