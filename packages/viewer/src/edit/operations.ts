import { ViewerError } from "../errors.js";
import { pointerSegment, validateSchema } from "./schema.js";
import type {
  BinaryData,
  OperationIssue,
  OperationSchemaSet,
} from "./types.js";

/**
 * Checks the batch-level preconditions of `apply()`: an array that is not empty
 * and not larger than the limit.
 */
export function assertBatchSize(operations: unknown, limit: number): void {
  if (!Array.isArray(operations) || operations.length === 0)
    throw invalidOperationError([
      {
        operationIndex: -1,
        path: "",
        code: "empty-batch",
        message: "A batch needs at least one operation",
      },
    ]);
  if (operations.length > limit)
    throw new ViewerError(
      "resource-limit",
      "The batch exceeds maxEditOperations",
      { details: { operations: operations.length, limit } },
    );
}

/**
 * Shape and schema checks for every operation of a batch: JSON-only values, a
 * known `op`, and the operation's JSON Schema. All issues are collected.
 */
export function checkOperations(
  operations: readonly unknown[],
  schemas: OperationSchemaSet,
): OperationIssue[] {
  const issues: OperationIssue[] = [];
  operations.forEach((operation, operationIndex) => {
    const shape = jsonIssues(operation, "", operationIndex, new Set());
    if (shape.length > 0) {
      issues.push(...shape);
      return;
    }
    if (!isPlainObject(operation) || typeof operation.op !== "string") {
      issues.push({
        operationIndex,
        path: "/op",
        code: "required",
        message: "An operation is an object with a string op",
      });
      return;
    }
    const schema = schemas.operations[operation.op];
    if (!schema) {
      issues.push({
        operationIndex,
        path: "/op",
        code: "unknown-operation",
        message: `Unknown ${schemas.format} operation ${operation.op}`,
      });
      return;
    }
    issues.push(...validateSchema(operation, schema, operationIndex));
  });
  return issues;
}

export function invalidOperationError(
  issues: readonly OperationIssue[],
): ViewerError {
  const first = issues[0];
  return new ViewerError(
    "invalid-operation",
    first
      ? `Invalid operation ${first.operationIndex} at ${first.path || "/"}: ${first.message}`
      : "Invalid operation",
    {
      details: {
        issues: Object.freeze(
          issues.map((issue) => Object.freeze({ ...issue })),
        ),
      },
    },
  );
}

/** Bytes of a `BinaryData` value that already passed schema validation. */
export function decodeBinary(data: BinaryData): Uint8Array {
  if (data instanceof Uint8Array) return data;
  return Uint8Array.from(atob(data), (character) => character.charCodeAt(0));
}

/** Deep copy of a validated batch, so later caller mutations cannot reach the history. */
/** The operation index a `"$<n>"` same-batch target names, if it is one. */
export function parseReference(target: string): number | undefined {
  const match = /^\$(\d{1,4})$/.exec(target);
  return match ? Number(match[1]) : undefined;
}

export function freezeOperations<T>(operations: readonly T[]): readonly T[] {
  const copies = new WeakMap<object, unknown>();
  return Object.freeze(
    operations.map((operation) => deepCopy(operation, copies) as T),
  );
}

function deepCopy(value: unknown, copies: WeakMap<object, unknown>): unknown {
  if (value instanceof Uint8Array) return value.slice();
  if (typeof value === "object" && value !== null && copies.has(value))
    return copies.get(value);
  // Snapshot cycles as cycles: queued JSON validation can then report their
  // precise path instead of the copier overflowing before apply returns.
  if (Array.isArray(value)) {
    const copy: unknown[] = new Array(value.length);
    copies.set(value, copy);
    value.forEach((child, index) => {
      copy[index] = deepCopy(child, copies);
    });
    return Object.freeze(copy);
  }
  if (isPlainObject(value)) {
    const copy: Record<string, unknown> = Object.fromEntries(
      Object.entries(value).filter(([, child]) => child !== undefined),
    );
    copies.set(value, copy);
    for (const [key, child] of Object.entries(copy))
      copy[key] = deepCopy(child, copies);
    return Object.freeze(copy);
  }
  return value;
}

function jsonIssues(
  value: unknown,
  path: string,
  operationIndex: number,
  ancestors: Set<object>,
): OperationIssue[] {
  const issue = (message: string): OperationIssue[] => [
    { operationIndex, path, code: "not-json", message },
  ];
  if (value === null || typeof value === "string" || typeof value === "boolean")
    return [];
  if (typeof value === "number")
    return Number.isFinite(value) ? [] : issue("Numbers must be finite");
  if (value instanceof Uint8Array) return [];
  if (typeof value !== "object")
    return issue(`A ${typeof value} is not a JSON value`);
  if (ancestors.has(value)) return issue("Cyclic value");
  const issues: OperationIssue[] = [];
  ancestors.add(value);
  if (Array.isArray(value)) {
    for (let index = 0; index < value.length; index += 1) {
      const childPath = `${path}${pointerSegment(index)}`;
      if (!(index in value) || value[index] === undefined)
        issues.push({
          operationIndex,
          path: childPath,
          code: "not-json",
          message: "Arrays cannot hold holes or undefined",
        });
      else
        issues.push(
          ...jsonIssues(value[index], childPath, operationIndex, ancestors),
        );
    }
  } else if (isPlainObject(value)) {
    for (const [key, child] of Object.entries(value))
      if (child !== undefined)
        issues.push(
          ...jsonIssues(
            child,
            `${path}${pointerSegment(key)}`,
            operationIndex,
            ancestors,
          ),
        );
  } else {
    issues.push(...issue("Class instances are not JSON values"));
  }
  ancestors.delete(value);
  return issues;
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  if (typeof value !== "object" || value === null || Array.isArray(value))
    return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}
