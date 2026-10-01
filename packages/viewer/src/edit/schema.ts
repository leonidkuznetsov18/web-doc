import type { JsonSchema, OperationIssue } from "./types.js";

/*
 * A validator for the subset of JSON Schema 2020-12 that operation schemas
 * use. The same schemas are exported to clients, so they must stay valid JSON
 * Schema; the only extension is `x-binary`, which marks a value that may be a
 * `Uint8Array` in-process or a base64 string in JSON.
 */

const SUPPORTED_KEYWORDS = new Set([
  "$schema",
  "$id",
  "$defs",
  "$ref",
  "title",
  "description",
  "default",
  "examples",
  "type",
  "properties",
  "required",
  "additionalProperties",
  "enum",
  "const",
  "minimum",
  "maximum",
  "exclusiveMinimum",
  "exclusiveMaximum",
  "minLength",
  "maxLength",
  "pattern",
  "items",
  "minItems",
  "maxItems",
  "oneOf",
  "contentEncoding",
  "x-binary",
]);

const BASE64 = /^[A-Za-z0-9+/]*={0,2}$/;

type SchemaType =
  "object" | "array" | "string" | "number" | "integer" | "boolean" | "null";

/** Collects every issue of `value` against `schema`; an empty result means valid. */
export function validateSchema(
  value: unknown,
  schema: JsonSchema,
  operationIndex: number,
): OperationIssue[] {
  const issues: OperationIssue[] = [];
  visit(value, schema, schema, "", operationIndex, issues);
  return issues;
}

/**
 * Throws when a schema uses a keyword the validator does not implement, so a
 * constraint can never be silently ignored. Run by the tests over every
 * operation schema.
 */
export function assertSupportedSchema(schema: JsonSchema, path = "#"): void {
  for (const [keyword, child] of Object.entries(schema)) {
    if (!SUPPORTED_KEYWORDS.has(keyword))
      throw new Error(`Unsupported JSON Schema keyword ${keyword} at ${path}`);
    if (keyword === "properties" || keyword === "$defs")
      for (const [name, nested] of Object.entries(asRecord(child)))
        assertSupportedSchema(asRecord(nested), `${path}/${keyword}/${name}`);
    else if (keyword === "items" || keyword === "additionalProperties") {
      if (isRecord(child)) assertSupportedSchema(child, `${path}/${keyword}`);
    } else if (keyword === "oneOf")
      asArray(child).forEach((nested, index) =>
        assertSupportedSchema(asRecord(nested), `${path}/oneOf/${index}`),
      );
  }
}

/** JSON Pointer segment escaping (RFC 6901). */
export function pointerSegment(segment: string | number): string {
  return `/${String(segment).replaceAll("~", "~0").replaceAll("/", "~1")}`;
}

function visit(
  value: unknown,
  schema: JsonSchema,
  root: JsonSchema,
  path: string,
  operationIndex: number,
  issues: OperationIssue[],
): void {
  const issue = (code: string, message: string): void => {
    issues.push({ operationIndex, path, code, message });
  };

  if (typeof schema.$ref === "string") {
    visit(
      value,
      resolveRef(schema.$ref, root),
      root,
      path,
      operationIndex,
      issues,
    );
    return;
  }

  if (schema["x-binary"] === true) {
    if (value instanceof Uint8Array) return;
    if (
      typeof value !== "string" ||
      value.length % 4 !== 0 ||
      !BASE64.test(value)
    )
      issue("binary", "Expected bytes or a base64 string");
    return;
  }

  if (Array.isArray(schema.oneOf)) {
    visitOneOf(
      value,
      schema.oneOf as JsonSchema[],
      root,
      path,
      operationIndex,
      issues,
    );
    return;
  }

  if (schema.const !== undefined && !sameJson(value, schema.const)) {
    issue("const", `Expected ${JSON.stringify(schema.const)}`);
    return;
  }
  if (
    Array.isArray(schema.enum) &&
    !schema.enum.some((candidate) => sameJson(value, candidate))
  ) {
    issue(
      "enum",
      `Expected one of ${schema.enum.map((candidate) => JSON.stringify(candidate)).join(", ")}`,
    );
    return;
  }

  const types = schemaTypes(schema.type);
  if (types && !types.some((type) => hasType(value, type))) {
    issue("type", `Expected ${types.join(" or ")}`);
    return;
  }

  if (typeof value === "number") checkNumber(value, schema, issue);
  else if (typeof value === "string") checkString(value, schema, issue);
  else if (Array.isArray(value))
    checkArray(value, schema, root, path, operationIndex, issues, issue);
  else if (isRecord(value))
    checkObject(value, schema, root, path, operationIndex, issues);
}

function visitOneOf(
  value: unknown,
  branches: readonly JsonSchema[],
  root: JsonSchema,
  path: string,
  operationIndex: number,
  issues: OperationIssue[],
): void {
  const results = branches.map((branch) =>
    collect(value, branch, root, path, operationIndex),
  );
  const matching = results.filter((result) => result.length === 0).length;
  if (matching === 1) return;
  if (matching > 1) {
    issues.push({
      operationIndex,
      path,
      code: "one-of",
      message: "Matches more than one allowed shape",
    });
    return;
  }
  // A branch whose discriminating constants all match explains the failure
  // better than a generic message, for example a known `op` with a bad field.
  const discriminated = branches.findIndex((branch) =>
    discriminatorMatches(value, branch, root),
  );
  if (discriminated >= 0) {
    issues.push(...results[discriminated]!);
    return;
  }
  issues.push({
    operationIndex,
    path,
    code: "one-of",
    message: "Does not match any allowed shape",
  });
}

function collect(
  value: unknown,
  schema: JsonSchema,
  root: JsonSchema,
  path: string,
  operationIndex: number,
): OperationIssue[] {
  const issues: OperationIssue[] = [];
  visit(value, schema, root, path, operationIndex, issues);
  return issues;
}

function discriminatorMatches(
  value: unknown,
  branch: JsonSchema,
  root: JsonSchema,
): boolean {
  const schema =
    typeof branch.$ref === "string" ? resolveRef(branch.$ref, root) : branch;
  if (!isRecord(value) || !isRecord(schema.properties)) return false;
  const constants = Object.entries(schema.properties).filter(
    ([, property]) => isRecord(property) && property.const !== undefined,
  );
  return (
    constants.length > 0 &&
    constants.every(([name, property]) =>
      sameJson(value[name], (property as JsonSchema).const),
    )
  );
}

function checkNumber(
  value: number,
  schema: JsonSchema,
  issue: (code: string, message: string) => void,
): void {
  if (typeof schema.minimum === "number" && value < schema.minimum)
    issue("minimum", `Must be at least ${schema.minimum}`);
  if (typeof schema.maximum === "number" && value > schema.maximum)
    issue("maximum", `Must be at most ${schema.maximum}`);
  if (
    typeof schema.exclusiveMinimum === "number" &&
    value <= schema.exclusiveMinimum
  )
    issue("minimum", `Must be greater than ${schema.exclusiveMinimum}`);
  if (
    typeof schema.exclusiveMaximum === "number" &&
    value >= schema.exclusiveMaximum
  )
    issue("maximum", `Must be less than ${schema.exclusiveMaximum}`);
}

function checkString(
  value: string,
  schema: JsonSchema,
  issue: (code: string, message: string) => void,
): void {
  // Lengths count code points, as JSON Schema specifies.
  const length = [...value].length;
  if (typeof schema.minLength === "number" && length < schema.minLength)
    issue("min-length", `Must have at least ${schema.minLength} characters`);
  if (typeof schema.maxLength === "number" && length > schema.maxLength)
    issue("max-length", `Must have at most ${schema.maxLength} characters`);
  if (
    typeof schema.pattern === "string" &&
    !new RegExp(schema.pattern, "u").test(value)
  )
    issue("pattern", `Must match ${schema.pattern}`);
}

function checkArray(
  value: readonly unknown[],
  schema: JsonSchema,
  root: JsonSchema,
  path: string,
  operationIndex: number,
  issues: OperationIssue[],
  issue: (code: string, message: string) => void,
): void {
  if (typeof schema.minItems === "number" && value.length < schema.minItems)
    issue("min-items", `Must have at least ${schema.minItems} items`);
  if (typeof schema.maxItems === "number" && value.length > schema.maxItems)
    issue("max-items", `Must have at most ${schema.maxItems} items`);
  if (isRecord(schema.items))
    value.forEach((item, index) =>
      visit(
        item,
        schema.items as JsonSchema,
        root,
        `${path}${pointerSegment(index)}`,
        operationIndex,
        issues,
      ),
    );
}

function checkObject(
  value: Readonly<Record<string, unknown>>,
  schema: JsonSchema,
  root: JsonSchema,
  path: string,
  operationIndex: number,
  issues: OperationIssue[],
): void {
  const properties = isRecord(schema.properties) ? schema.properties : {};
  // An explicit `undefined` counts as absent, as it does in JSON.
  for (const name of asArray(schema.required))
    if (typeof name === "string" && value[name] === undefined)
      issues.push({
        operationIndex,
        path: `${path}${pointerSegment(name)}`,
        code: "required",
        message: "Required",
      });
  for (const [name, child] of Object.entries(value)) {
    if (child === undefined) continue;
    const childPath = `${path}${pointerSegment(name)}`;
    const property = properties[name];
    if (isRecord(property)) {
      visit(child, property, root, childPath, operationIndex, issues);
    } else if (schema.additionalProperties === false) {
      issues.push({
        operationIndex,
        path: childPath,
        code: "additional-property",
        message: "Unknown property",
      });
    } else if (isRecord(schema.additionalProperties)) {
      visit(
        child,
        schema.additionalProperties,
        root,
        childPath,
        operationIndex,
        issues,
      );
    }
  }
}

function resolveRef(reference: string, root: JsonSchema): JsonSchema {
  const prefix = "#/$defs/";
  const definitions = isRecord(root.$defs) ? root.$defs : {};
  const target = reference.startsWith(prefix)
    ? definitions[reference.slice(prefix.length)]
    : undefined;
  if (!isRecord(target))
    throw new Error(`Unresolvable JSON Schema reference ${reference}`);
  return target;
}

function schemaTypes(type: unknown): readonly SchemaType[] | undefined {
  if (typeof type === "string") return [type as SchemaType];
  if (Array.isArray(type)) return type as SchemaType[];
  return undefined;
}

function hasType(value: unknown, type: SchemaType): boolean {
  switch (type) {
    case "object":
      return isRecord(value);
    case "array":
      return Array.isArray(value);
    case "string":
      return typeof value === "string";
    case "number":
      return typeof value === "number" && Number.isFinite(value);
    case "integer":
      return Number.isInteger(value);
    case "boolean":
      return typeof value === "boolean";
    case "null":
      return value === null;
  }
}

function sameJson(left: unknown, right: unknown): boolean {
  if (left === right) return true;
  if (Array.isArray(left) && Array.isArray(right))
    return (
      left.length === right.length &&
      left.every((item, index) => sameJson(item, right[index]))
    );
  if (isRecord(left) && isRecord(right)) {
    const leftKeys = Object.keys(left);
    return (
      leftKeys.length === Object.keys(right).length &&
      leftKeys.every((key) => sameJson(left[key], right[key]))
    );
  }
  return false;
}

function isRecord(value: unknown): value is Readonly<Record<string, unknown>> {
  return (
    typeof value === "object" &&
    value !== null &&
    !Array.isArray(value) &&
    !(value instanceof Uint8Array)
  );
}

function asRecord(value: unknown): JsonSchema {
  return isRecord(value) ? value : {};
}

function asArray(value: unknown): readonly unknown[] {
  return Array.isArray(value) ? value : [];
}
