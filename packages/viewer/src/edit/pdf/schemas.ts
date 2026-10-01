import type { OperationSchemaSet } from "../types.js";

/**
 * JSON Schemas of the PDF operations, shared by the worker (engine
 * validation), the main thread (shape checks) and clients (`session.schemas`).
 */
export const pdfOperationSchemas: OperationSchemaSet = Object.freeze({
  format: "pdf",
  version: 1,
  operations: Object.freeze({}),
});
