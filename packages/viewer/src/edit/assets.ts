import { ViewerError } from "../errors.js";
import type { BinaryData, JsonSchema } from "./types.js";

/*
 * Binary payloads of operations live in a content-addressed store for the
 * session: an operation refers to them as `asset:<sha-256 hex>`, so the
 * history, undo, redo, dry runs and recovery never copy image bytes again.
 */

const REFERENCE = /^asset:[0-9a-f]{64}$/;

export function isAssetReference(value: unknown): value is string {
  return typeof value === "string" && REFERENCE.test(value);
}

/** The content-addressed id of `bytes`. */
export async function assetIdOf(bytes: Uint8Array): Promise<string> {
  const digest = await globalThis.crypto.subtle.digest(
    "SHA-256",
    bytes.slice().buffer,
  );
  let hex = "";
  for (const byte of new Uint8Array(digest))
    hex += byte.toString(16).padStart(2, "0");
  return `asset:${hex}`;
}

/** Something that answers an asset reference with its bytes. */
export interface AssetSource {
  get(id: string): Uint8Array | undefined;
}

export class AssetStore implements AssetSource {
  readonly #bytes = new Map<string, Uint8Array>();
  #total = 0;

  get(id: string): Uint8Array | undefined {
    return this.#bytes.get(id);
  }

  has(id: string): boolean {
    return this.#bytes.has(id);
  }

  set(id: string, bytes: Uint8Array): void {
    if (this.#bytes.has(id)) return;
    this.#bytes.set(id, bytes);
    this.#total += bytes.byteLength;
  }

  /** Bytes held, for limits. */
  get byteLength(): number {
    return this.#total;
  }
}

/**
 * Bytes of a `BinaryData` value that passed schema validation: inline bytes,
 * a base64 string, or an asset reference the store knows.
 */
export function resolveBinary(
  data: BinaryData,
  assets: AssetSource,
): Uint8Array {
  if (data instanceof Uint8Array) return data;
  if (isAssetReference(data)) {
    const bytes = assets.get(data);
    if (!bytes)
      throw new ViewerError("invalid-operation", `Unknown asset ${data}`, {
        details: {
          issues: [
            {
              operationIndex: -1,
              path: "",
              code: "unknown-asset",
              message: `Unknown asset ${data}`,
            },
          ],
        },
      });
    return bytes;
  }
  return Uint8Array.from(atob(data), (character) => character.charCodeAt(0));
}

/** Names of an operation's top-level properties its schema marks as binary. */
export function binaryFields(schema: JsonSchema | undefined): string[] {
  const properties = schema?.properties;
  if (!properties || typeof properties !== "object") return [];
  return Object.entries(properties as Record<string, unknown>)
    .filter(
      ([, property]) =>
        typeof property === "object" &&
        property !== null &&
        (property as Record<string, unknown>)["x-binary"] === true,
    )
    .map(([name]) => name);
}
