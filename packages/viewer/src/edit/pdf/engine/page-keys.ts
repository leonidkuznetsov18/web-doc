import { ViewerError } from "../../../errors.js";
import { attachInfoDictionary } from "./compact.js";
import type { Pdfium } from "./pdfium.js";

/** Native page dictionary identity survives PDFium full and incremental saves. */
const METADATA_KEY = "WebDocPageKeys";
const MAX_METADATA_BYTES = 1_048_576;
const MAX_KEY_LENGTH = 64;

/** PDFium writes a new Info object but omits its trailer link on no-Info imports. */
export function preservePageKeysMetadata(
  bytes: Uint8Array,
  prefixLength = 0,
): Uint8Array {
  return attachInfoDictionary(bytes, METADATA_KEY, prefixLength);
}

export function isPageKey(value: unknown): value is string {
  return (
    typeof value === "string" &&
    value.length <= MAX_KEY_LENGTH &&
    /^(?:p\d+|q\d+\.\d+(?:~[1-9]\d*)?)$/.test(value)
  );
}

function pageNumbers(pdfium: Pdfium, document: number): number[] {
  const { lib } = pdfium;
  return Array.from({ length: lib.FPDF_GetPageCount(document) }, (_, index) =>
    lib.EPDFDoc_GetPageObjectNumberByIndex(document, index),
  );
}

/** Damaged or externally stale metadata cannot assign a key to another page. */
export function readPageKeys(
  pdfium: Pdfium,
  document: number,
): string[] | undefined {
  const raw = pdfium.readWideString(
    (buffer, bytes) =>
      pdfium.lib.FPDF_GetMetaText(document, METADATA_KEY, buffer, bytes),
    MAX_METADATA_BYTES,
  );
  if (!raw) return undefined;
  let value: unknown;
  try {
    value = JSON.parse(raw);
  } catch {
    return undefined;
  }
  if (
    !value ||
    typeof value !== "object" ||
    !("version" in value) ||
    value.version !== 1 ||
    !("pages" in value) ||
    !Array.isArray(value.pages) ||
    value.pages.length !== pdfium.lib.FPDF_GetPageCount(document)
  )
    return undefined;
  const byObject = new Map<number, string>();
  const used = new Set<string>();
  for (const entry of value.pages) {
    if (
      !entry ||
      typeof entry !== "object" ||
      !("key" in entry) ||
      !isPageKey(entry.key) ||
      !("object" in entry) ||
      typeof entry.object !== "number" ||
      !Number.isSafeInteger(entry.object) ||
      entry.object <= 0 ||
      used.has(entry.key) ||
      byObject.has(entry.object)
    )
      return undefined;
    byObject.set(entry.object, entry.key);
    used.add(entry.key);
  }
  const keys: string[] = [];
  for (const number of pageNumbers(pdfium, document)) {
    const key = byObject.get(number);
    if (key === undefined) return undefined;
    keys.push(key);
  }
  return keys;
}

export function writePageKeys(
  pdfium: Pdfium,
  document: number,
  keys: readonly string[],
): void {
  const numbers = pageNumbers(pdfium, document);
  if (
    numbers.length !== keys.length ||
    new Set(keys).size !== keys.length ||
    keys.some((key) => !isPageKey(key)) ||
    numbers.some((number) => number <= 0)
  )
    throw new ViewerError(
      "edit-failed",
      "The native page identities could not be preserved",
    );
  const raw = JSON.stringify({
    version: 1,
    pages: numbers.map((object, index) => ({ object, key: keys[index] })),
  });
  if ((raw.length + 1) * 2 > MAX_METADATA_BYTES)
    throw new ViewerError(
      "resource-limit",
      "The page identity metadata is too large",
    );
  const pointer = pdfium.writeWideString(raw);
  try {
    if (!pdfium.lib.EPDF_SetMetaText(document, METADATA_KEY, pointer))
      throw new ViewerError(
        "edit-failed",
        "PDFium could not preserve page identities",
      );
  } finally {
    pdfium.free(pointer);
  }
}
