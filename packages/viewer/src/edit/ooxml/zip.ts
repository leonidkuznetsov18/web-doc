import type { ResourceLimits } from "../../contracts.js";
import { ViewerError } from "../../errors.js";

/*
 * The ZIP container of an OOXML package, read in place. The central
 * directory is parsed once; an entry's local record is located when the
 * entry is first used and must agree with its central record; inflation
 * streams through the platform DecompressionStream, metered against the
 * declared size and the resource limits, and the CRC-32 is checked. Entries
 * the writer later copies verbatim are described by their exact byte ranges.
 * ZIP64, encryption, methods other than stored and deflated, and multi-disk
 * archives are refused; nothing here guesses.
 */

export interface ZipEntry {
  /** The entry name as stored, decoded as UTF-8 (OPC names are ASCII). */
  readonly name: string;
  readonly nameBytes: Uint8Array;
  readonly method: number;
  readonly flags: number;
  readonly crc32: number;
  readonly compressedSize: number;
  readonly uncompressedSize: number;
  readonly localHeaderOffset: number;
  /** Byte range of the central record, for verbatim copies. */
  readonly centralOffset: number;
  readonly centralLength: number;
  readonly extra: Uint8Array;
  readonly comment: Uint8Array;
  readonly modTime: number;
  readonly modDate: number;
  readonly versionMadeBy: number;
  readonly versionNeeded: number;
  readonly internalAttributes: number;
  readonly externalAttributes: number;
}

export interface ZipArchive {
  readonly bytes: Uint8Array;
  /** In central-directory order. */
  readonly entries: readonly ZipEntry[];
  readonly centralDirectoryOffset: number;
  readonly centralDirectoryLength: number;
  readonly eocdOffset: number;
  readonly comment: Uint8Array;
}

/** Where an entry's local record lies: header, data and any data descriptor. */
export interface LocalRecord {
  readonly headerOffset: number;
  readonly dataStart: number;
  readonly dataEnd: number;
  /** After the data descriptor when the entry has one. */
  readonly recordEnd: number;
}

export const METHOD_STORED = 0;
export const METHOD_DEFLATED = 8;

const SIGNATURE_LOCAL = 0x04034b50;
const SIGNATURE_CENTRAL = 0x02014b50;
const SIGNATURE_EOCD = 0x06054b50;
const SIGNATURE_DESCRIPTOR = 0x08074b50;
const SIGNATURE_ZIP64_LOCATOR = 0x07064b50;
const EOCD_LENGTH = 22;
const MAX_COMMENT = 0xffff;
const FLAG_ENCRYPTED = 0x0001;
const FLAG_DESCRIPTOR = 0x0008;
const FLAG_STRONG_ENCRYPTION = 0x0040;
const ZIP64_MARKER = 0xffffffff;

/** Parses the end-of-central-directory record and the central directory. */
export function parseZip(
  bytes: Uint8Array,
  limits: ResourceLimits,
): ZipArchive {
  if (bytes.byteLength > limits.maxInputBytes)
    throw limit(
      "Input exceeds maxInputBytes",
      bytes.byteLength,
      limits.maxInputBytes,
    );
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(bytes, view);
  if (eocdOffset === undefined)
    throw invalid("No end-of-central-directory record");
  if (
    eocdOffset >= 20 &&
    view.getUint32(eocdOffset - 20, true) === SIGNATURE_ZIP64_LOCATOR
  )
    throw unsupported("ZIP64 archives are not supported");
  const thisDisk = view.getUint16(eocdOffset + 4, true);
  const directoryDisk = view.getUint16(eocdOffset + 6, true);
  const countOnDisk = view.getUint16(eocdOffset + 8, true);
  const count = view.getUint16(eocdOffset + 10, true);
  const centralDirectoryLength = view.getUint32(eocdOffset + 12, true);
  const centralDirectoryOffset = view.getUint32(eocdOffset + 16, true);
  const commentLength = view.getUint16(eocdOffset + 20, true);
  if (thisDisk !== 0 || directoryDisk !== 0 || countOnDisk !== count)
    throw unsupported("Multi-disk archives are not supported");
  if (
    count === 0xffff ||
    centralDirectoryLength === ZIP64_MARKER ||
    centralDirectoryOffset === ZIP64_MARKER
  )
    throw unsupported("ZIP64 archives are not supported");
  if (centralDirectoryOffset + centralDirectoryLength > eocdOffset)
    throw invalid("The central directory runs past its end record");
  const comment = bytes.subarray(
    eocdOffset + EOCD_LENGTH,
    eocdOffset + EOCD_LENGTH + commentLength,
  );

  const entries: ZipEntry[] = [];
  let offset = centralDirectoryOffset;
  let expanded = 0;
  for (let index = 0; index < count; index += 1) {
    if (
      offset + 46 > eocdOffset ||
      view.getUint32(offset, true) !== SIGNATURE_CENTRAL
    )
      throw invalid(`Central-directory entry ${index} is missing or malformed`);
    const versionMadeBy = view.getUint16(offset + 4, true);
    const versionNeeded = view.getUint16(offset + 6, true);
    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const modTime = view.getUint16(offset + 12, true);
    const modDate = view.getUint16(offset + 14, true);
    const crc32 = view.getUint32(offset + 16, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const uncompressedSize = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const diskNumber = view.getUint16(offset + 34, true);
    const internalAttributes = view.getUint16(offset + 36, true);
    const externalAttributes = view.getUint32(offset + 38, true);
    const localHeaderOffset = view.getUint32(offset + 42, true);
    const centralLength = 46 + nameLength + extraLength + commentLength;
    if (offset + centralLength > eocdOffset)
      throw invalid(`Central-directory entry ${index} runs past the directory`);
    if (
      compressedSize === ZIP64_MARKER ||
      uncompressedSize === ZIP64_MARKER ||
      localHeaderOffset === ZIP64_MARKER
    )
      throw unsupported("ZIP64 entries are not supported");
    if (diskNumber !== 0)
      throw unsupported("Multi-disk archives are not supported");
    if (flags & (FLAG_ENCRYPTED | FLAG_STRONG_ENCRYPTION))
      throw unsupported("Encrypted archives are not supported");
    if (method !== METHOD_STORED && method !== METHOD_DEFLATED)
      throw unsupported(`Compression method ${method} is not supported`);
    if (uncompressedSize > limits.maxZipEntryBytes)
      throw limit(
        "ZIP entry exceeds maxZipEntryBytes",
        uncompressedSize,
        limits.maxZipEntryBytes,
      );
    expanded += uncompressedSize;
    if (expanded > limits.maxExpandedOfficeBytes)
      throw limit(
        "Office package exceeds maxExpandedOfficeBytes",
        expanded,
        limits.maxExpandedOfficeBytes,
      );
    if (localHeaderOffset + 30 > centralDirectoryOffset)
      throw invalid(`Entry ${index} points outside the archive`);
    const nameBytes = bytes.subarray(offset + 46, offset + 46 + nameLength);
    entries.push({
      name: new TextDecoder("utf-8", { fatal: false }).decode(nameBytes),
      nameBytes,
      method,
      flags,
      crc32,
      compressedSize,
      uncompressedSize,
      localHeaderOffset,
      centralOffset: offset,
      centralLength,
      extra: bytes.subarray(
        offset + 46 + nameLength,
        offset + 46 + nameLength + extraLength,
      ),
      comment: bytes.subarray(
        offset + 46 + nameLength + extraLength,
        offset + centralLength,
      ),
      modTime,
      modDate,
      versionMadeBy,
      versionNeeded,
      internalAttributes,
      externalAttributes,
    });
    offset += centralLength;
  }
  if (offset !== centralDirectoryOffset + centralDirectoryLength)
    throw invalid("The central directory's length does not match its entries");
  return {
    bytes,
    entries,
    centralDirectoryOffset,
    centralDirectoryLength,
    eocdOffset,
    comment,
  };
}

/**
 * Locates an entry's local record and checks it against the central record:
 * same name, same method, and the same sizes and CRC unless the entry
 * carries them in a data descriptor instead.
 */
export function localRecordOf(
  archive: ZipArchive,
  entry: ZipEntry,
): LocalRecord {
  const { bytes } = archive;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const headerOffset = entry.localHeaderOffset;
  if (
    headerOffset + 30 > bytes.byteLength ||
    view.getUint32(headerOffset, true) !== SIGNATURE_LOCAL
  )
    throw invalid(`Entry ${entry.name} has no local header`);
  const flags = view.getUint16(headerOffset + 6, true);
  const method = view.getUint16(headerOffset + 8, true);
  const crc32 = view.getUint32(headerOffset + 14, true);
  const compressedSize = view.getUint32(headerOffset + 18, true);
  const uncompressedSize = view.getUint32(headerOffset + 22, true);
  const nameLength = view.getUint16(headerOffset + 26, true);
  const extraLength = view.getUint16(headerOffset + 28, true);
  const dataStart = headerOffset + 30 + nameLength + extraLength;
  const localName = bytes.subarray(
    headerOffset + 30,
    headerOffset + 30 + nameLength,
  );
  if (!sameBytes(localName, entry.nameBytes))
    throw invalid(
      `Entry ${entry.name} is named differently in its local header`,
    );
  if (method !== entry.method)
    throw invalid(
      `Entry ${entry.name} has a different method in its local header`,
    );
  const hasDescriptor = (flags & FLAG_DESCRIPTOR) !== 0;
  if (
    !hasDescriptor &&
    (crc32 !== entry.crc32 ||
      compressedSize !== entry.compressedSize ||
      uncompressedSize !== entry.uncompressedSize)
  )
    throw invalid(
      `Entry ${entry.name} has different sizes in its local header`,
    );
  const dataEnd = dataStart + entry.compressedSize;
  if (dataEnd > archive.centralDirectoryOffset)
    throw invalid(`Entry ${entry.name} runs into the central directory`);
  let recordEnd = dataEnd;
  if (hasDescriptor) {
    let at = dataEnd;
    if (
      at + 4 <= bytes.byteLength &&
      view.getUint32(at, true) === SIGNATURE_DESCRIPTOR
    )
      at += 4;
    if (at + 12 > bytes.byteLength)
      throw invalid(`Entry ${entry.name} has a truncated data descriptor`);
    const descriptorCrc = view.getUint32(at, true);
    const descriptorCompressed = view.getUint32(at + 4, true);
    const descriptorUncompressed = view.getUint32(at + 8, true);
    if (
      descriptorCrc !== entry.crc32 ||
      descriptorCompressed !== entry.compressedSize ||
      descriptorUncompressed !== entry.uncompressedSize
    )
      throw invalid(
        `Entry ${entry.name} has a data descriptor that disagrees with the directory`,
      );
    recordEnd = at + 12;
  }
  return { headerOffset, dataStart, dataEnd, recordEnd };
}

/** The entry's bytes, inflated when needed, metered and CRC-checked. */
export async function inflateEntry(
  archive: ZipArchive,
  entry: ZipEntry,
  limits: ResourceLimits,
  signal?: AbortSignal,
): Promise<Uint8Array> {
  const local = localRecordOf(archive, entry);
  const data = archive.bytes.subarray(local.dataStart, local.dataEnd);
  let output: Uint8Array;
  if (entry.method === METHOD_STORED) {
    if (data.byteLength !== entry.uncompressedSize)
      throw invalid(
        `Entry ${entry.name} is stored with a different size than declared`,
      );
    output = data.slice();
  } else {
    output = await inflateRaw(data, entry, limits, signal);
  }
  if (crc32(output) !== entry.crc32)
    throw invalid(`Entry ${entry.name} does not match its CRC-32`);
  return output;
}

async function inflateRaw(
  data: Uint8Array,
  entry: ZipEntry,
  limits: ResourceLimits,
  signal: AbortSignal | undefined,
): Promise<Uint8Array> {
  if (typeof DecompressionStream !== "function")
    throw new ViewerError(
      "edit-unsupported",
      "This environment has no DecompressionStream for OOXML parts",
      { details: { reason: "no-decompression-stream" } },
    );
  const expected = entry.uncompressedSize;
  const ceiling = Math.min(expected, limits.maxZipEntryBytes);
  const output = new Uint8Array(expected);
  let total = 0;
  const stream = new DecompressionStream("deflate-raw");
  const writer = stream.writable.getWriter();
  // A copy: the stream wants an ArrayBuffer-backed view, and the slice
  // keeps a shared or detached source from leaking into the stream.
  const writing = writer
    .write(data.slice() as Uint8Array<ArrayBuffer>)
    .then(() => writer.close())
    .catch(() => undefined);
  const reader = stream.readable.getReader();
  try {
    for (;;) {
      if (signal?.aborted)
        throw new ViewerError("aborted", "Inflation was aborted");
      let chunk: ReadableStreamReadResult<Uint8Array>;
      try {
        chunk = await reader.read();
      } catch (error) {
        throw invalid(`Entry ${entry.name} is not valid deflate data`, error);
      }
      if (chunk.done) break;
      if (total + chunk.value.byteLength > ceiling) {
        await reader.cancel().catch(() => undefined);
        throw limit(
          `Entry ${entry.name} inflates beyond its declared size`,
          total + chunk.value.byteLength,
          ceiling,
        );
      }
      output.set(chunk.value, total);
      total += chunk.value.byteLength;
    }
  } finally {
    reader.releaseLock();
    await writing;
  }
  if (total !== expected)
    throw invalid(
      `Entry ${entry.name} inflates to ${total} bytes, not the declared ${expected}`,
    );
  return output;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

/** CRC-32 as ZIP uses it (IEEE 802.3, reflected). */
export function crc32(bytes: Uint8Array, seed = 0): number {
  let crc = (seed ^ 0xffffffff) >>> 0;
  for (let index = 0; index < bytes.byteLength; index += 1)
    crc = CRC_TABLE[(crc ^ bytes[index]!) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function findEndOfCentralDirectory(
  bytes: Uint8Array,
  view: DataView,
): number | undefined {
  const floor = Math.max(0, bytes.byteLength - EOCD_LENGTH - MAX_COMMENT);
  for (
    let offset = bytes.byteLength - EOCD_LENGTH;
    offset >= floor;
    offset -= 1
  ) {
    if (view.getUint32(offset, true) !== SIGNATURE_EOCD) continue;
    const commentLength = view.getUint16(offset + 20, true);
    if (offset + EOCD_LENGTH + commentLength === bytes.byteLength)
      return offset;
  }
  return undefined;
}

function sameBytes(a: Uint8Array, b: Uint8Array): boolean {
  if (a.byteLength !== b.byteLength) return false;
  for (let index = 0; index < a.byteLength; index += 1)
    if (a[index] !== b[index]) return false;
  return true;
}

function invalid(message: string, cause?: unknown): ViewerError {
  return new ViewerError(
    "invalid-file",
    message,
    cause === undefined ? {} : { cause },
  );
}

function unsupported(message: string): ViewerError {
  return new ViewerError("unsupported-package", message);
}

function limit(message: string, actual: number, limit: number): ViewerError {
  return new ViewerError("resource-limit", message, {
    details: { actual, limit },
  });
}
