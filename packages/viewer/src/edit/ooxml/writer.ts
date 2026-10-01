import { ViewerError } from "../../errors.js";
import { entryNameOf, partKey } from "./names.js";
import {
  crc32,
  localRecordOf,
  METHOD_DEFLATED,
  METHOD_STORED,
  type ZipArchive,
  type ZipEntry,
} from "./zip.js";

/*
 * Writes a package from an archive and an overlay of changed, added and
 * removed parts. Untouched entries are copied as the exact bytes of their
 * local record and their central record (only the local-header offset is
 * rewritten); changed entries get fresh headers around stored — or, on
 * request, deflated — data; new entries come last with a fixed time; removed
 * entries are left out; the archive comment is kept. The output never needs
 * ZIP64: a package that would is refused before any byte is written.
 */

export interface PartChange {
  /** Absolute part name. */
  readonly name: string;
  readonly bytes: Uint8Array;
}

export interface WriteOverlay {
  /** Replaced parts by absolute name. */
  readonly changed: ReadonlyMap<string, PartChange>;
  /** New parts by absolute name, in insertion order. */
  readonly added: ReadonlyMap<string, PartChange>;
  /** Removed parts by absolute name. */
  readonly removed: ReadonlySet<string>;
}

export interface WriteOptions {
  /** How changed and new entries are written; stored keeps output identical across engines. */
  readonly compression?: "store" | "deflate";
  readonly signal?: AbortSignal;
}

const SIGNATURE_LOCAL = 0x04034b50;
const SIGNATURE_CENTRAL = 0x02014b50;
const SIGNATURE_EOCD = 0x06054b50;
const VERSION_NEEDED = 20;
const VERSION_MADE_BY = 0x0314; // UNIX, 2.0
/** 1980-01-01 00:00 in MS-DOS time and date. */
const FIXED_TIME = 0;
const FIXED_DATE = 0x0021;
const MAX_32 = 0xffffffff;
const MAX_ENTRIES = 0xffff;

const encoder = new TextEncoder();

/** The package bytes for the archive with the overlay applied. */
export async function writeZip(
  archive: ZipArchive,
  overlay: WriteOverlay,
  options: WriteOptions = {},
): Promise<Uint8Array> {
  const compression = options.compression ?? "store";
  const changed = keyed(overlay.changed);
  const added = keyed(overlay.added);
  const removed = new Set([...overlay.removed].map(partKey));
  const locals: Uint8Array[][] = [];
  const centrals: Uint8Array[] = [];
  let offset = 0;
  const seen = new Set<string>();

  for (const entry of archive.entries) {
    throwIfAborted(options.signal);
    const key = partKey(entry.name);
    seen.add(key);
    if (removed.has(key)) continue;
    const change = changed.get(key);
    if (change) {
      const written = await fresh(entry.name, change.bytes, compression, entry);
      locals.push([written.local, written.data]);
      centrals.push(written.central(offset));
      offset += written.local.byteLength + written.data.byteLength;
      continue;
    }
    const local = localRecordOf(archive, entry);
    const record = archive.bytes.subarray(local.headerOffset, local.recordEnd);
    locals.push([record]);
    centrals.push(centralWithOffset(archive, entry, offset));
    offset += record.byteLength;
  }
  for (const [key, change] of added) {
    if (seen.has(key) && !removed.has(key))
      throw new ViewerError("invalid-file", `Part ${key} is added but exists`);
    throwIfAborted(options.signal);
    const name = entryNameOf(change.name);
    const written = await fresh(name, change.bytes, compression);
    locals.push([written.local, written.data]);
    centrals.push(written.central(offset));
    offset += written.local.byteLength + written.data.byteLength;
  }
  if (centrals.length > MAX_ENTRIES)
    throw new ViewerError(
      "unsupported-package",
      "The package would need ZIP64: more than 65,535 entries",
    );
  const centralOffset = offset;
  let centralLength = 0;
  for (const central of centrals) centralLength += central.byteLength;
  if (centralOffset + centralLength > MAX_32)
    throw new ViewerError(
      "unsupported-package",
      "The package would need ZIP64: it exceeds 4 GiB",
    );
  const eocd = new Uint8Array(22 + archive.comment.byteLength);
  const view = new DataView(eocd.buffer);
  view.setUint32(0, SIGNATURE_EOCD, true);
  view.setUint16(4, 0, true);
  view.setUint16(6, 0, true);
  view.setUint16(8, centrals.length, true);
  view.setUint16(10, centrals.length, true);
  view.setUint32(12, centralLength, true);
  view.setUint32(16, centralOffset, true);
  view.setUint16(20, archive.comment.byteLength, true);
  eocd.set(archive.comment, 22);

  const out = new Uint8Array(centralOffset + centralLength + eocd.byteLength);
  let at = 0;
  for (const parts of locals)
    for (const part of parts) {
      out.set(part, at);
      at += part.byteLength;
    }
  for (const central of centrals) {
    out.set(central, at);
    at += central.byteLength;
  }
  out.set(eocd, at);
  return out;
}

function keyed(
  parts: ReadonlyMap<string, PartChange>,
): Map<string, PartChange> {
  const map = new Map<string, PartChange>();
  for (const change of parts.values()) map.set(partKey(change.name), change);
  return map;
}

/** The central record of an untouched entry with the new local-header offset. */
function centralWithOffset(
  archive: ZipArchive,
  entry: ZipEntry,
  localHeaderOffset: number,
): Uint8Array {
  const record = archive.bytes.slice(
    entry.centralOffset,
    entry.centralOffset + entry.centralLength,
  );
  new DataView(record.buffer).setUint32(42, localHeaderOffset, true);
  return record;
}

interface FreshEntry {
  readonly local: Uint8Array;
  readonly data: Uint8Array;
  central(localHeaderOffset: number): Uint8Array;
}

/** Fresh local and central records around stored or deflated data. */
async function fresh(
  name: string,
  bytes: Uint8Array,
  compression: "store" | "deflate",
  previous?: ZipEntry,
): Promise<FreshEntry> {
  if (bytes.byteLength > MAX_32)
    throw new ViewerError(
      "unsupported-package",
      `Part ${name} would need ZIP64: it exceeds 4 GiB`,
    );
  const nameBytes = previous?.nameBytes ?? encoder.encode(name);
  const extra = previous?.extra ?? new Uint8Array(0);
  const comment = previous?.comment ?? new Uint8Array(0);
  const method = compression === "deflate" ? METHOD_DEFLATED : METHOD_STORED;
  const data = method === METHOD_DEFLATED ? await deflateRaw(bytes) : bytes;
  const crc = crc32(bytes);
  const modTime = previous?.modTime ?? FIXED_TIME;
  const modDate = previous?.modDate ?? FIXED_DATE;
  const local = new Uint8Array(30 + nameBytes.byteLength + extra.byteLength);
  const localView = new DataView(local.buffer);
  localView.setUint32(0, SIGNATURE_LOCAL, true);
  localView.setUint16(4, VERSION_NEEDED, true);
  localView.setUint16(6, 0, true);
  localView.setUint16(8, method, true);
  localView.setUint16(10, modTime, true);
  localView.setUint16(12, modDate, true);
  localView.setUint32(14, crc, true);
  localView.setUint32(18, data.byteLength, true);
  localView.setUint32(22, bytes.byteLength, true);
  localView.setUint16(26, nameBytes.byteLength, true);
  localView.setUint16(28, extra.byteLength, true);
  local.set(nameBytes, 30);
  local.set(extra, 30 + nameBytes.byteLength);
  return {
    local,
    data,
    central(localHeaderOffset) {
      const central = new Uint8Array(
        46 + nameBytes.byteLength + extra.byteLength + comment.byteLength,
      );
      const view = new DataView(central.buffer);
      view.setUint32(0, SIGNATURE_CENTRAL, true);
      view.setUint16(4, previous?.versionMadeBy ?? VERSION_MADE_BY, true);
      view.setUint16(6, VERSION_NEEDED, true);
      view.setUint16(8, 0, true);
      view.setUint16(10, method, true);
      view.setUint16(12, modTime, true);
      view.setUint16(14, modDate, true);
      view.setUint32(16, crc, true);
      view.setUint32(20, data.byteLength, true);
      view.setUint32(24, bytes.byteLength, true);
      view.setUint16(28, nameBytes.byteLength, true);
      view.setUint16(30, extra.byteLength, true);
      view.setUint16(32, comment.byteLength, true);
      view.setUint16(34, 0, true);
      view.setUint16(36, previous?.internalAttributes ?? 0, true);
      view.setUint32(38, previous?.externalAttributes ?? 0, true);
      view.setUint32(42, localHeaderOffset, true);
      central.set(nameBytes, 46);
      central.set(extra, 46 + nameBytes.byteLength);
      central.set(comment, 46 + nameBytes.byteLength + extra.byteLength);
      return central;
    },
  };
}

async function deflateRaw(bytes: Uint8Array): Promise<Uint8Array> {
  if (typeof CompressionStream !== "function")
    throw new ViewerError(
      "edit-unsupported",
      'This environment has no CompressionStream; save with compression: "store"',
      { details: { reason: "no-compression-stream" } },
    );
  const stream = new CompressionStream("deflate-raw");
  const writer = stream.writable.getWriter();
  const writing = writer
    .write(bytes.slice() as Uint8Array<ArrayBuffer>)
    .then(() => writer.close());
  const chunks: Uint8Array[] = [];
  let total = 0;
  const reader = stream.readable.getReader();
  for (;;) {
    const chunk = await reader.read();
    if (chunk.done) break;
    chunks.push(chunk.value);
    total += chunk.value.byteLength;
  }
  await writing;
  const out = new Uint8Array(total);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.byteLength;
  }
  return out;
}

function throwIfAborted(signal: AbortSignal | undefined): void {
  if (signal?.aborted)
    throw new ViewerError("aborted", "Writing the package was aborted");
}
