import { deflateRawSync } from "node:zlib";

import { crc32 } from "../../src/edit/ooxml/zip.js";

/*
 * A ZIP writer for tests only: it can produce the shapes the reader must
 * accept (stored, deflated, data descriptors with and without a signature,
 * extra fields, a comment) and the ones it must refuse (ZIP64 markers,
 * encryption flags, unknown methods, lying directories, mismatched local
 * headers, bad CRCs).
 */

export interface ZipFileSpec {
  readonly name: string;
  readonly data: Uint8Array | string;
  /** 0 stored (default for non-XML), 8 deflated (default for text). */
  readonly method?: number;
  /** Write sizes and CRC in a trailing data descriptor; "unsigned" omits its signature. */
  readonly descriptor?: boolean | "unsigned";
  readonly extra?: Uint8Array;
  readonly comment?: string;
  /** Overrides, to build archives the reader must refuse. */
  readonly override?: {
    readonly flags?: number;
    readonly crc?: number;
    readonly centralCompressedSize?: number;
    readonly centralUncompressedSize?: number;
    /** Written in every header, so only inflation can catch the lie. */
    readonly declaredUncompressedSize?: number;
    readonly localName?: string;
    readonly localMethod?: number;
    readonly centralMethod?: number;
    readonly localHeaderOffset?: number;
  };
}

export interface ZipBuildOptions {
  readonly comment?: string;
  /** Prepend a ZIP64 end-of-central-directory locator. */
  readonly zip64Locator?: boolean;
  /** Write 0xffff as the entry count. */
  readonly zip64Count?: boolean;
}

const encoder = new TextEncoder();

export function buildZip(
  files: readonly ZipFileSpec[],
  options: ZipBuildOptions = {},
): Uint8Array {
  const chunks: Uint8Array[] = [];
  let offset = 0;
  const push = (bytes: Uint8Array): void => {
    chunks.push(bytes);
    offset += bytes.byteLength;
  };
  const centrals: Uint8Array[] = [];
  for (const file of files) {
    const data =
      typeof file.data === "string" ? encoder.encode(file.data) : file.data;
    const method = file.method ?? (typeof file.data === "string" ? 8 : 0);
    const stored = method === 8 ? new Uint8Array(deflateRawSync(data)) : data;
    const crc = file.override?.crc ?? crc32(data);
    const nameBytes = encoder.encode(file.name);
    const localNameBytes = encoder.encode(
      file.override?.localName ?? file.name,
    );
    const extra = file.extra ?? new Uint8Array(0);
    const descriptor = file.descriptor ?? false;
    const flags = (file.override?.flags ?? 0) | (descriptor ? 0x0008 : 0);
    const localMethod = file.override?.localMethod ?? method;
    const headerOffset = offset;
    const local = new Uint8Array(
      30 + localNameBytes.byteLength + extra.byteLength,
    );
    const view = new DataView(local.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, flags, true);
    view.setUint16(8, localMethod, true);
    view.setUint16(10, 0, true);
    view.setUint16(12, 0x21, true);
    view.setUint32(14, descriptor ? 0 : crc, true);
    view.setUint32(18, descriptor ? 0 : stored.byteLength, true);
    const declared = file.override?.declaredUncompressedSize ?? data.byteLength;
    view.setUint32(22, descriptor ? 0 : declared, true);
    view.setUint16(26, localNameBytes.byteLength, true);
    view.setUint16(28, extra.byteLength, true);
    local.set(localNameBytes, 30);
    local.set(extra, 30 + localNameBytes.byteLength);
    push(local);
    push(stored);
    if (descriptor) {
      const signed = descriptor !== "unsigned";
      const trailer = new Uint8Array(signed ? 16 : 12);
      const trailerView = new DataView(trailer.buffer);
      let at = 0;
      if (signed) {
        trailerView.setUint32(0, 0x08074b50, true);
        at = 4;
      }
      trailerView.setUint32(at, crc, true);
      trailerView.setUint32(at + 4, stored.byteLength, true);
      trailerView.setUint32(at + 8, declared, true);
      push(trailer);
    }
    const commentBytes = encoder.encode(file.comment ?? "");
    const central = new Uint8Array(
      46 + nameBytes.byteLength + extra.byteLength + commentBytes.byteLength,
    );
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 0x0314, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, flags, true);
    centralView.setUint16(10, file.override?.centralMethod ?? method, true);
    centralView.setUint16(12, 0, true);
    centralView.setUint16(14, 0x21, true);
    centralView.setUint32(16, crc, true);
    centralView.setUint32(
      20,
      file.override?.centralCompressedSize ?? stored.byteLength,
      true,
    );
    centralView.setUint32(
      24,
      file.override?.centralUncompressedSize ?? declared,
      true,
    );
    centralView.setUint16(28, nameBytes.byteLength, true);
    centralView.setUint16(30, extra.byteLength, true);
    centralView.setUint16(32, commentBytes.byteLength, true);
    centralView.setUint16(34, 0, true);
    centralView.setUint16(36, 0, true);
    centralView.setUint32(38, 0, true);
    centralView.setUint32(
      42,
      file.override?.localHeaderOffset ?? headerOffset,
      true,
    );
    central.set(nameBytes, 46);
    central.set(extra, 46 + nameBytes.byteLength);
    central.set(commentBytes, 46 + nameBytes.byteLength + extra.byteLength);
    centrals.push(central);
  }
  const centralOffset = offset;
  for (const central of centrals) push(central);
  const centralLength = offset - centralOffset;
  if (options.zip64Locator) {
    const locator = new Uint8Array(20);
    new DataView(locator.buffer).setUint32(0, 0x07064b50, true);
    push(locator);
  }
  const commentBytes = encoder.encode(options.comment ?? "");
  const eocd = new Uint8Array(22 + commentBytes.byteLength);
  const eocdView = new DataView(eocd.buffer);
  eocdView.setUint32(0, 0x06054b50, true);
  eocdView.setUint16(4, 0, true);
  eocdView.setUint16(6, 0, true);
  const count = options.zip64Count ? 0xffff : files.length;
  eocdView.setUint16(8, count, true);
  eocdView.setUint16(10, count, true);
  eocdView.setUint32(12, centralLength, true);
  eocdView.setUint32(16, centralOffset, true);
  eocdView.setUint16(20, commentBytes.byteLength, true);
  eocd.set(commentBytes, 22);
  push(eocd);
  const out = new Uint8Array(offset);
  let at = 0;
  for (const chunk of chunks) {
    out.set(chunk, at);
    at += chunk.byteLength;
  }
  return out;
}

/** The smallest package the layer accepts: content types and one part. */
export function minimalPackage(
  extraFiles: readonly ZipFileSpec[] = [],
): Uint8Array {
  return buildZip([
    {
      name: "[Content_Types].xml",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>',
    },
    {
      name: "_rels/.rels",
      data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
    },
    ...extraFiles,
  ]);
}
