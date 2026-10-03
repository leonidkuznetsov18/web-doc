import { deflateSync, inflateSync } from "node:zlib";

import type { DecodedImage } from "../../src/edit/pdf/engine/images.js";

/*
 * A minimal PNG codec for tests: 8-bit RGBA, no interlace, filter type 0 on
 * encode and all five filters on decode. Browsers decode with
 * createImageBitmap; Node tests use this instead.
 */

export function encodePng(image: DecodedImage): Uint8Array {
  const { width, height, rgba } = image;
  const raw = new Uint8Array((width * 4 + 1) * height);
  for (let row = 0; row < height; row += 1) {
    raw[row * (width * 4 + 1)] = 0;
    raw.set(
      rgba.subarray(row * width * 4, (row + 1) * width * 4),
      row * (width * 4 + 1) + 1,
    );
  }
  const header = new Uint8Array(13);
  const view = new DataView(header.buffer);
  view.setUint32(0, width);
  view.setUint32(4, height);
  header.set([8, 6, 0, 0, 0], 8);
  return concat([
    Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a),
    chunk("IHDR", header),
    chunk("IDAT", new Uint8Array(deflateSync(raw))),
    chunk("IEND", new Uint8Array(0)),
  ]);
}

export function decodePng(bytes: Uint8Array): DecodedImage {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  const colorType = bytes[25]!;
  const channels = colorType === 6 ? 4 : colorType === 2 ? 3 : 1;
  const idat: Uint8Array[] = [];
  let offset = 8;
  while (offset < bytes.length) {
    const length = view.getUint32(offset);
    const type = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8));
    if (type === "IDAT")
      idat.push(bytes.subarray(offset + 8, offset + 8 + length));
    offset += 12 + length;
  }
  const raw = new Uint8Array(inflateSync(concat(idat)));
  const stride = width * channels;
  const rgba = new Uint8Array(width * height * 4);
  let previous = new Uint8Array(stride);
  for (let row = 0; row < height; row += 1) {
    const filter = raw[row * (stride + 1)]!;
    const line = raw.slice(row * (stride + 1) + 1, (row + 1) * (stride + 1));
    for (let i = 0; i < stride; i += 1) {
      const a = i >= channels ? line[i - channels]! : 0;
      const b = previous[i]!;
      const c = i >= channels ? previous[i - channels]! : 0;
      let predictor = 0;
      if (filter === 1) predictor = a;
      else if (filter === 2) predictor = b;
      else if (filter === 3) predictor = (a + b) >> 1;
      else if (filter === 4) {
        const p = a + b - c;
        const pa = Math.abs(p - a);
        const pb = Math.abs(p - b);
        const pc = Math.abs(p - c);
        predictor = pa <= pb && pa <= pc ? a : pb <= pc ? b : c;
      }
      line[i] = (line[i]! + predictor) & 0xff;
    }
    for (let x = 0; x < width; x += 1) {
      const target = (row * width + x) * 4;
      if (channels === 4) rgba.set(line.subarray(x * 4, x * 4 + 4), target);
      else if (channels === 3) {
        rgba.set(line.subarray(x * 3, x * 3 + 3), target);
        rgba[target + 3] = 255;
      } else {
        rgba[target] = rgba[target + 1] = rgba[target + 2] = line[x]!;
        rgba[target + 3] = 255;
      }
    }
    previous = line;
  }
  return { width, height, rgba };
}

/** A gradient with a transparent band, for alpha checks. */
export function samplePng(width = 24, height = 16): Uint8Array {
  const rgba = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y += 1)
    for (let x = 0; x < width; x += 1) {
      const offset = (y * width + x) * 4;
      rgba[offset] = Math.round((x / (width - 1)) * 255);
      rgba[offset + 1] = 64;
      rgba[offset + 2] = Math.round((y / (height - 1)) * 255);
      rgba[offset + 3] = y < height / 2 ? 255 : 90;
    }
  return encodePng({ width, height, rgba });
}

/**
 * Whether a PNG is whole: the signature, then chunks from IHDR to IEND that
 * each fit the file and pass their CRC. A stand-in for the browser's decoder
 * in Node: Chromium's createImageBitmap rejects the picture of
 * tests/fixtures/docx/everything.docx, whose IDAT fails its checksum.
 */
export function pngChecksumsHold(bytes: Uint8Array): boolean {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const signature = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];
  if (!signature.every((byte, index) => bytes[index] === byte)) return false;
  let offset = 8;
  let first = true;
  while (offset + 12 <= bytes.length) {
    const length = view.getUint32(offset);
    if (offset + 12 + length > bytes.length) return false;
    const type = String.fromCharCode(...bytes.subarray(offset + 4, offset + 8));
    const stored = view.getUint32(offset + 8 + length);
    if (crc32(bytes.subarray(offset + 4, offset + 8 + length)) !== stored)
      return false;
    if (first && type !== "IHDR") return false;
    first = false;
    offset += 12 + length;
    if (type === "IEND") return offset === bytes.length;
  }
  return false;
}

function chunk(type: string, data: Uint8Array): Uint8Array {
  const out = new Uint8Array(12 + data.length);
  const view = new DataView(out.buffer);
  view.setUint32(0, data.length);
  out.set(
    [...type].map((c) => c.charCodeAt(0)),
    4,
  );
  out.set(data, 8);
  view.setUint32(8 + data.length, crc32(out.subarray(4, 8 + data.length)));
  return out;
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(
    parts.reduce((total, part) => total + part.length, 0),
  );
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});

function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC_TABLE[(crc ^ byte) & 0xff]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
