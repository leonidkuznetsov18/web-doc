import { ViewerError } from "../../../errors.js";

/*
 * Compaction of a full PDFium save. PDFium writes every object it holds,
 * including content streams and resources a regenerated page no longer
 * references, so deleted text would stay recoverable inside the file. This
 * pass keeps only the objects reachable from the trailer and rewrites the
 * cross-reference table. It relies on the shape PDFium produces: classic
 * `N G obj … endobj` objects, direct `/Length` values, no object streams and
 * a plain `trailer` dictionary.
 */

interface ParsedObject {
  readonly number: number;
  readonly generation: number;
  /** Byte range of the whole object, header to `endobj` inclusive. */
  readonly start: number;
  readonly end: number;
  /** Object numbers referenced from the object's dictionary and arrays. */
  readonly references: readonly number[];
}

const HEADER = /^(\d+) (\d+) obj\b/;
const REFERENCE = /(\d+) (\d+) R\b/g;
const LENGTH = /\/Length (\d+)/;

/** The reachable objects of `bytes` with a fresh cross-reference table. */
export function compactPdf(bytes: Uint8Array): Uint8Array {
  const text = latin1(bytes);
  const trailerAt = text.lastIndexOf("trailer");
  const xrefAt = text.lastIndexOf("xref", trailerAt);
  if (trailerAt < 0 || xrefAt < 0)
    throw new ViewerError("internal", "The saved PDF has no classic trailer");
  const trailerDictionary = dictionaryAt(text, trailerAt + "trailer".length);
  const objects = parseObjects(text, xrefAt);
  const byNumber = new Map(objects.map((object) => [object.number, object]));

  const roots = [...trailerDictionary.matchAll(REFERENCE)].map((match) =>
    Number(match[1]),
  );
  const reachable = new Set<number>();
  const queue = [...roots];
  while (queue.length > 0) {
    const number = queue.pop()!;
    if (reachable.has(number)) continue;
    const object = byNumber.get(number);
    if (!object) continue;
    reachable.add(number);
    queue.push(...object.references);
  }

  const kept = objects.filter((object) => reachable.has(object.number));
  const headerEnd = objects[0]?.start ?? xrefAt;
  const parts: Uint8Array[] = [bytes.subarray(0, headerEnd)];
  const offsets = new Map<number, number>();
  let position = headerEnd;
  for (const object of kept) {
    offsets.set(object.number, position);
    const slice = bytes.subarray(object.start, object.end);
    parts.push(slice);
    position += slice.byteLength;
    if (!text.endsWith("\n", object.end)) {
      parts.push(NEWLINE);
      position += 1;
    }
  }
  const size = Math.max(0, ...kept.map((object) => object.number)) + 1;
  let xref = `xref\n0 ${size}\n`;
  for (let number = 0; number < size; number += 1) {
    const offset = offsets.get(number);
    xref +=
      offset === undefined
        ? "0000000000 65535 f \n"
        : `${String(offset).padStart(10, "0")} ${String(byNumber.get(number)!.generation).padStart(5, "0")} n \n`;
  }
  const trailer = trailerDictionary.replace(/\/Size \d+/, `/Size ${size}`);
  const tail = `${xref}trailer\n${trailer}\nstartxref\n${position}\n%%EOF\n`;
  parts.push(latin1Bytes(tail));
  return concat(parts);
}

function parseObjects(text: string, until: number): ParsedObject[] {
  const objects: ParsedObject[] = [];
  let cursor = 0;
  while (cursor < until) {
    const headerAt = text.indexOf(" obj", cursor);
    if (headerAt < 0 || headerAt >= until) break;
    const lineStart = text.lastIndexOf("\n", headerAt) + 1;
    const header = HEADER.exec(text.slice(lineStart, headerAt + 4));
    if (!header) {
      cursor = headerAt + 4;
      continue;
    }
    const bodyStart = headerAt + 4;
    const streamAt = text.indexOf("stream", bodyStart);
    const endobjAt = text.indexOf("endobj", bodyStart);
    if (endobjAt < 0)
      throw new ViewerError(
        "internal",
        "The saved PDF has an unterminated object",
      );
    let dictionaryEnd = endobjAt;
    let end = endobjAt + "endobj".length;
    if (streamAt >= 0 && streamAt < endobjAt) {
      // The dictionary ends before `stream`; the data follows the end of that
      // line and runs /Length bytes.
      dictionaryEnd = streamAt;
      const length = LENGTH.exec(text.slice(bodyStart, streamAt));
      if (!length)
        throw new ViewerError("internal", "A stream has no direct /Length");
      let dataStart = streamAt + "stream".length;
      if (text[dataStart] === "\r") dataStart += 1;
      if (text[dataStart] === "\n") dataStart += 1;
      const dataEnd = dataStart + Number(length[1]);
      const endstreamAt = text.indexOf("endstream", dataEnd);
      const realEndobj = text.indexOf("endobj", endstreamAt);
      if (endstreamAt < 0 || realEndobj < 0)
        throw new ViewerError("internal", "A stream is not terminated");
      end = realEndobj + "endobj".length;
    }
    const references = [
      ...text.slice(bodyStart, dictionaryEnd).matchAll(REFERENCE),
    ].map((match) => Number(match[1]));
    objects.push({
      number: Number(header[1]),
      generation: Number(header[2]),
      start: lineStart,
      end,
      references,
    });
    cursor = end;
  }
  return objects;
}

function dictionaryAt(text: string, from: number): string {
  const start = text.indexOf("<<", from);
  let depth = 0;
  for (let index = start; index < text.length - 1; index += 1) {
    if (text.startsWith("<<", index)) {
      depth += 1;
      index += 1;
    } else if (text.startsWith(">>", index)) {
      depth -= 1;
      index += 1;
      if (depth === 0) return text.slice(start, index + 1);
    }
  }
  throw new ViewerError("internal", "The trailer dictionary is not terminated");
}

const NEWLINE = Uint8Array.of(0x0a);

function latin1(bytes: Uint8Array): string {
  let text = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000)
    text += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  return text;
}

function latin1Bytes(text: string): Uint8Array {
  return Uint8Array.from(text, (character) => character.charCodeAt(0) & 0xff);
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(
    parts.reduce((total, part) => total + part.byteLength, 0),
  );
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.byteLength;
  }
  return out;
}
