import { ViewerError } from "../../../errors.js";

/*
 * Compaction of a full PDFium save. PDFium writes every object it holds,
 * including content streams and resources a regenerated page no longer
 * references, so deleted text would stay recoverable inside the file. This
 * pass keeps only the objects reachable from the trailer and rewrites the
 * cross-reference table.
 *
 * The file is read with a small PDF lexer, not with string searches: comments,
 * literal and hex strings and names are skipped as units, so an "endobj" or
 * "stream" inside a string cannot end an object, tokens may be separated by
 * any whitespace (a reference split over a line still counts), and a stream's
 * length may be direct or an indirect reference. Nothing is decoded; objects
 * are copied byte for byte. The pass works on the shape PDFium produces —
 * classic objects with a `trailer` dictionary — and refuses anything else.
 */

interface ParsedObject {
  readonly number: number;
  readonly generation: number;
  /** Byte range of the whole object, header to `endobj` inclusive. */
  readonly start: number;
  readonly end: number;
  /** Object numbers referenced from the object's dictionary and arrays. */
  readonly references: readonly number[];
  /** The object's value when it is a bare integer, for indirect lengths. */
  readonly integer?: number;
  /** Byte range of the object's value when it is a dictionary or an array. */
  readonly value?: readonly [number, number];
  /** Byte range of the object's stream data, when it has a stream. */
  readonly data?: readonly [number, number];
}

/** An object of a file PDFium wrote: its dictionary or array as text, its raw stream bytes. */
export interface PdfObjectBytes {
  readonly value?: string;
  readonly data?: Uint8Array;
}

/**
 * The objects of a file PDFium wrote, by number, read with the same lexer
 * as the compaction. Throws `PdfCompactionError` on a shape it cannot read.
 */
export function readObjects(bytes: Uint8Array): Map<number, PdfObjectBytes> {
  return new Map(
    parseFile(bytes).objects.map((object) => [
      object.number,
      {
        ...(object.value
          ? { value: latin1(bytes.subarray(...object.value)) }
          : {}),
        ...(object.data ? { data: bytes.subarray(...object.data) } : {}),
      },
    ]),
  );
}

/** Attach PDFium's newly written Info object when its source had no `/Info`. */
export function attachInfoDictionary(
  bytes: Uint8Array,
  metadataKey: string,
  prefixLength = 0,
): Uint8Array {
  // Incremental saves can have an arbitrary original encoding. Only inspect
  // PDFium's classic appended section; every byte of the original stays intact.
  const appended = bytes.subarray(prefixLength);
  const parsed = parseFile(appended);
  if (dictionaryHasKey(appended.subarray(...parsed.trailerRange), "Info"))
    return bytes;
  const latest = new Map(
    parsed.objects.map((object) => [object.number, object]),
  );
  const candidates = [...latest.values()].filter(
    (object) =>
      !object.data &&
      object.value &&
      dictionaryHasKey(appended.subarray(...object.value), metadataKey),
  );
  if (candidates.length !== 1)
    throw new PdfCompactionError(
      "the page identity Info object is ambiguous or absent",
    );
  const info = candidates[0]!;
  const end = prefixLength + parsed.trailerRange[1] - 2;
  return concat([
    bytes.subarray(0, end),
    ascii(` /Info ${info.number} ${info.generation} R `),
    bytes.subarray(end),
  ]);
}

/** Inspect dictionary entries, ignoring names inside values or stream bytes. */
function dictionaryHasKey(bytes: Uint8Array, key: string): boolean {
  const lexer = new Lexer(bytes);
  const opening = lexer.next();
  if (opening.kind !== "delimiter" || opening.value !== "<<") return false;
  for (;;) {
    const name = lexer.next();
    if (name.kind === "delimiter" && name.value === ">>") return false;
    if (name.kind !== "name")
      throw new PdfCompactionError("a dictionary entry is malformed");
    if (lexer.lastNameIs(key)) return true;
    const first = lexer.peek();
    skipValue(lexer, []);
    if (first.kind !== "number" || lexer.peek().kind !== "number") continue;
    lexer.next();
    const reference = lexer.next();
    if (reference.kind !== "keyword" || reference.value !== "R")
      throw new PdfCompactionError("a dictionary reference is malformed");
  }
}

/** A token with its byte range; `start` is past any whitespace and comments. */
type Token = { readonly start: number; readonly end: number } & (
  | { readonly kind: "number"; readonly value: number }
  | { readonly kind: "keyword"; readonly value: string }
  | {
      readonly kind: "delimiter";
      readonly value: "<<" | ">>" | "[" | "]" | "{" | "}";
    }
  | { readonly kind: "name" | "string" }
  | { readonly kind: "end" }
);

export class PdfCompactionError extends ViewerError {
  constructor(message: string) {
    super("edit-failed", `The full save could not be compacted: ${message}`, {
      details: { stage: "materialize", reason: "pdf-compaction" },
    });
  }
}

/** The reachable objects of `bytes` with a fresh cross-reference table. */
export function compactPdf(bytes: Uint8Array): Uint8Array {
  const { objects, trailer, trailerReferences, headerEnd } = parseFile(bytes);
  const byNumber = new Map(objects.map((object) => [object.number, object]));

  const reachable = new Set<number>();
  const queue = [...trailerReferences];
  while (queue.length > 0) {
    const number = queue.pop()!;
    if (reachable.has(number)) continue;
    const object = byNumber.get(number);
    if (!object) continue;
    reachable.add(number);
    queue.push(...object.references);
  }

  const kept = objects.filter((object) => reachable.has(object.number));
  const parts: Uint8Array[] = [bytes.subarray(0, headerEnd)];
  const offsets = new Map<number, number>();
  let position = headerEnd;
  for (const object of kept) {
    offsets.set(object.number, position);
    const slice = bytes.subarray(object.start, object.end);
    parts.push(slice, NEWLINE);
    position += slice.byteLength + 1;
  }
  // One subsection per run of consecutive numbers, as PDFium writes it, so
  // sparse numbering does not cost twenty bytes per gap.
  const numbers = kept.map((object) => object.number).sort((a, b) => a - b);
  const size = (numbers.at(-1) ?? 0) + 1;
  let xref = "xref\n0 1\n0000000000 65535 f \n";
  for (let index = 0; index < numbers.length;) {
    let count = 1;
    while (numbers[index + count] === numbers[index]! + count) count += 1;
    xref += `${numbers[index]} ${count}\n`;
    for (const number of numbers.slice(index, index + count))
      xref += `${String(offsets.get(number)).padStart(10, "0")} ${String(byNumber.get(number)!.generation).padStart(5, "0")} n \n`;
    index += count;
  }
  const dictionary = trailer.replace(/\/Size\s+\d+/, `/Size ${size}`);
  parts.push(
    ascii(`${xref}trailer\n${dictionary}\nstartxref\n${position}\n%%EOF\n`),
  );
  return concat(parts);
}

function parseFile(bytes: Uint8Array): {
  readonly objects: ParsedObject[];
  readonly trailer: string;
  readonly trailerReferences: readonly number[];
  readonly trailerRange: readonly [number, number];
  readonly headerEnd: number;
} {
  const lexer = new Lexer(bytes);
  const objects: ParsedObject[] = [];
  let trailer: string | undefined;
  let trailerRange: readonly [number, number] | undefined;
  let trailerReferences: number[] = [];
  let headerEnd: number | undefined;
  // Two integers followed by `obj` open an object; a dangling pair is kept
  // until the next token tells what it is.
  let pending: Token[] = [];
  for (;;) {
    const token = lexer.next();
    if (token.kind === "end") break;
    if (token.kind === "number") {
      pending.push(token);
      if (pending.length > 2) pending.shift();
      continue;
    }
    if (token.kind === "keyword" && token.value === "obj") {
      const [first, second] = pending;
      if (
        pending.length !== 2 ||
        first!.kind !== "number" ||
        second!.kind !== "number"
      )
        throw new PdfCompactionError("an object header is malformed");
      // The object starts at its number: the cross-reference offset must
      // point there, not at the whitespace before it.
      const object = parseObject(
        lexer,
        bytes,
        first!.value,
        second!.value,
        first!.start,
      );
      headerEnd ??= object.start;
      objects.push(object);
      pending = [];
      continue;
    }
    pending = [];
    if (token.kind === "keyword" && token.value === "trailer") {
      const dictionaryStart = lexer.skipWhitespace();
      const references: number[] = [];
      const dictionaryEnd = skipValue(lexer, references);
      trailer = latin1(bytes.subarray(dictionaryStart, dictionaryEnd));
      trailerRange = [dictionaryStart, dictionaryEnd];
      trailerReferences = references;
      continue;
    }
    if (token.kind === "keyword" && token.value === "xref") {
      // A classic table: `xref`, subsections of `start count` and entries.
      // PDFium always writes one; its entries are skipped as plain tokens.
      continue;
    }
    if (token.kind === "keyword" && token.value === "startxref") {
      lexer.next();
      continue;
    }
    if (token.kind === "delimiter" && token.value === "<<")
      // A dictionary outside any object: an xref stream's or a stray one.
      throw new PdfCompactionError(
        "the file uses a cross-reference stream instead of a trailer",
      );
  }
  if (trailer === undefined || trailerRange === undefined)
    throw new PdfCompactionError("the file has no trailer dictionary");
  if (headerEnd === undefined)
    throw new PdfCompactionError("the file has no objects");
  // Indirect stream lengths were resolved while parsing where the length
  // object came first; the rest were measured from `endstream`.
  return { objects, trailer, trailerRange, trailerReferences, headerEnd };
}

function parseObject(
  lexer: Lexer,
  bytes: Uint8Array,
  number: number,
  generation: number,
  start: number,
): ParsedObject {
  const references: number[] = [];
  let integer: number | undefined;
  let lengthValue: number | undefined;
  let lengthReference: number | undefined;
  let value: readonly [number, number] | undefined;
  let data: readonly [number, number] | undefined;
  // The object's value: a dictionary, an array, a scalar or nothing.
  const first = lexer.peek();
  if (first.kind === "number") {
    lexer.next();
    const second = lexer.peek();
    if (second.kind === "number") {
      lexer.next();
      const third = lexer.peek();
      if (third.kind === "keyword" && third.value === "R") {
        lexer.next();
        references.push(first.value);
      }
    } else integer = first.value;
  } else if (first.kind !== "keyword" || first.value !== "endobj") {
    const dictionaryStart = lexer.position;
    const dictionaryEnd = skipValue(lexer, references);
    if (
      first.kind === "delimiter" &&
      (first.value === "<<" || first.value === "[")
    )
      value = [dictionaryStart, dictionaryEnd];
    if (first.kind === "delimiter" && first.value === "<<") {
      const length = findLength(bytes, dictionaryStart, dictionaryEnd);
      lengthValue = length?.value;
      lengthReference = length?.reference;
    }
  }
  let next = lexer.next();
  if (next.kind === "keyword" && next.value === "stream") {
    let dataStart = next.end;
    if (bytes[dataStart] === CR) dataStart += 1;
    if (bytes[dataStart] === LF) dataStart += 1;
    let dataEnd: number | undefined;
    if (lengthValue !== undefined) dataEnd = dataStart + lengthValue;
    else if (lengthReference !== undefined) {
      const known = lexer.integerObjects.get(lengthReference);
      if (known !== undefined) dataEnd = dataStart + known;
    }
    if (
      dataEnd === undefined ||
      dataEnd > bytes.length ||
      !followedByEndstream(bytes, dataEnd)
    ) {
      // Length unknown or wrong: measure to the `endstream` that closes the
      // object. Stream data that spells "endstream" would mislead this, but
      // a wrong direct length is the more common fault.
      dataEnd = findEndstream(bytes, dataStart);
      if (dataEnd === undefined)
        throw new PdfCompactionError(`object ${number} has no endstream`);
    }
    data = [dataStart, dataEnd];
    lexer.seek(dataEnd);
    const endstream = lexer.next();
    if (endstream.kind !== "keyword" || endstream.value !== "endstream")
      throw new PdfCompactionError(`object ${number} is not terminated`);
    next = lexer.next();
  }
  if (next.kind !== "keyword" || next.value !== "endobj")
    throw new PdfCompactionError(`object ${number} has no endobj`);
  if (integer !== undefined) lexer.integerObjects.set(number, integer);
  return {
    number,
    generation,
    start,
    end: next.end,
    references,
    ...(integer === undefined ? {} : { integer }),
    ...(value ? { value } : {}),
    ...(data ? { data } : {}),
  };
}

/** Skips one value (dictionary, array, scalar), collecting `n g R` references. */
function skipValue(lexer: Lexer, references: number[]): number {
  let depth = 0;
  let numbers: number[] = [];
  let end = lexer.position;
  do {
    const token = lexer.next();
    end = token.end;
    switch (token.kind) {
      case "end":
        throw new PdfCompactionError("a value is not terminated");
      case "delimiter":
        if (token.value === "<<" || token.value === "[") depth += 1;
        else if (token.value === ">>" || token.value === "]") depth -= 1;
        numbers = [];
        break;
      case "number":
        numbers.push(token.value);
        if (numbers.length > 2) numbers.shift();
        break;
      case "keyword":
        if (token.value === "R" && numbers.length === 2)
          references.push(numbers[0]!);
        else if (depth === 0 && !SCALAR_KEYWORDS.has(token.value))
          throw new PdfCompactionError(`unexpected keyword ${token.value}`);
        numbers = [];
        break;
      default:
        numbers = [];
    }
    if (depth < 0) throw new PdfCompactionError("unbalanced delimiters");
  } while (depth > 0);
  return end;
}

/** `/Length` of a dictionary's bytes: a direct integer or an indirect reference. */
function findLength(
  bytes: Uint8Array,
  start: number,
  end: number,
): { value?: number; reference?: number } | undefined {
  const lexer = new Lexer(bytes.subarray(start, end));
  for (;;) {
    const token = lexer.next();
    if (token.kind === "end") return undefined;
    if (token.kind !== "name" || !lexer.lastNameIs("Length")) continue;
    const value = lexer.next();
    if (value.kind !== "number") return undefined;
    const generation = lexer.peek();
    if (generation.kind !== "number") return { value: value.value };
    lexer.next();
    const reference = lexer.peek();
    if (reference.kind === "keyword" && reference.value === "R")
      return { reference: value.value };
    return { value: value.value };
  }
}

function followedByEndstream(bytes: Uint8Array, at: number): boolean {
  let index = at;
  while (index < bytes.length && isWhitespace(bytes[index]!)) index += 1;
  return matchesAscii(bytes, index, "endstream");
}

function findEndstream(bytes: Uint8Array, from: number): number | undefined {
  for (let index = from; index <= bytes.length - 9; index += 1) {
    if (bytes[index] !== 0x65 || !matchesAscii(bytes, index, "endstream"))
      continue;
    const after = index + 9;
    if (after < bytes.length && !isWhitespace(bytes[after]!)) continue;
    // Trim the end-of-line that precedes `endstream`.
    let end = index;
    if (end > from && bytes[end - 1] === LF) end -= 1;
    if (end > from && bytes[end - 1] === CR) end -= 1;
    return end;
  }
  return undefined;
}

/** A PDF tokenizer that skips comments, strings and names as units. */
class Lexer {
  readonly #bytes: Uint8Array;
  #position = 0;
  #lastName = "";
  /** Objects whose value is a bare integer, as seen so far: indirect lengths. */
  readonly integerObjects = new Map<number, number>();

  constructor(bytes: Uint8Array) {
    this.#bytes = bytes;
  }

  get position(): number {
    return this.#position;
  }

  seek(position: number): void {
    this.#position = position;
  }

  lastNameIs(name: string): boolean {
    return this.#lastName === name;
  }

  peek(): Token {
    const position = this.#position;
    const token = this.next();
    this.#position = position;
    return token;
  }

  /** Skips whitespace and comments; returns where the next token starts. */
  skipWhitespace(): number {
    const bytes = this.#bytes;
    for (;;) {
      while (
        this.#position < bytes.length &&
        isWhitespace(bytes[this.#position]!)
      )
        this.#position += 1;
      if (bytes[this.#position] !== 0x25) return this.#position; // %
      while (
        this.#position < bytes.length &&
        bytes[this.#position] !== LF &&
        bytes[this.#position] !== CR
      )
        this.#position += 1;
    }
  }

  next(): Token {
    const bytes = this.#bytes;
    const start = this.skipWhitespace();
    if (start >= bytes.length) return { kind: "end", start, end: start };
    const byte = bytes[start]!;
    if (byte === 0x28) {
      // ( literal string with nesting and escapes
      let depth = 0;
      let index = start;
      for (; index < bytes.length; index += 1) {
        const current = bytes[index]!;
        if (current === 0x5c)
          index += 1; // backslash escapes the next byte
        else if (current === 0x28) depth += 1;
        else if (current === 0x29 && --depth === 0) break;
      }
      this.#position = Math.min(index + 1, bytes.length);
      return { kind: "string", start, end: this.#position };
    }
    if (byte === 0x3c) {
      if (bytes[start + 1] === 0x3c) {
        this.#position = start + 2;
        return {
          kind: "delimiter",
          value: "<<",
          start,
          end: this.#position,
        };
      }
      let index = start + 1;
      while (index < bytes.length && bytes[index] !== 0x3e) index += 1;
      this.#position = Math.min(index + 1, bytes.length);
      return { kind: "string", start, end: this.#position };
    }
    if (byte === 0x3e && bytes[start + 1] === 0x3e) {
      this.#position = start + 2;
      return {
        kind: "delimiter",
        value: ">>",
        start,
        end: this.#position,
      };
    }
    if (byte === 0x5b || byte === 0x5d || byte === 0x7b || byte === 0x7d) {
      this.#position = start + 1;
      return {
        kind: "delimiter",
        value: String.fromCharCode(byte) as "[" | "]" | "{" | "}",
        start,
        end: this.#position,
      };
    }
    if (byte === 0x2f) {
      let index = start + 1;
      while (index < bytes.length && isRegular(bytes[index]!)) index += 1;
      this.#position = index;
      this.#lastName = latin1(bytes.subarray(start + 1, index));
      return { kind: "name", start, end: index };
    }
    let index = start;
    while (index < bytes.length && isRegular(bytes[index]!)) index += 1;
    if (index === start) index += 1; // a stray delimiter byte
    this.#position = index;
    const text = latin1(bytes.subarray(start, index));
    const value = Number(text);
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(text) && Number.isFinite(value))
      return { kind: "number", value, start, end: index };
    return { kind: "keyword", value: text, start, end: index };
  }
}

/** Keywords that are whole values on their own: `null`, `true`, `false`. */
const SCALAR_KEYWORDS = new Set(["null", "true", "false"]);
const CR = 0x0d;
const LF = 0x0a;
const NEWLINE = Uint8Array.of(LF);

function isWhitespace(byte: number): boolean {
  return (
    byte === 0x20 ||
    byte === LF ||
    byte === CR ||
    byte === 0x09 ||
    byte === 0x0c ||
    byte === 0x00
  );
}

/** Regular characters: anything but whitespace and the delimiters. */
function isRegular(byte: number): boolean {
  return (
    !isWhitespace(byte) &&
    byte !== 0x28 &&
    byte !== 0x29 &&
    byte !== 0x3c &&
    byte !== 0x3e &&
    byte !== 0x5b &&
    byte !== 0x5d &&
    byte !== 0x7b &&
    byte !== 0x7d &&
    byte !== 0x2f &&
    byte !== 0x25
  );
}

function matchesAscii(bytes: Uint8Array, at: number, text: string): boolean {
  if (at + text.length > bytes.length) return false;
  for (let index = 0; index < text.length; index += 1)
    if (bytes[at + index] !== text.charCodeAt(index)) return false;
  return true;
}

function latin1(bytes: Uint8Array): string {
  let text = "";
  for (let offset = 0; offset < bytes.length; offset += 0x8000)
    text += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000));
  return text;
}

function ascii(text: string): Uint8Array {
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
