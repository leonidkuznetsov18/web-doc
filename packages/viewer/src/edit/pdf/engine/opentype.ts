import { cffWithSafeName, type CffFont } from "./cff.js";

/*
 * Fonts a browser loads with `FontFace`. A PDF's CFF program gets a minimal
 * OpenType font around it: the CFF table as the PDF holds it, a Unicode cmap,
 * and the tables a font loader requires (head, hhea, hmtx, maxp, name, OS/2,
 * post), filled from what the CFF says. A PDF's TrueType program keeps its
 * tables; subsetting tools leave out the OS/2, name and post tables and pack
 * tables off four-byte boundaries, which browsers refuse, so those tables are
 * added and the file laid out again. Nothing in a program's tables changes.
 */

const SFNT_TRUETYPE = 0x00010000;
const SFNT_CFF = 0x4f54544f; // "OTTO"

/**
 * The OpenType file, or `undefined` when no glyph has a code point or the
 * program's name cannot be made one a browser accepts.
 */
export function cffOpenType(
  program: Uint8Array,
  font: CffFont,
  unicode: ReadonlyMap<number, number>,
): Uint8Array | undefined {
  const safe = cffWithSafeName(program);
  if (unicode.size === 0 || !safe) return undefined;
  const em = font.unitsPerEm;
  const [xMin, yMin, xMax, yMax] = fontBox(font.bbox, em);
  const glyphCount = font.widths.length;
  const widths = font.widths.map((width) =>
    Math.max(0, Math.min(0xffff, Math.round(width))),
  );
  const codes = [...unicode.keys()].sort((a, b) => a - b);
  const tables = new Map<string, Uint8Array>([
    ["CFF ", safe],
    ["OS/2", os2(em, widths, codes, yMin, yMax)],
    ["cmap", cmap(unicode, codes)],
    ["head", head(em, xMin, yMin, xMax, yMax)],
    ["hhea", hhea(widths, glyphCount, xMin, yMin, xMax, yMax)],
    ["hmtx", concat(widths.map((width) => bytes(u16(width), i16(0))))],
    ["maxp", bytes(u32(0x00005000), u16(glyphCount))],
    ["name", names(postScriptName(font.name))],
    ["post", post(em)],
  ]);
  return assemble(tables, SFNT_CFF);
}

/**
 * The font box in int16 font units, in order; a font without one, which CFF
 * allows, gets a box of the usual proportions so its line has a height.
 */
function fontBox(
  box: readonly number[],
  em: number,
): [number, number, number, number] {
  const clamp = (value: number) =>
    Math.max(-0x8000, Math.min(0x7fff, Math.round(value)));
  const [x0 = 0, y0 = 0, x1 = 0, y1 = 0] = box.map(clamp);
  if (y1 <= y0) return [0, clamp(-em * 0.2), clamp(em), clamp(em * 0.8)];
  return [Math.min(x0, x1), y0, Math.max(x0, x1), y1];
}

/**
 * An embedded TrueType or OpenType program with the OS/2, name and post
 * tables it lacks, laid out on four-byte boundaries; `undefined` when its
 * table directory or core tables are broken.
 */
export function completeSfnt(
  program: Uint8Array,
  fontName: string,
): Uint8Array | undefined {
  const tables = sfntTables(program);
  const headTable = tables?.get("head");
  const hheaTable = tables?.get("hhea");
  const maxpTable = tables?.get("maxp");
  const hmtxTable = tables?.get("hmtx");
  if (
    !tables ||
    !tables.has("cmap") ||
    !headTable ||
    headTable.length < 54 ||
    !hheaTable ||
    hheaTable.length < 36 ||
    !maxpTable ||
    maxpTable.length < 6 ||
    !hmtxTable
  )
    return undefined;
  const view = (table: Uint8Array) =>
    new DataView(table.buffer, table.byteOffset, table.byteLength);
  const headView = view(headTable);
  const em = headView.getUint16(18);
  const yMin = headView.getInt16(38);
  const yMax = headView.getInt16(42);
  // The adjustment is the file's, so it is worked out again.
  const ownHead = headTable.slice();
  ownHead.fill(0, 8, 12);
  tables.set("head", ownHead);
  if (!tables.has("OS/2")) {
    const metrics = view(hheaTable).getUint16(34);
    const widths = Array.from(
      { length: Math.min(metrics, hmtxTable.length >> 2) },
      (_, glyph) => view(hmtxTable).getUint16(glyph * 4),
    );
    tables.set("OS/2", os2(em, widths, [], yMin, yMax));
  }
  if (!tables.has("name")) tables.set("name", names(postScriptName(fontName)));
  if (!tables.has("post")) tables.set("post", post(em));
  const version = view(program).getUint32(0);
  return assemble(tables, version === SFNT_CFF ? SFNT_CFF : SFNT_TRUETYPE);
}

/** The sfnt file maps Unicode: a cmap subtable for platform 0, or for Windows Unicode BMP or full repertoire. */
export function hasUnicodeCmap(program: Uint8Array): boolean {
  const table = sfntTables(program)?.get("cmap");
  if (!table || table.length < 4) return false;
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const count = Math.min(view.getUint16(2), (table.length - 4) >> 3);
  for (let index = 0; index < count; index += 1) {
    const platform = view.getUint16(4 + index * 8);
    const encoding = view.getUint16(6 + index * 8);
    if (
      platform === 0 ||
      (platform === 3 && (encoding === 1 || encoding === 10))
    )
      return true;
  }
  return false;
}

/** The tables of an sfnt file by tag, or `undefined` when its directory is broken. */
function sfntTables(program: Uint8Array): Map<string, Uint8Array> | undefined {
  if (program.length < 12) return undefined;
  const view = new DataView(
    program.buffer,
    program.byteOffset,
    program.byteLength,
  );
  const count = view.getUint16(4);
  if (count === 0 || 12 + count * 16 > program.length) return undefined;
  const tables = new Map<string, Uint8Array>();
  for (let index = 0; index < count; index += 1) {
    const record = 12 + index * 16;
    const tag = String.fromCharCode(...program.subarray(record, record + 4));
    const offset = view.getUint32(record + 8);
    const length = view.getUint32(record + 12);
    if (tables.has(tag) || offset + length > program.length) return undefined;
    tables.set(tag, program.subarray(offset, offset + length));
  }
  return tables;
}

/** The PostScript name without a subset tag, in the characters a name table takes. */
function postScriptName(name: string): string {
  return (
    name
      .replace(/^[A-Z]{6}\+/, "")
      .replace(/[^\x21-\x7e]|[[\](){}<>/%]/g, "")
      .slice(0, 63) || "Embedded"
  );
}

function post(em: number): Uint8Array {
  // Version 3: no glyph names.
  return bytes(
    u32(0x00030000),
    u32(0),
    i16(-Math.round(em / 10)),
    i16(Math.round(em / 20)),
    u32(0),
    u32(0),
    u32(0),
    u32(0),
    u32(0),
  );
}

function head(
  em: number,
  xMin: number,
  yMin: number,
  xMax: number,
  yMax: number,
): Uint8Array {
  return bytes(
    u32(0x00010000),
    u32(0x00010000),
    u32(0), // checkSumAdjustment, set once the font is assembled
    u32(0x5f0f3cf5),
    u16(0x000b),
    u16(em),
    u32(0),
    u32(0),
    u32(0),
    u32(0),
    i16(xMin),
    i16(yMin),
    i16(xMax),
    i16(yMax),
    u16(0),
    u16(8),
    i16(2),
    i16(0),
    i16(0),
  );
}

function hhea(
  widths: readonly number[],
  glyphCount: number,
  xMin: number,
  yMin: number,
  xMax: number,
  yMax: number,
): Uint8Array {
  return bytes(
    u32(0x00010000),
    i16(yMax),
    i16(yMin),
    i16(0),
    u16(widths.reduce((most, width) => Math.max(most, width), 0)),
    i16(xMin),
    i16(0),
    i16(xMax),
    i16(1),
    i16(0),
    i16(0),
    i16(0),
    i16(0),
    i16(0),
    i16(0),
    i16(0),
    u16(glyphCount),
  );
}

function os2(
  em: number,
  widths: readonly number[],
  codes: readonly number[],
  yMin: number,
  yMax: number,
): Uint8Array {
  const scaled = (fraction: number) => Math.round(em * fraction);
  const drawn = widths.filter((width) => width > 0);
  const average = drawn.length
    ? Math.round(drawn.reduce((sum, width) => sum + width, 0) / drawn.length)
    : 0;
  return bytes(
    u16(4),
    i16(average),
    u16(400),
    u16(5),
    u16(0),
    i16(scaled(0.65)),
    i16(scaled(0.6)),
    i16(0),
    i16(scaled(0.075)),
    i16(scaled(0.65)),
    i16(scaled(0.6)),
    i16(0),
    i16(scaled(0.35)),
    i16(scaled(0.05)),
    i16(scaled(0.25)),
    i16(0),
    new Array<number>(10).fill(0),
    u32(0),
    u32(0),
    u32(0),
    u32(0),
    [0x57, 0x44, 0x4f, 0x43], // "WDOC"
    u16(0x0040), // regular
    u16(Math.min(codes[0] ?? 0x20, 0xffff)),
    u16(Math.min(codes.at(-1) ?? 0xffff, 0xffff)),
    i16(yMax),
    i16(yMin),
    i16(0),
    u16(Math.max(0, yMax)),
    u16(Math.max(0, -yMin)),
    u32(1),
    u32(0),
    i16(scaled(0.5)),
    i16(scaled(0.7)),
    u16(0),
    u16(32),
    u16(0),
  );
}

/**
 * Format 4 for the BMP, and format 12 as well when a code point lies beyond
 * it or the BMP takes more segments than format 4 holds. Records are in
 * platform and encoding order, as the specification requires.
 */
function cmap(
  unicode: ReadonlyMap<number, number>,
  codes: readonly number[],
): Uint8Array {
  const basic = codes.filter((code) => code <= 0xffff);
  const format4 = cmapFormat4(unicode, basic);
  const format12 =
    !format4 || codes.length > basic.length
      ? cmapFormat12(unicode, codes)
      : undefined;
  const subtables: { platform: number; encoding: number; data: Uint8Array }[] =
    [
      ...(format4 ? [{ platform: 0, encoding: 3, data: format4 }] : []),
      ...(format12 ? [{ platform: 0, encoding: 4, data: format12 }] : []),
      ...(format4 ? [{ platform: 3, encoding: 1, data: format4 }] : []),
      ...(format12 ? [{ platform: 3, encoding: 10, data: format12 }] : []),
    ];
  // Subtables shared by two records are written once.
  const unique = [...new Set(subtables.map((subtable) => subtable.data))];
  const headerSize = 4 + subtables.length * 8;
  const offsets = new Map<Uint8Array, number>();
  let offset = headerSize;
  for (const data of unique) {
    offsets.set(data, offset);
    offset += data.length;
  }
  return concat([
    bytes(u16(0), u16(subtables.length)),
    ...subtables.map(({ platform, encoding, data }) =>
      bytes(u16(platform), u16(encoding), u32(offsets.get(data)!)),
    ),
    ...unique,
  ]);
}

/** `undefined` when the segments overflow the subtable's 16-bit length. */
function cmapFormat4(
  unicode: ReadonlyMap<number, number>,
  codes: readonly number[],
): Uint8Array | undefined {
  // Runs of consecutive code points on consecutive glyphs share one delta.
  const segments: { start: number; end: number; glyph: number }[] = [];
  for (const code of codes) {
    const glyph = unicode.get(code)!;
    const last = segments.at(-1);
    if (
      last &&
      code === last.end + 1 &&
      glyph === last.glyph + (code - last.start)
    )
      last.end = code;
    else segments.push({ start: code, end: code, glyph });
  }
  segments.push({ start: 0xffff, end: 0xffff, glyph: 0 });
  const count = segments.length;
  if (16 + count * 8 > 0xffff) return undefined;
  const searchRange = 2 * 2 ** Math.floor(Math.log2(count));
  const delta = (segment: (typeof segments)[number]) =>
    segment.start === 0xffff ? 1 : (segment.glyph - segment.start) & 0xffff;
  return bytes(
    u16(4),
    u16(16 + count * 8),
    u16(0),
    u16(count * 2),
    u16(searchRange),
    u16(Math.floor(Math.log2(count))),
    u16(count * 2 - searchRange),
    segments.flatMap((segment) => u16(segment.end)),
    u16(0),
    segments.flatMap((segment) => u16(segment.start)),
    segments.flatMap((segment) => u16(delta(segment))),
    segments.flatMap(() => u16(0)),
  );
}

function cmapFormat12(
  unicode: ReadonlyMap<number, number>,
  codes: readonly number[],
): Uint8Array {
  const groups: { start: number; end: number; glyph: number }[] = [];
  for (const code of codes) {
    const glyph = unicode.get(code)!;
    const last = groups.at(-1);
    if (
      last &&
      code === last.end + 1 &&
      glyph === last.glyph + (code - last.start)
    )
      last.end = code;
    else groups.push({ start: code, end: code, glyph });
  }
  return bytes(
    u16(12),
    u16(0),
    u32(16 + groups.length * 12),
    u32(0),
    u32(groups.length),
    groups.flatMap((group) => [
      ...u32(group.start),
      ...u32(group.end),
      ...u32(group.glyph),
    ]),
  );
}

function names(postScript: string): Uint8Array {
  const records: [number, string][] = [
    [1, postScript],
    [2, "Regular"],
    [3, `${postScript}-web-doc`],
    [4, postScript],
    [5, "Version 1.0"],
    [6, postScript],
  ];
  const strings = records.map(([, text]) =>
    [...text].flatMap((character) => u16(character.charCodeAt(0))),
  );
  let offset = 0;
  const entries = records.map(([id], index) => {
    const entry = [
      ...u16(3),
      ...u16(1),
      ...u16(0x0409),
      ...u16(id),
      ...u16(strings[index]!.length),
      ...u16(offset),
    ];
    offset += strings[index]!.length;
    return entry;
  });
  return bytes(
    u16(0),
    u16(records.length),
    u16(6 + records.length * 12),
    entries.flat(),
    strings.flat(),
  );
}

/** The table directory, the tables padded to four bytes, and the head checksum adjustment. */
function assemble(
  tables: ReadonlyMap<string, Uint8Array>,
  version: number,
): Uint8Array {
  const tags = [...tables.keys()].sort();
  const count = tags.length;
  const selector = Math.floor(Math.log2(count));
  const headerSize = 12 + count * 16;
  let offset = headerSize;
  const placed = tags.map((tag) => {
    const data = tables.get(tag)!;
    const at = offset;
    offset += (data.length + 3) & ~3;
    return { tag, data, at };
  });
  const font = new Uint8Array(offset);
  font.set(
    bytes(
      u32(version),
      u16(count),
      u16(2 ** selector * 16),
      u16(selector),
      u16(count * 16 - 2 ** selector * 16),
      placed.flatMap(({ tag, data, at }) => [
        ...[...tag].map((character) => character.charCodeAt(0)),
        ...u32(checksum(data)),
        ...u32(at),
        ...u32(data.length),
      ]),
    ),
  );
  for (const { data, at } of placed) font.set(data, at);
  const head = placed.find((table) => table.tag === "head")!;
  const adjustment = (0xb1b0afba - checksum(font)) >>> 0;
  font.set(u32(adjustment), head.at + 8);
  return font;
}

function checksum(data: Uint8Array): number {
  let sum = 0;
  for (let offset = 0; offset < data.length; offset += 4)
    sum =
      (sum +
        (((data[offset] ?? 0) << 24) |
          ((data[offset + 1] ?? 0) << 16) |
          ((data[offset + 2] ?? 0) << 8) |
          (data[offset + 3] ?? 0))) >>>
      0;
  return sum;
}

function u16(value: number): number[] {
  return [(value >> 8) & 0xff, value & 0xff];
}

function i16(value: number): number[] {
  return u16(value < 0 ? value + 0x10000 : value);
}

function u32(value: number): number[] {
  return [
    (value >>> 24) & 0xff,
    (value >>> 16) & 0xff,
    (value >>> 8) & 0xff,
    value & 0xff,
  ];
}

function bytes(...parts: (number[] | Uint8Array)[]): Uint8Array {
  return concat(parts.map((part) => Uint8Array.from(part)));
}

function concat(parts: readonly Uint8Array[]): Uint8Array {
  const out = new Uint8Array(parts.reduce((sum, part) => sum + part.length, 0));
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}
