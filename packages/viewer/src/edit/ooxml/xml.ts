import { ViewerError } from "../../errors.js";

/*
 * An offset-preserving scanner for OOXML parts. It tokenizes the decoded
 * text of a part once and builds an element tree whose every node carries
 * its exact range in that text, plus the regions it does not interpret
 * (declaration, comments, processing instructions, CDATA, DOCTYPE). It is
 * non-validating: it checks well-formedness as far as nesting, attribute
 * syntax and a single root go, resolves namespace prefixes, decodes the
 * predefined and numeric entities, and never re-serializes anything. Patches
 * splice the text by these ranges and leave every other character as it is.
 */

export interface XmlRange {
  readonly kind: "declaration" | "comment" | "pi" | "cdata" | "doctype";
  readonly start: number;
  readonly end: number;
}

export interface XmlAttribute {
  /** Qualified, as written. */
  readonly name: string;
  /** Entities decoded. */
  readonly value: string;
  /** As written, without the quotes. */
  readonly rawValue: string;
  /** Of the attribute name. */
  readonly start: number;
  /** After the closing quote. */
  readonly end: number;
}

export interface XmlElement {
  /** Qualified, as written: "p:sp". */
  readonly name: string;
  readonly prefix: string;
  readonly local: string;
  /** Resolved through xmlns declarations in scope, when declared. */
  readonly namespace?: string;
  readonly attributes: readonly XmlAttribute[];
  /** "<" of the start tag. */
  readonly start: number;
  /** After ">" of the end tag, or after "/>". */
  readonly end: number;
  /** The content between the tags; empty and equal for a self-closing element. */
  readonly contentStart: number;
  readonly contentEnd: number;
  readonly selfClosing: boolean;
  readonly children: readonly XmlElement[];
  readonly parent?: XmlElement;
  /** Index among the parent's element children. */
  readonly index: number;
  /** Child-index path from the root. */
  readonly path: readonly number[];
}

export interface XmlPart {
  readonly name: string;
  /** The decoded text the ranges index into (UTF-16 code units). */
  readonly text: string;
  readonly encoding: "utf-8";
  readonly hasBom: boolean;
  /** The overlay revision the part was scanned at; a patch from another revision is stale. */
  readonly revision: number;
  readonly declaration: XmlRange | undefined;
  readonly root: XmlElement;
  /** Comments, processing instructions, CDATA sections and a DOCTYPE, in document order. */
  readonly opaque: readonly XmlRange[];
  find(name: string, under?: XmlElement): XmlElement | undefined;
  findAll(name: string, under?: XmlElement): readonly XmlElement[];
  at(path: readonly number[]): XmlElement | undefined;
  /** Concatenated character data of a subtree, entities decoded, CDATA included. */
  textOf(node: XmlElement): string;
  attribute(node: XmlElement, name: string): string | undefined;
}

const encoder = new TextEncoder();
const BOM = "﻿";

/** Decodes a part as UTF-8; a part that is not UTF-8 cannot be patched. */
export function decodePart(
  name: string,
  bytes: Uint8Array,
): { readonly text: string; readonly hasBom: boolean } {
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true, ignoreBOM: true }).decode(
      bytes,
    );
  } catch (error) {
    throw new ViewerError("unsupported-part", `Part ${name} is not UTF-8`, {
      cause: error,
      details: { part: name },
    });
  }
  return { text, hasBom: text.startsWith(BOM) };
}

/** The bytes of a part's text; `decodePart` guarantees they equal the original. */
export function encodePart(text: string): Uint8Array {
  return encoder.encode(text);
}

/** Scans a part's text; throws `malformed-xml` for what it cannot read. */
export function scanXml(name: string, text: string, revision = 0): XmlPart {
  return new Scanner(name, text, revision).scan();
}

/** Escapes character data for element content. */
export function escapeText(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

/** Escapes an attribute value and quotes it with double quotes. */
export function escapeAttribute(value: string): string {
  return `"${value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll('"', "&quot;")
    .replaceAll("\t", "&#9;")
    .replaceAll("\n", "&#10;")
    .replaceAll("\r", "&#13;")}"`;
}

/** Decodes the predefined and numeric entities; anything else stays as written. */
export function decodeEntities(raw: string): string {
  if (!raw.includes("&")) return raw;
  return raw.replace(
    /&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);/g,
    (match, body: string) => {
      switch (body) {
        case "lt":
          return "<";
        case "gt":
          return ">";
        case "amp":
          return "&";
        case "quot":
          return '"';
        case "apos":
          return "'";
        default: {
          const code =
            body[1] === "x"
              ? Number.parseInt(body.slice(2), 16)
              : Number.parseInt(body.slice(1), 10);
          return Number.isFinite(code) && code >= 0 && code <= 0x10ffff
            ? String.fromCodePoint(code)
            : match;
        }
      }
    },
  );
}

interface MutableElement {
  name: string;
  prefix: string;
  local: string;
  namespace?: string;
  attributes: XmlAttribute[];
  start: number;
  end: number;
  contentStart: number;
  contentEnd: number;
  selfClosing: boolean;
  children: MutableElement[];
  parent?: MutableElement;
  index: number;
  path: number[];
  scope: ReadonlyMap<string, string>;
}

class Scanner {
  readonly #name: string;
  readonly #text: string;
  readonly #revision: number;
  #at = 0;
  readonly #opaque: XmlRange[] = [];
  #declaration: XmlRange | undefined;
  #root: MutableElement | undefined;
  readonly #stack: MutableElement[] = [];

  constructor(name: string, text: string, revision: number) {
    this.#name = name;
    this.#text = text;
    this.#revision = revision;
  }

  scan(): XmlPart {
    const text = this.#text;
    const hasBom = text.startsWith(BOM);
    this.#at = hasBom ? 1 : 0;
    while (this.#at < text.length) {
      const open = text.indexOf("<", this.#at);
      if (open < 0) {
        this.#characterData(this.#at, text.length);
        this.#at = text.length;
        break;
      }
      if (open > this.#at) this.#characterData(this.#at, open);
      this.#at = open;
      if (text.startsWith("<?", open)) this.#processingInstruction();
      else if (text.startsWith("<!--", open)) this.#comment();
      else if (text.startsWith("<![CDATA[", open)) this.#cdata();
      else if (text.startsWith("<!DOCTYPE", open)) this.#doctype();
      else if (text.startsWith("</", open)) this.#endTag();
      else this.#startTag();
    }
    if (this.#stack.length > 0)
      throw this.#malformed(
        `element ${this.#stack.at(-1)!.name} is never closed`,
        text.length,
      );
    if (!this.#root) throw this.#malformed("no root element", 0);
    return new ScannedPart(
      this.#name,
      text,
      hasBom,
      this.#revision,
      this.#declaration,
      this.#root as XmlElement,
      this.#opaque,
    );
  }

  #characterData(start: number, end: number): void {
    if (this.#stack.length > 0) return;
    // Outside the root only whitespace may appear.
    for (let index = start; index < end; index += 1) {
      const code = this.#text.charCodeAt(index);
      if (code !== 0x20 && code !== 0x09 && code !== 0x0a && code !== 0x0d)
        throw this.#malformed("text outside the root element", index);
    }
  }

  #processingInstruction(): void {
    const close = this.#text.indexOf("?>", this.#at + 2);
    if (close < 0)
      throw this.#malformed("unterminated processing instruction", this.#at);
    const isDeclaration =
      this.#text.startsWith("<?xml", this.#at) &&
      /^<\?xml[\s?]/.test(this.#text.slice(this.#at, this.#at + 6));
    const range: XmlRange = {
      kind: isDeclaration ? "declaration" : "pi",
      start: this.#at,
      end: close + 2,
    };
    if (isDeclaration) {
      if (this.#declaration || this.#root || this.#stack.length > 0)
        throw this.#malformed("an XML declaration after the start", this.#at);
      this.#declaration = range;
    } else this.#opaque.push(range);
    this.#at = close + 2;
  }

  #comment(): void {
    const close = this.#text.indexOf("-->", this.#at + 4);
    if (close < 0) throw this.#malformed("unterminated comment", this.#at);
    this.#opaque.push({ kind: "comment", start: this.#at, end: close + 3 });
    this.#at = close + 3;
  }

  #cdata(): void {
    if (this.#stack.length === 0)
      throw this.#malformed("CDATA outside the root element", this.#at);
    const close = this.#text.indexOf("]]>", this.#at + 9);
    if (close < 0)
      throw this.#malformed("unterminated CDATA section", this.#at);
    this.#opaque.push({ kind: "cdata", start: this.#at, end: close + 3 });
    this.#at = close + 3;
  }

  #doctype(): void {
    if (this.#root || this.#stack.length > 0)
      throw this.#malformed("DOCTYPE after the root element", this.#at);
    // The internal subset may hold ">" inside brackets.
    let depth = 0;
    let index = this.#at + 9;
    for (; index < this.#text.length; index += 1) {
      const char = this.#text[index];
      if (char === "[") depth += 1;
      else if (char === "]") depth -= 1;
      else if (char === ">" && depth <= 0) break;
    }
    if (index >= this.#text.length)
      throw this.#malformed("unterminated DOCTYPE", this.#at);
    this.#opaque.push({ kind: "doctype", start: this.#at, end: index + 1 });
    this.#at = index + 1;
  }

  #endTag(): void {
    const text = this.#text;
    const nameStart = this.#at + 2;
    const nameEnd = this.#nameEnd(nameStart);
    if (nameEnd === nameStart)
      throw this.#malformed("end tag without a name", this.#at);
    const name = text.slice(nameStart, nameEnd);
    let close = nameEnd;
    while (close < text.length && isWhitespace(text.charCodeAt(close)))
      close += 1;
    if (text[close] !== ">")
      throw this.#malformed(`end tag </${name}> is not closed`, this.#at);
    const open = this.#stack.pop();
    if (!open)
      throw this.#malformed(`end tag </${name}> without a start tag`, this.#at);
    if (open.name !== name)
      throw this.#malformed(
        `end tag </${name}> closes <${open.name}>`,
        this.#at,
      );
    open.contentEnd = this.#at;
    open.end = close + 1;
    this.#at = close + 1;
  }

  #startTag(): void {
    const text = this.#text;
    const start = this.#at;
    const nameStart = start + 1;
    const nameEnd = this.#nameEnd(nameStart);
    if (nameEnd === nameStart)
      throw this.#malformed("start tag without a name", start);
    const name = text.slice(nameStart, nameEnd);
    const attributes: XmlAttribute[] = [];
    let index = nameEnd;
    let selfClosing = false;
    for (;;) {
      while (index < text.length && isWhitespace(text.charCodeAt(index)))
        index += 1;
      if (index >= text.length)
        throw this.#malformed(`start tag <${name}> is not closed`, start);
      const char = text[index];
      if (char === ">") {
        index += 1;
        break;
      }
      if (char === "/") {
        if (text[index + 1] !== ">")
          throw this.#malformed(`stray "/" in <${name}>`, index);
        selfClosing = true;
        index += 2;
        break;
      }
      const attributeStart = index;
      const attributeNameEnd = this.#nameEnd(index);
      if (attributeNameEnd === index)
        throw this.#malformed(`bad attribute in <${name}>`, index);
      const attributeName = text.slice(index, attributeNameEnd);
      index = attributeNameEnd;
      while (index < text.length && isWhitespace(text.charCodeAt(index)))
        index += 1;
      if (text[index] !== "=")
        throw this.#malformed(`attribute ${attributeName} has no value`, index);
      index += 1;
      while (index < text.length && isWhitespace(text.charCodeAt(index)))
        index += 1;
      const quote = text[index];
      if (quote !== '"' && quote !== "'")
        throw this.#malformed(
          `attribute ${attributeName} is not quoted`,
          index,
        );
      const valueEnd = text.indexOf(quote, index + 1);
      if (valueEnd < 0)
        throw this.#malformed(
          `attribute ${attributeName} is not closed`,
          index,
        );
      const rawValue = text.slice(index + 1, valueEnd);
      if (rawValue.includes("<"))
        throw this.#malformed(`attribute ${attributeName} contains "<"`, index);
      attributes.push({
        name: attributeName,
        value: decodeEntities(rawValue),
        rawValue,
        start: attributeStart,
        end: valueEnd + 1,
      });
      index = valueEnd + 1;
    }
    const parent = this.#stack.at(-1);
    if (!parent && this.#root)
      throw this.#malformed("a second root element", start);
    const colon = name.indexOf(":");
    const prefix = colon < 0 ? "" : name.slice(0, colon);
    const local = colon < 0 ? name : name.slice(colon + 1);
    const scope = scopeOf(parent?.scope, attributes);
    const namespace = scope.get(prefix);
    const element: MutableElement = {
      name,
      prefix,
      local,
      ...(namespace === undefined ? {} : { namespace }),
      attributes,
      start,
      end: index,
      contentStart: index,
      contentEnd: index,
      selfClosing,
      children: [],
      ...(parent ? { parent } : {}),
      index: parent ? parent.children.length : 0,
      path: parent ? [...parent.path, parent.children.length] : [],
      scope,
    };
    if (parent) parent.children.push(element);
    else this.#root = element;
    if (selfClosing) {
      element.contentStart = element.contentEnd = index - 2;
    } else this.#stack.push(element);
    this.#at = index;
  }

  #nameEnd(from: number): number {
    const text = this.#text;
    let index = from;
    while (index < text.length) {
      const code = text.charCodeAt(index);
      if (
        isWhitespace(code) ||
        code === 0x3e || // >
        code === 0x2f || // /
        code === 0x3d || // =
        code === 0x3c || // <
        code === 0x22 || // "
        code === 0x27 // '
      )
        break;
      index += 1;
    }
    return index;
  }

  #malformed(message: string, at: number): ViewerError {
    return new ViewerError("malformed-xml", `Part ${this.#name}: ${message}`, {
      details: { part: this.#name, offset: at },
    });
  }
}

function scopeOf(
  parent: ReadonlyMap<string, string> | undefined,
  attributes: readonly XmlAttribute[],
): ReadonlyMap<string, string> {
  let scope: Map<string, string> | undefined;
  for (const attribute of attributes) {
    if (attribute.name === "xmlns")
      (scope ??= new Map(parent)).set("", attribute.value);
    else if (attribute.name.startsWith("xmlns:"))
      (scope ??= new Map(parent)).set(attribute.name.slice(6), attribute.value);
  }
  return scope ?? parent ?? EMPTY_SCOPE;
}

const EMPTY_SCOPE: ReadonlyMap<string, string> = new Map();

function isWhitespace(code: number): boolean {
  return code === 0x20 || code === 0x09 || code === 0x0a || code === 0x0d;
}

class ScannedPart implements XmlPart {
  readonly encoding = "utf-8" as const;
  readonly #opaqueByStart: readonly XmlRange[];

  constructor(
    readonly name: string,
    readonly text: string,
    readonly hasBom: boolean,
    readonly revision: number,
    readonly declaration: XmlRange | undefined,
    readonly root: XmlElement,
    readonly opaque: readonly XmlRange[],
  ) {
    this.#opaqueByStart = opaque;
  }

  find(name: string, under: XmlElement = this.root): XmlElement | undefined {
    const matches = matcher(name);
    const stack: XmlElement[] = [under];
    while (stack.length > 0) {
      const node = stack.pop()!;
      if (matches(node)) return node;
      for (let index = node.children.length - 1; index >= 0; index -= 1)
        stack.push(node.children[index]!);
    }
    return undefined;
  }

  findAll(name: string, under: XmlElement = this.root): readonly XmlElement[] {
    const matches = matcher(name);
    const found: XmlElement[] = [];
    const stack: XmlElement[] = [under];
    while (stack.length > 0) {
      const node = stack.pop()!;
      if (matches(node)) found.push(node);
      for (let index = node.children.length - 1; index >= 0; index -= 1)
        stack.push(node.children[index]!);
    }
    return found;
  }

  at(path: readonly number[]): XmlElement | undefined {
    let node: XmlElement | undefined = this.root;
    for (const index of path) {
      node = node?.children[index];
      if (!node) return undefined;
    }
    return node;
  }

  textOf(node: XmlElement): string {
    if (node.selfClosing) return "";
    let out = "";
    let cursor = node.contentStart;
    const stop = node.contentEnd;
    const children = node.children;
    let child = 0;
    let opaqueIndex = firstOpaqueAt(this.#opaqueByStart, cursor);
    while (cursor < stop) {
      const nextChild = child < children.length ? children[child]! : undefined;
      const nextOpaque =
        opaqueIndex < this.#opaqueByStart.length
          ? this.#opaqueByStart[opaqueIndex]!
          : undefined;
      const childStart = nextChild ? nextChild.start : Number.POSITIVE_INFINITY;
      const opaqueStart =
        nextOpaque && nextOpaque.start < stop
          ? nextOpaque.start
          : Number.POSITIVE_INFINITY;
      const next = Math.min(childStart, opaqueStart, stop);
      if (next > cursor) out += decodeEntities(this.text.slice(cursor, next));
      if (next === stop) break;
      if (next === childStart) {
        out += this.textOf(nextChild!);
        cursor = nextChild!.end;
        child += 1;
        while (
          opaqueIndex < this.#opaqueByStart.length &&
          this.#opaqueByStart[opaqueIndex]!.start < cursor
        )
          opaqueIndex += 1;
      } else {
        if (nextOpaque!.kind === "cdata")
          out += this.text.slice(nextOpaque!.start + 9, nextOpaque!.end - 3);
        cursor = nextOpaque!.end;
        opaqueIndex += 1;
      }
    }
    return out;
  }

  attribute(node: XmlElement, name: string): string | undefined {
    return node.attributes.find((attribute) => attribute.name === name)?.value;
  }
}

function matcher(name: string): (node: XmlElement) => boolean {
  return name.includes(":")
    ? (node) => node.name === name
    : (node) => node.local === name;
}

function firstOpaqueAt(opaque: readonly XmlRange[], position: number): number {
  let low = 0;
  let high = opaque.length;
  while (low < high) {
    const middle = (low + high) >>> 1;
    if (opaque[middle]!.start < position) low = middle + 1;
    else high = middle;
  }
  return low;
}
