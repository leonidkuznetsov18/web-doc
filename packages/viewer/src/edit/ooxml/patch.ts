import { ViewerError } from "../../errors.js";
import {
  escapeAttribute,
  escapeText,
  scanXml,
  type XmlElement,
  type XmlPart,
} from "./xml.js";

/*
 * Range patches on a scanned part. A patch replaces a half-open range of
 * the part's text and says what the patched position must read back as
 * after a full re-scan: an element whose outer XML equals the fragment, an
 * element's content, an attribute's value, or nothing at all. Patches are
 * applied from the end so earlier ranges stay valid, then the part is
 * re-scanned and every expectation checked; any failure discards the new
 * text. The builders produce patches with the right ranges and
 * expectations for the common shapes.
 */

export type PatchExpectation =
  | { readonly kind: "element" }
  /** The element starting at `at` (in the unpatched text) has this content afterwards. */
  | { readonly kind: "content"; readonly at: number }
  | {
      readonly kind: "attribute";
      readonly at: number;
      readonly name: string;
      readonly value: string | undefined;
    }
  | {
      readonly kind: "removed";
      /** Where the parent starts, how many element children it has without the element, and the element's own text. */
      readonly parentAt: number;
      readonly count: number;
      readonly xml: string;
    }
  | { readonly kind: "none" };

export interface XmlPatch {
  /** Half-open range in the part's text; patches of one part must not overlap. */
  readonly start: number;
  readonly end: number;
  readonly text: string;
  readonly expect: PatchExpectation;
}

export const patches = {
  /** Replaces the content between the tags; a self-closing element is opened. */
  replaceContent(part: XmlPart, node: XmlElement, xml: string): XmlPatch {
    if (node.selfClosing)
      return {
        start: node.end - 2,
        end: node.end,
        text: `>${xml}</${node.name}>`,
        expect: { kind: "content", at: node.start },
      };
    return {
      start: node.contentStart,
      end: node.contentEnd,
      text: xml,
      expect: { kind: "content", at: node.start },
    };
  },
  /** Replaces the whole element with one element. */
  replaceElement(part: XmlPart, node: XmlElement, xml: string): XmlPatch {
    return {
      start: node.start,
      end: node.end,
      text: xml,
      expect: { kind: "element" },
    };
  },
  /** Removes the element and nothing around it. */
  removeElement(part: XmlPart, node: XmlElement): XmlPatch {
    if (!node.parent)
      throw new ViewerError(
        "invalid-patch",
        `Part ${part.name}: the root element cannot be removed`,
        { details: { part: part.name } },
      );
    return {
      start: node.start,
      end: node.end,
      text: "",
      expect: {
        kind: "removed",
        parentAt: node.parent.start,
        count: node.parent.children.length - 1,
        xml: part.text.slice(node.start, node.end),
      },
    };
  },
  insertBefore(part: XmlPart, node: XmlElement, xml: string): XmlPatch {
    return {
      start: node.start,
      end: node.start,
      text: xml,
      expect: { kind: "element" },
    };
  },
  insertAfter(part: XmlPart, node: XmlElement, xml: string): XmlPatch {
    return {
      start: node.end,
      end: node.end,
      text: xml,
      expect: { kind: "element" },
    };
  },
  /** Appends one element at the end of the content; a self-closing element is opened. */
  appendChild(part: XmlPart, node: XmlElement, xml: string): XmlPatch {
    if (node.selfClosing)
      return {
        start: node.end - 2,
        end: node.end,
        text: `>${xml}</${node.name}>`,
        expect: { kind: "content", at: node.start },
      };
    return {
      start: node.contentEnd,
      end: node.contentEnd,
      text: xml,
      expect: { kind: "element" },
    };
  },
  setAttribute(
    part: XmlPart,
    node: XmlElement,
    name: string,
    value: string,
  ): XmlPatch {
    const existing = node.attributes.find(
      (attribute) => attribute.name === name,
    );
    const expect: PatchExpectation = {
      kind: "attribute",
      at: node.start,
      name,
      value,
    };
    if (existing)
      return {
        start: existing.start,
        end: existing.end,
        text: `${name}=${escapeAttribute(value)}`,
        expect,
      };
    const at = startTagClose(part, node);
    return {
      start: at,
      end: at,
      text: ` ${name}=${escapeAttribute(value)}`,
      expect,
    };
  },
  removeAttribute(part: XmlPart, node: XmlElement, name: string): XmlPatch {
    const existing = node.attributes.find(
      (attribute) => attribute.name === name,
    );
    const expect: PatchExpectation = {
      kind: "attribute",
      at: node.start,
      name,
      value: undefined,
    };
    if (!existing)
      return { start: node.start, end: node.start, text: "", expect };
    // Take the whitespace before the attribute with it.
    let start = existing.start;
    while (start > node.start && /\s/.test(part.text[start - 1]!)) start -= 1;
    return { start, end: existing.end, text: "", expect };
  },
  text: escapeText,
  attr: escapeAttribute,
};

/** Where the start tag's ">" or "/>" begins. */
function startTagClose(part: XmlPart, node: XmlElement): number {
  if (node.selfClosing) return node.end - 2;
  const last = node.attributes.at(-1);
  const from = last ? last.end : node.start + 1 + node.name.length;
  const close = part.text.indexOf(">", from);
  if (close < 0 || close >= node.contentStart)
    throw new ViewerError("internal", `Start tag of ${node.name} has no close`);
  return close;
}

/**
 * Applies the patches to the part's text, re-scans it and checks every
 * expectation; returns the new part scanned at `revision`. Throws
 * `invalid-patch` and leaves nothing behind when anything is wrong.
 */
export function applyPatches(
  part: XmlPart,
  items: readonly XmlPatch[],
  revision: number,
): XmlPart {
  const sorted = [...items].sort((a, b) => a.start - b.start || a.end - b.end);
  let cursor = 0;
  for (const patch of sorted) {
    if (
      !Number.isInteger(patch.start) ||
      !Number.isInteger(patch.end) ||
      patch.start < 0 ||
      patch.end < patch.start ||
      patch.end > part.text.length
    )
      throw invalid(part, patch, "the range is outside the part");
    if (patch.start < cursor) throw invalid(part, patch, "the ranges overlap");
    cursor = patch.end;
  }
  // Build the new text from the front, remembering where each patch lands.
  let out = "";
  let from = 0;
  const landed: number[] = [];
  for (const patch of sorted) {
    out += part.text.slice(from, patch.start);
    landed.push(out.length);
    out += patch.text;
    from = patch.end;
  }
  out += part.text.slice(from);
  let rescanned: XmlPart;
  try {
    rescanned = scanXml(part.name, out, revision);
  } catch (error) {
    throw new ViewerError(
      "invalid-patch",
      `Part ${part.name}: the patched part is not well-formed`,
      {
        cause: error,
        details: {
          part: part.name,
          reason: error instanceof Error ? error.message : String(error),
        },
      },
    );
  }
  const byStart = new Map<number, XmlElement>();
  const stack: XmlElement[] = [rescanned.root];
  while (stack.length > 0) {
    const node = stack.pop()!;
    byStart.set(node.start, node);
    for (const child of node.children) stack.push(child);
  }
  /** Where an unpatched offset lands: shifted by every patch that ends at or before it. */
  const mapped = (offset: number): number => {
    let delta = 0;
    for (const patch of sorted) {
      if (patch.end > offset) break;
      delta += patch.text.length - (patch.end - patch.start);
    }
    return offset + delta;
  };
  sorted.forEach((patch, index) => {
    const at = landed[index]!;
    const { expect } = patch;
    switch (expect.kind) {
      case "element": {
        const node = byStart.get(at);
        if (!node || node.end !== at + patch.text.length)
          throw invalid(
            part,
            patch,
            "the fragment does not read back as one element",
            rescanned.text.slice(at, at + patch.text.length),
          );
        break;
      }
      case "content": {
        const node = byStart.get(mapped(expect.at));
        if (!node) throw invalid(part, patch, "the patched element is gone");
        const content = rescanned.text.slice(
          node.contentStart,
          node.contentEnd,
        );
        const wanted = node.selfClosing
          ? ""
          : patch.text.startsWith(">") && patch.text.endsWith(`</${node.name}>`)
            ? patch.text.slice(1, -(node.name.length + 3))
            : patch.text;
        if (content !== wanted)
          throw invalid(part, patch, "the content does not read back", content);
        break;
      }
      case "attribute": {
        const node = byStart.get(mapped(expect.at));
        if (!node) throw invalid(part, patch, "the patched element is gone");
        const read = rescanned.attribute(node, expect.name);
        if (read !== expect.value)
          throw invalid(
            part,
            patch,
            `attribute ${expect.name} reads back differently`,
            read,
          );
        break;
      }
      case "removed": {
        const parent = byStart.get(mapped(expect.parentAt));
        if (!parent)
          throw invalid(part, patch, "the removed element's parent is gone");
        // Another patch may have put a sibling in the same place; then the
        // removed element's own text must at least be gone from there.
        if (
          parent.children.length !== expect.count &&
          rescanned.text.startsWith(expect.xml, at)
        )
          throw invalid(
            part,
            patch,
            "the element is still there",
            String(parent.children.length),
          );
        break;
      }
      case "none":
        break;
    }
  });
  return rescanned;
}

function invalid(
  part: XmlPart,
  patch: XmlPatch,
  reason: string,
  read?: string,
): ViewerError {
  return new ViewerError("invalid-patch", `Part ${part.name}: ${reason}`, {
    details: {
      part: part.name,
      range: [patch.start, patch.end],
      reason,
      ...(read === undefined ? {} : { read: read.slice(0, 200) }),
    },
  });
}
