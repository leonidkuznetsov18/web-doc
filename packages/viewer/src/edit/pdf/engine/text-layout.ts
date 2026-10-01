import type { PdfTextAlign } from "../types.js";

/*
 * Text layout for boxes web-doc draws: paragraphs on "\n", greedy word
 * wrapping, words longer than a line broken by character, one line per
 * text object. Pure, so it is tested without PDFium; the measurer supplies
 * widths.
 */

export interface LayoutInput {
  readonly text: string;
  readonly width: number;
  readonly height: number;
  readonly fontSize: number;
  /** Multiple of the font size. */
  readonly lineHeight: number;
  readonly align: PdfTextAlign;
  readonly ascent: number;
  /** Width of a string in points at the font size. */
  readonly advance: (text: string) => number;
}

export interface LayoutLine {
  readonly text: string;
  /** Offset of the line's left edge from the box's left edge. */
  readonly x: number;
  /** Offset of the baseline from the box's top edge. */
  readonly baseline: number;
  readonly width: number;
}

export interface Layout {
  readonly lines: readonly LayoutLine[];
  /** Height the lines need; may exceed the box. */
  readonly height: number;
  readonly overflow: boolean;
}

export function layoutText(input: LayoutInput): Layout {
  const step = input.fontSize * input.lineHeight;
  const texts: string[] = [];
  for (const paragraph of input.text.replaceAll("\r\n", "\n").split("\n"))
    texts.push(...wrapParagraph(paragraph.replaceAll("\t", " "), input));
  const lines = texts.map((text, index) => {
    const width = input.advance(text);
    return {
      text,
      x: offset(input.align, input.width, width),
      baseline: index * step + input.ascent,
      width,
    };
  });
  const height = lines.length * step;
  return { lines, height, overflow: height > input.height + 0.001 };
}

function wrapParagraph(paragraph: string, input: LayoutInput): string[] {
  const lines: string[] = [];
  let current = "";
  const push = (): void => {
    lines.push(current.trimEnd());
    current = "";
  };
  // Runs of spaces travel with the line they are in; a word that does not
  // fit moves to the next line, and one wider than the box is cut.
  for (const token of paragraph.match(/ +|[^ ]+/g) ?? []) {
    if (token.startsWith(" ")) {
      current += token;
      continue;
    }
    const candidate = current + token;
    if (input.advance(candidate) <= input.width) {
      current = candidate;
      continue;
    }
    if (current.trim()) push();
    else current = "";
    if (input.advance(token) <= input.width) {
      current = token;
      continue;
    }
    for (const character of token) {
      const next = current + character;
      if (current && input.advance(next) > input.width) {
        push();
        current = character;
      } else current = next;
    }
  }
  lines.push(current.trimEnd());
  return lines;
}

function offset(
  align: PdfTextAlign,
  boxWidth: number,
  lineWidth: number,
): number {
  switch (align) {
    case "center":
      return Math.max(0, (boxWidth - lineWidth) / 2);
    case "right":
      return Math.max(0, boxWidth - lineWidth);
    default:
      return 0;
  }
}
