import type { TextRun } from "../../contracts.js";
import type {
  PagePoint,
  PageRect,
  TextLayoutGlyph,
  TextLayoutLine,
  TextPosition,
} from "../types.js";

/*
 * The lines of a DOCX paragraph from the renderer's text runs. A run is a
 * piece of one line: `y` and `height` are the line's box (its top and the
 * line pitch), `x` and `width` the pen origin and the advance it fills,
 * `advanceBounds`, when the renderer reports it, the ascent-to-descent box
 * around the baseline. Runs of a line share their top, so consecutive runs
 * at the same top form a line. Run text is aligned with the paragraph's
 * logical text, so every glyph names its offset; a glyph the text does not
 * hold (list numbering, a tab drawn as nothing) is left out. Advances come
 * from the run's font measured on a canvas and scaled to the run's width,
 * or are even shares of the width where nothing can measure.
 */

/** Measures text the way the renderer's canvas does; absent outside a browser. */
export interface TextMeasurer {
  /** Advance of each code point of `text` in a CSS `font`, in CSS pixels. */
  advances(font: string, text: string): readonly number[];
  /** Ascent and descent of a CSS `font`, in CSS pixels. */
  metrics(font: string): { readonly ascent: number; readonly descent: number };
}

/** A run with the page it was laid out on. */
export interface PlacedRun {
  readonly run: TextRun;
  readonly pageIndex: number;
}

/** A laid-out line without its colour, which the engine resolves. */
export interface LaidLine extends Omit<TextLayoutLine, "color"> {
  readonly pageIndex: number;
  /** The line's box: from its top to its bottom by the line pitch. */
  readonly box: PageRect;
  /** Per glyph, the offset after its character (a surrogate pair counts two). */
  readonly ends: readonly number[];
}

/** Share of the font size above the baseline when no font can be measured. */
const ASCENT_SHARE = 0.8;
/** CSS pixels per point. */
const PX_PER_PT = 96 / 72;
/** Runs whose tops differ by less than this many pixels share a line. */
const SAME_LINE = 0.5;
/** Characters the renderer may draw as nothing or as something else. */
const UNDRAWN = /[\s\ufffc\u00ad\u200b-\u200d\ufeff]/;

/** A canvas measurer when the environment has one, else undefined. */
export function canvasMeasurer(): TextMeasurer | undefined {
  const context = canvasContext();
  if (!context) return undefined;
  const metrics = new Map<string, { ascent: number; descent: number }>();
  return {
    advances(font, text) {
      context.font = font;
      const out: number[] = [];
      let before = 0;
      let prefix = "";
      for (const character of text) {
        prefix += character;
        const width = context.measureText(prefix).width;
        out.push(width - before);
        before = width;
      }
      return out;
    },
    metrics(font) {
      let known = metrics.get(font);
      if (!known) {
        context.font = font;
        const measured = context.measureText("Hg");
        known = {
          ascent:
            measured.fontBoundingBoxAscent ?? measured.actualBoundingBoxAscent,
          descent:
            measured.fontBoundingBoxDescent ??
            measured.actualBoundingBoxDescent,
        };
        metrics.set(font, known);
      }
      return known;
    },
  };
}

type MeasureContext = Pick<CanvasRenderingContext2D, "font" | "measureText">;

function canvasContext(): MeasureContext | undefined {
  try {
    if (typeof OffscreenCanvas === "function")
      return new OffscreenCanvas(1, 1).getContext("2d") ?? undefined;
    if (typeof document !== "undefined")
      return document.createElement("canvas").getContext("2d") ?? undefined;
  } catch {
    // A canvas that cannot be created measures nothing; shares are used.
  }
  return undefined;
}

/**
 * The lines a paragraph draws, across the pages of `runs`, in order.
 * `fallback.fontSize` (points) sizes runs that carry no size of their own.
 */
export function layParagraph(
  elementId: string,
  text: string,
  runs: readonly PlacedRun[],
  fallback: { readonly fontFamily: string; readonly fontSize: number },
  measurer: TextMeasurer | undefined,
): LaidLine[] {
  const lines: LaidLine[] = [];
  let cursor = 0;
  let group: { run: TextRun; offsets: (number | undefined)[] }[] = [];
  let groupPage = -1;
  const flush = (): void => {
    const line = layLine(elementId, text, groupPage, group, fallback, measurer);
    if (line) lines.push(line);
    group = [];
  };
  for (const { run, pageIndex } of runs) {
    const previous = group.at(-1)?.run;
    if (
      previous &&
      (pageIndex !== groupPage || Math.abs(run.y - previous.y) >= SAME_LINE)
    )
      flush();
    groupPage = pageIndex;
    const aligned = alignRun(text, cursor, run.text);
    cursor = aligned.cursor;
    group.push({ run, offsets: aligned.offsets });
  }
  if (group.length > 0) flush();
  return lines;
}

/**
 * The offset of each code point of a run's text in the paragraph's text,
 * from `cursor` on: the whole run where the text continues with it (after
 * characters the renderer may skip), else character by character, case
 * folded. A run whose visible characters the text does not all hold next
 * (list numbering, a bullet) gets no offsets and leaves the cursor; only
 * its whitespace may go unmatched.
 */
function alignRun(
  text: string,
  cursor: number,
  runText: string,
): { offsets: (number | undefined)[]; cursor: number } {
  const characters = [...runText];
  const whole = text.indexOf(runText, cursor);
  if (whole >= 0 && skippable(text.slice(cursor, whole))) {
    let at = whole;
    const offsets = characters.map((character) => {
      const offset = at;
      at += character.length;
      return offset;
    });
    return { offsets, cursor: at };
  }
  const offsets: (number | undefined)[] = [];
  let at = cursor;
  for (const character of characters) {
    let probe = at;
    while (
      probe < text.length &&
      !sameCharacter(text, probe, character) &&
      UNDRAWN.test(text[probe]!)
    )
      probe += 1;
    if (probe < text.length && sameCharacter(text, probe, character)) {
      offsets.push(probe);
      at = probe + character.length;
    } else if (UNDRAWN.test(character)) offsets.push(undefined);
    else return { offsets: characters.map(() => undefined), cursor };
  }
  return { offsets, cursor: at };
}

function skippable(between: string): boolean {
  return [...between].every((character) => UNDRAWN.test(character));
}

/** Case folded, so text drawn in capitals (`w:caps`) still aligns. */
function sameCharacter(text: string, at: number, character: string): boolean {
  const held = text.slice(at, at + character.length);
  return (
    held === character || held.toLocaleLowerCase() === character.toLowerCase()
  );
}

function layLine(
  elementId: string,
  text: string,
  pageIndex: number,
  group: readonly { run: TextRun; offsets: readonly (number | undefined)[] }[],
  fallback: { readonly fontFamily: string; readonly fontSize: number },
  measurer: TextMeasurer | undefined,
): LaidLine | undefined {
  const runs = group.map(({ run }) => ({
    run,
    ...vertical(run, fallback.fontSize, measurer),
  }));
  // Only runs that draw the paragraph's text size the line: numbering
  // left out of the glyphs is left out of the metrics too. The pen line is
  // the one of the largest text; a raised or lowered run keeps its own
  // ascent box.
  const drawing = runs.filter((_, index) =>
    group[index]!.offsets.some((offset) => offset !== undefined),
  );
  if (drawing.length === 0) return undefined;
  const main = drawing.reduce((best, entry) =>
    entry.fontSize > best.fontSize ? entry : best,
  );
  const ascent = Math.max(...drawing.map((entry) => entry.ascent));
  const descent = Math.max(...drawing.map((entry) => entry.descent));
  const glyphs: TextLayoutGlyph[] = [];
  const ends: number[] = [];
  for (const [index, { run, offsets }] of group.entries()) {
    const entry = runs[index]!;
    const advances = advancesOf(run, measurer);
    let x = run.x;
    for (const [at, character] of [...run.text].entries()) {
      const advance = advances[at] ?? 0;
      const offset = offsets[at];
      if (offset !== undefined) {
        glyphs.push({
          offset,
          box: {
            x,
            y: entry.baseline - entry.ascent,
            width: advance,
            height: entry.ascent + entry.descent,
          },
          advance,
          origin: { x, y: entry.baseline },
        });
        ends.push(offset + character.length);
      }
      x += advance;
    }
  }
  const first = glyphs[0];
  const last = glyphs.at(-1);
  if (!first || !last) return undefined;
  const start = first.offset;
  const end = Math.max(...ends);
  const left = first.origin!.x;
  const advanceBounds = {
    x: left,
    y: main.baseline - ascent,
    width: last.origin!.x + last.advance - left,
    height: ascent + descent,
  };
  const top = Math.min(...group.map(({ run }) => run.y));
  const bottom = Math.max(...group.map(({ run }) => run.y + run.height));
  return {
    pageIndex,
    range: {
      start: { elementId, offset: start },
      end: { elementId, offset: end },
    },
    text: text.slice(start, end),
    bounds: unionRects(glyphs.map((glyph) => glyph.box)),
    advanceBounds,
    baseline: { x: left, y: main.baseline },
    glyphs,
    fontFamily: main.run.fontFamily ?? fallback.fontFamily,
    fontSize: main.fontSize,
    box: {
      x: left,
      y: top,
      width: advanceBounds.width,
      height: bottom - top,
    },
    ends,
  };
}

/**
 * Where a run's baseline lies and how far its font reaches above and below
 * it, in CSS pixels. With the renderer's ascent box, the baseline divides
 * it by the font's measured ascent share; without one, the run sits in its
 * line box as CSS lays out a line of that height.
 */
function vertical(
  run: TextRun,
  fallbackPoints: number,
  measurer: TextMeasurer | undefined,
): { baseline: number; ascent: number; descent: number; fontSize: number } {
  const fontSize = run.fontSize ?? fallbackPoints * PX_PER_PT;
  const measured = run.font ? measurer?.metrics(run.font) : undefined;
  const ascent = measured?.ascent ?? fontSize * ASCENT_SHARE;
  const descent = measured?.descent ?? fontSize * (1 - ASCENT_SHARE);
  const share = ascent + descent > 0 ? ascent / (ascent + descent) : 1;
  const box = run.advanceBounds;
  if (box && box.height > 0 && (box.y !== run.y || box.height !== run.height))
    return {
      baseline: box.y + box.height * share,
      ascent: box.height * share,
      descent: box.height * (1 - share),
      fontSize,
    };
  return {
    baseline: run.y + (run.height - ascent - descent) / 2 + ascent,
    ascent,
    descent,
    fontSize,
  };
}

/** Each code point's advance, scaled so the run's advances fill its width. */
function advancesOf(
  run: TextRun,
  measurer: TextMeasurer | undefined,
): readonly number[] {
  const count = [...run.text].length;
  if (count === 0) return [];
  const spacing = run.letterSpacingPx ?? 0;
  const measured =
    run.font && measurer
      ? measurer.advances(run.font, run.text).map((width) => width + spacing)
      : undefined;
  const raw = measured ?? new Array<number>(count).fill(1);
  const total = raw.reduce((sum, width) => sum + width, 0);
  if (total <= 0) return new Array<number>(count).fill(run.width / count);
  return raw.map((width) => (width * run.width) / total);
}

/**
 * The caret nearest to `point` among `lines`: the line whose box is
 * nearest (vertically first, then horizontally), and in it the edge of the
 * glyph under the point, the one after it past the glyph's middle.
 */
export function caretAt(
  lines: readonly LaidLine[],
  point: PagePoint,
): TextPosition | undefined {
  let best: LaidLine | undefined;
  let bestDistance: readonly [number, number] = [Infinity, Infinity];
  for (const line of lines) {
    const distance = [
      gap(point.y, line.box.y, line.box.y + line.box.height),
      gap(point.x, line.box.x, line.box.x + line.box.width),
    ] as const;
    if (
      distance[0] < bestDistance[0] ||
      (distance[0] === bestDistance[0] && distance[1] < bestDistance[1])
    ) {
      best = line;
      bestDistance = distance;
    }
  }
  if (!best) return undefined;
  const elementId = best.range.start.elementId;
  for (const [index, glyph] of best.glyphs.entries()) {
    const middle = glyph.origin!.x + glyph.advance / 2;
    if (point.x < middle) return { elementId, offset: glyph.offset };
    if (point.x < glyph.origin!.x + glyph.advance)
      return { elementId, offset: best.ends[index]! };
  }
  return { elementId, offset: best.ends.at(-1)! };
}

function gap(value: number, from: number, to: number): number {
  if (value < from) return from - value;
  if (value > to) return value - to;
  return 0;
}

export function unionRects(rects: readonly PageRect[]): PageRect {
  const x = Math.min(...rects.map((rect) => rect.x));
  const y = Math.min(...rects.map((rect) => rect.y));
  return {
    x,
    y,
    width: Math.max(...rects.map((rect) => rect.x + rect.width)) - x,
    height: Math.max(...rects.map((rect) => rect.y + rect.height)) - y,
  };
}
