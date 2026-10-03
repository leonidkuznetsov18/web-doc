import { ViewerError } from "../errors.js";
import type { TextRange } from "./types.js";

/** Half-open UTF-16 offsets in an element's text. */
export interface TextSpan {
  readonly start: number;
  readonly end: number;
}

/**
 * The span a range names on its target, or `undefined` for no range (the
 * whole text). Both ends must lie on the target.
 */
export function spanOnTarget(
  target: string,
  range: TextRange | undefined,
): TextSpan | undefined {
  if (!range) return undefined;
  if (range.start.elementId !== target || range.end.elementId !== target)
    throw new ViewerError(
      "invalid-operation",
      "Both ends of the range must lie on the target",
      { details: { target } },
    );
  return { start: range.start.offset, end: range.end.offset };
}

/**
 * The runs a span covers. A collapsed span reads the run before it, whose
 * style text typed there takes, or the first run after it at the very start.
 */
export function runsCovering<R extends TextSpan>(
  runs: readonly R[],
  span: TextSpan,
): R[] {
  if (span.start < span.end)
    return runs.filter((run) => run.start < span.end && run.end > span.start);
  const before = runs
    .filter((run) => run.start < span.start && run.end >= span.start)
    .at(-1);
  const at = before ?? runs.find((run) => run.start >= span.start);
  return at ? [at] : [];
}

/**
 * What every style has in common: a property two of them differ on, or one
 * some leave out, is left out. Values compare as JSON, so a theme colour and
 * its plain form differ.
 */
export function sharedStyle<T extends object>(
  styles: readonly T[],
): Partial<T> | undefined {
  const [first, ...rest] = styles;
  if (!first) return undefined;
  const shared: Partial<T> = {};
  for (const key of Object.keys(first) as (keyof T)[]) {
    const value = JSON.stringify(first[key]);
    if (
      rest.every(
        (style) => key in style && JSON.stringify(style[key]) === value,
      )
    )
      shared[key] = first[key];
  }
  return shared;
}

/** Whether a span lies inside a text of `length` code units. */
export function spanFits(span: TextSpan, length: number): boolean {
  return span.start >= 0 && span.start <= span.end && span.end <= length;
}
