import type { PagePoint, PageRect } from "../../types.js";

/**
 * How a page is displayed: its visible box in PDF user space (media box ∩
 * crop box, y up) and its /Rotate quarter turns. Page space is the displayed
 * page in points with the origin at its top-left corner and y down — the same
 * frame PDF.js uses for its viewport at scale 1.
 */
export interface PageGeometry {
  /** User-space box: left and bottom are the minimum x and y. */
  readonly box: {
    readonly left: number;
    readonly bottom: number;
    readonly right: number;
    readonly top: number;
  };
  /** Quarter turns clockwise, 0–3. */
  readonly rotation: number;
}

/** Displayed size in points, rotation applied. */
export function displayedSize(geometry: PageGeometry): {
  readonly width: number;
  readonly height: number;
} {
  const width = geometry.box.right - geometry.box.left;
  const height = geometry.box.top - geometry.box.bottom;
  return geometry.rotation % 2 === 0
    ? { width, height }
    : { width: height, height: width };
}

export function userToPage(
  geometry: PageGeometry,
  x: number,
  y: number,
): PagePoint {
  const width = geometry.box.right - geometry.box.left;
  const height = geometry.box.top - geometry.box.bottom;
  // Unrotated page space: flip y so it grows downwards from the top edge.
  const px = x - geometry.box.left;
  const py = geometry.box.top - y;
  switch (geometry.rotation) {
    case 1:
      return { x: height - py, y: px };
    case 2:
      return { x: width - px, y: height - py };
    case 3:
      return { x: py, y: width - px };
    default:
      return { x: px, y: py };
  }
}

export function pageToUser(
  geometry: PageGeometry,
  x: number,
  y: number,
): { readonly x: number; readonly y: number } {
  const width = geometry.box.right - geometry.box.left;
  const height = geometry.box.top - geometry.box.bottom;
  let px: number;
  let py: number;
  switch (geometry.rotation) {
    case 1:
      px = y;
      py = height - x;
      break;
    case 2:
      px = width - x;
      py = height - y;
      break;
    case 3:
      px = width - y;
      py = x;
      break;
    default:
      px = x;
      py = y;
  }
  return { x: px + geometry.box.left, y: geometry.box.top - py };
}

/** Axis-aligned page-space box of a user-space rectangle. */
export function userRectToPage(
  geometry: PageGeometry,
  left: number,
  bottom: number,
  right: number,
  top: number,
): PageRect {
  const corners = [
    userToPage(geometry, left, bottom),
    userToPage(geometry, right, bottom),
    userToPage(geometry, right, top),
    userToPage(geometry, left, top),
  ];
  const xs = corners.map((corner) => corner.x);
  const ys = corners.map((corner) => corner.y);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, width: Math.max(...xs) - x, height: Math.max(...ys) - y };
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

export function rectContains(rect: PageRect, point: PagePoint): boolean {
  return (
    point.x >= rect.x &&
    point.x <= rect.x + rect.width &&
    point.y >= rect.y &&
    point.y <= rect.y + rect.height
  );
}

export function rectsIntersect(a: PageRect, b: PageRect): boolean {
  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

/** Rounds geometry to a thousandth of a point so results are tidy and stable. */
export function roundRect(rect: PageRect): PageRect {
  return {
    x: round(rect.x),
    y: round(rect.y),
    width: round(rect.width),
    height: round(rect.height),
  };
}

export function round(value: number): number {
  return Math.round(value * 1000) / 1000;
}
