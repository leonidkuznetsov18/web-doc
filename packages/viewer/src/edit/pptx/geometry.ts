import type { PageRect } from "../types.js";
import type { ElementFrame } from "../types.js";
import type { XmlElement, XmlPart } from "../ooxml/xml.js";

/*
 * DrawingML geometry: EMU, 60,000ths of a degree, and the affine transforms
 * of groups. Slide space is CSS pixels at 96 dpi.
 */

export const EMU_PER_PX = 9525;
export const EMU_PER_PT = 12700;

export function emuToPx(emu: number): number {
  return emu / EMU_PER_PX;
}

export function pxToEmu(px: number): number {
  return Math.round(px * EMU_PER_PX);
}

/** An `a:xfrm` as the file stores it, in EMU and degrees. */
export interface Xfrm {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** Clockwise degrees. */
  readonly rotation: number;
  readonly flipH: boolean;
  readonly flipV: boolean;
  /** Child space of a group. */
  readonly child?: {
    readonly x: number;
    readonly y: number;
    readonly cx: number;
    readonly cy: number;
  };
}

/** Reads `a:xfrm` (or `p:xfrm` of a graphic frame); undefined without an offset and an extent. */
export function readXfrm(
  part: XmlPart,
  node: XmlElement | undefined,
): Xfrm | undefined {
  if (!node) return undefined;
  const off = node.children.find((child) => child.local === "off");
  const ext = node.children.find((child) => child.local === "ext");
  if (!off || !ext) return undefined;
  const x = integer(part.attribute(off, "x"));
  const y = integer(part.attribute(off, "y"));
  const cx = integer(part.attribute(ext, "cx"));
  const cy = integer(part.attribute(ext, "cy"));
  if ([x, y, cx, cy].some((value) => value === undefined)) return undefined;
  const chOff = node.children.find((child) => child.local === "chOff");
  const chExt = node.children.find((child) => child.local === "chExt");
  const child =
    chOff && chExt
      ? {
          x: integer(part.attribute(chOff, "x")) ?? 0,
          y: integer(part.attribute(chOff, "y")) ?? 0,
          cx: integer(part.attribute(chExt, "cx")) ?? 0,
          cy: integer(part.attribute(chExt, "cy")) ?? 0,
        }
      : undefined;
  return {
    x: x!,
    y: y!,
    cx: cx!,
    cy: cy!,
    rotation: (integer(part.attribute(node, "rot")) ?? 0) / 60000,
    flipH: part.attribute(node, "flipH") === "1",
    flipV: part.attribute(node, "flipV") === "1",
    ...(child ? { child } : {}),
  };
}

function integer(value: string | undefined): number | undefined {
  if (value === undefined) return undefined;
  const number = Number(value);
  return Number.isFinite(number) ? Math.trunc(number) : undefined;
}

/** A 2D affine transform: x' = a·x + c·y + e, y' = b·x + d·y + f. */
export interface Matrix {
  readonly a: number;
  readonly b: number;
  readonly c: number;
  readonly d: number;
  readonly e: number;
  readonly f: number;
}

export const IDENTITY: Matrix = { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };

export function multiply(m: Matrix, n: Matrix): Matrix {
  // m ∘ n: apply n first, then m.
  return {
    a: m.a * n.a + m.c * n.b,
    b: m.b * n.a + m.d * n.b,
    c: m.a * n.c + m.c * n.d,
    d: m.b * n.c + m.d * n.d,
    e: m.a * n.e + m.c * n.f + m.e,
    f: m.b * n.e + m.d * n.f + m.f,
  };
}

export function translate(tx: number, ty: number): Matrix {
  return { a: 1, b: 0, c: 0, d: 1, e: tx, f: ty };
}

export function scale(sx: number, sy: number): Matrix {
  return { a: sx, b: 0, c: 0, d: sy, e: 0, f: 0 };
}

function rotate(degrees: number): Matrix {
  const radians = (degrees * Math.PI) / 180;
  const cos = Math.cos(radians);
  const sin = Math.sin(radians);
  return { a: cos, b: sin, c: -sin, d: cos, e: 0, f: 0 };
}

export function apply(
  m: Matrix,
  point: { readonly x: number; readonly y: number },
): { x: number; y: number } {
  return {
    x: m.a * point.x + m.c * point.y + m.e,
    y: m.b * point.x + m.d * point.y + m.f,
  };
}

export function invert(m: Matrix): Matrix {
  const det = m.a * m.d - m.b * m.c;
  if (det === 0) return IDENTITY;
  return {
    a: m.d / det,
    b: -m.b / det,
    c: -m.c / det,
    d: m.a / det,
    e: (m.c * m.f - m.d * m.e) / det,
    f: (m.b * m.e - m.a * m.f) / det,
  };
}

/** Rotation and flips of a frame about its centre, in the frame's own space. */
function frameTransform(frame: Xfrm): Matrix {
  const centre = translate(frame.x + frame.cx / 2, frame.y + frame.cy / 2);
  const back = translate(-(frame.x + frame.cx / 2), -(frame.y + frame.cy / 2));
  const flip = scale(frame.flipH ? -1 : 1, frame.flipV ? -1 : 1);
  return multiply(
    centre,
    multiply(rotate(frame.rotation), multiply(flip, back)),
  );
}

/**
 * The transform from a group's child space to its parent space: child
 * coordinates are scaled from `chOff`/`chExt` into `off`/`ext`, then the
 * group's rotation and flips apply about its centre.
 */
export function groupMatrix(group: Xfrm): Matrix {
  const child = group.child ?? { x: 0, y: 0, cx: group.cx, cy: group.cy };
  const sx = child.cx === 0 ? 1 : group.cx / child.cx;
  const sy = child.cy === 0 ? 1 : group.cy / child.cy;
  const place = multiply(
    translate(group.x, group.y),
    multiply(scale(sx, sy), translate(-child.x, -child.y)),
  );
  return multiply(frameTransform(group), place);
}

/** A frame in slide space (px) with its axis-aligned bounds. */
export interface PlacedFrame {
  readonly frame: ElementFrame;
  readonly bounds: PageRect;
  /** Exact local-to-slide EMU mapping, including nonuniform group transforms. */
  readonly matrix: Matrix;
  readonly sourceFrame: Xfrm;
}

/**
 * Places a frame given in its own space (EMU) through the transform of the
 * groups above it. The reported frame keeps the element's box scaled and
 * positioned by the groups, with the rotations and flips combined; the
 * bounds are the axis-aligned box of its four transformed corners.
 */
export function placeFrame(frame: Xfrm, parents: Matrix): PlacedFrame {
  const own = frameTransform(frame);
  const total = multiply(parents, own);
  const corners = [
    { x: frame.x, y: frame.y },
    { x: frame.x + frame.cx, y: frame.y },
    { x: frame.x + frame.cx, y: frame.y + frame.cy },
    { x: frame.x, y: frame.y + frame.cy },
  ].map((corner) => apply(total, corner));
  const xs = corners.map((corner) => corner.x);
  const ys = corners.map((corner) => corner.y);
  const left = Math.min(...xs);
  const top = Math.min(...ys);
  const bounds: PageRect = {
    x: emuToPx(left),
    y: emuToPx(top),
    width: emuToPx(Math.max(...xs) - left),
    height: emuToPx(Math.max(...ys) - top),
  };
  // The unrotated box in slide space: the parents' transform applied to the
  // element's box, measured by its centre and its scaled extent.
  const centre = apply(parents, {
    x: frame.x + frame.cx / 2,
    y: frame.y + frame.cy / 2,
  });
  const scaleX = Math.hypot(parents.a, parents.b);
  const scaleY = Math.hypot(parents.c, parents.d);
  const width = frame.cx * scaleX;
  const height = frame.cy * scaleY;
  const parentFlip = parents.a * parents.d - parents.b * parents.c < 0;
  // A reflecting parent is reported as a rotation followed by flipH, so its
  // angle comes from the matrix with that flip taken out.
  const parentRotation = parentFlip
    ? (Math.atan2(-parents.b, -parents.a) * 180) / Math.PI
    : (Math.atan2(parents.b, parents.a) * 180) / Math.PI;
  const rotation = normalizeDegrees(
    parentRotation + (parentFlip ? -frame.rotation : frame.rotation),
  );
  return {
    matrix: total,
    sourceFrame: frame,
    frame: {
      x: emuToPx(centre.x - width / 2),
      y: emuToPx(centre.y - height / 2),
      width: emuToPx(width),
      height: emuToPx(height),
      rotation,
      flipH: frame.flipH !== parentFlip,
      flipV: frame.flipV,
    },
    bounds,
  };
}

function normalizeDegrees(degrees: number): number {
  const value = ((degrees % 360) + 360) % 360;
  return Math.abs(value) < 1e-9 ? 0 : Number(value.toFixed(6));
}

/** Whether a slide-space point lies inside a placed frame, rotation and flips included. */
export function frameContains(
  placed: PlacedFrame,
  point: { readonly x: number; readonly y: number },
): boolean {
  const { matrix, sourceFrame } = placed;
  const det = matrix.a * matrix.d - matrix.b * matrix.c;
  if (!Number.isFinite(det) || det === 0) return false;
  const local = apply(invert(matrix), {
    x: point.x * EMU_PER_PX,
    y: point.y * EMU_PER_PX,
  });
  if (!Number.isFinite(local.x) || !Number.isFinite(local.y)) return false;
  const tolerance = EMU_PER_PX * 1e-6;
  return (
    local.x >= sourceFrame.x - tolerance &&
    local.x <= sourceFrame.x + sourceFrame.cx + tolerance &&
    local.y >= sourceFrame.y - tolerance &&
    local.y <= sourceFrame.y + sourceFrame.cy + tolerance
  );
}

export function rectsIntersect(a: PageRect, b: PageRect): boolean {
  return (
    a.x < b.x + b.width &&
    b.x < a.x + a.width &&
    a.y < b.y + b.height &&
    b.y < a.y + a.height
  );
}
