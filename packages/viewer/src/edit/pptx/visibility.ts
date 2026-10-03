import type { PptxElement } from "./types.js";

/**
 * Whether an element is drawn: neither it nor a group it sits in is hidden.
 * Hidden elements stay listed and editable by id, but a pointer never lands
 * on one, since nothing of it is on the slide.
 */
export function isDrawn(
  element: PptxElement,
  byId: ReadonlyMap<string, PptxElement>,
): boolean {
  for (
    let current: PptxElement | undefined = element;
    current;
    current =
      current.parentId === undefined ? undefined : byId.get(current.parentId)
  )
    if (current.hidden) return false;
  return true;
}
