export const MOBILE_LAYOUT =
  "(max-width: 760px), (max-width: 1024px) and (pointer: coarse)";

export function visibleViewport(window) {
  const visual = window.visualViewport;
  const zoomed = visual ? Math.abs(visual.scale - 1) > 0.01 : false;
  return {
    height:
      visual && !zoomed
        ? Math.min(window.innerHeight, visual.height)
        : window.innerHeight,
    top: visual && !zoomed ? Math.max(0, visual.offsetTop) : 0,
    zoomed,
  };
}

export function revealInScroller(field, scroller, viewport) {
  const bounds = scroller.getBoundingClientRect();
  const rect = field.getBoundingClientRect();
  const top = Math.max(bounds.top, viewport.top) + 8;
  const bottom = Math.min(bounds.bottom, viewport.top + viewport.height) - 8;
  if (bottom <= top) return;
  let distance = 0;
  if (rect.top < top || rect.height > bottom - top) distance = rect.top - top;
  else if (rect.bottom > bottom) distance = rect.bottom - bottom;
  if (distance) scroller.scrollTop += distance;
}
