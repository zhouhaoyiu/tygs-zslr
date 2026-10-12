import test from "node:test";
import assert from "node:assert/strict";
import { revealInScroller, visibleViewport } from "../src/viewport.js";

test("a keyboard can reduce only the visible area and pan it down", () => {
  const viewport = visibleViewport({
    innerHeight: 844,
    visualViewport: { height: 524, offsetTop: 90, scale: 1 },
  });
  assert.equal(viewport.height, 524);
  assert.equal(viewport.top, 90);
  assert.equal(viewport.zoomed, false);
});

test("a resized layout also bounds a temporarily stale visual viewport", () => {
  assert.equal(
    visibleViewport({
      innerHeight: 480,
      visualViewport: { height: 800, offsetTop: 0, scale: 1 },
    }).height,
    480,
  );
});

test("window height is the fallback when VisualViewport is unavailable", () => {
  assert.deepEqual(visibleViewport({ innerHeight: 480 }), {
    height: 480,
    top: 0,
    zoomed: false,
  });
});

test("pinch zoom is not treated as a keyboard resize", () => {
  assert.deepEqual(
    visibleViewport({
      innerHeight: 844,
      visualViewport: { height: 422, offsetTop: 160, scale: 2 },
    }),
    { height: 844, top: 0, zoomed: true },
  );
});

function recordingArea({
  top = 140,
  bottom = 463,
  inputTop = 572,
  scrollTop = 0,
} = {}) {
  const scroller = {
    scrollTop,
    getBoundingClientRect: () => ({ top, bottom }),
  };
  const field = {
    getBoundingClientRect: () => ({
      top: inputTop - scroller.scrollTop,
      bottom: inputTop + 58 - scroller.scrollTop,
      height: 58,
    }),
  };
  return { field, scroller };
}

test("after advancing, a focused reading below the scroll area is brought into view", () => {
  const { field, scroller } = recordingArea();
  const viewport = { top: 0, height: 524 };
  revealInScroller(field, scroller, viewport);
  const bounds = field.getBoundingClientRect();
  assert.ok(bounds.top >= 140 && bounds.bottom <= 463);
});

test("an already visible field does not move on repeated resize or focus events", () => {
  const { field, scroller } = recordingArea({ inputTop: 180 });
  revealInScroller(field, scroller, { top: 0, height: 524 });
  revealInScroller(field, scroller, { top: 0, height: 524 });
  assert.equal(scroller.scrollTop, 0);
});

test("a field above the visible area is recovered after changing users or viewport offset", () => {
  const { field, scroller } = recordingArea({ inputTop: 180, scrollTop: 100 });
  revealInScroller(field, scroller, { top: 150, height: 374 });
  const bounds = field.getBoundingClientRect();
  assert.ok(bounds.top >= 150 && bounds.bottom <= 463);
});
