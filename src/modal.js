export const vModal = {
  mounted(element) {
    const previous = document.activeElement;
    const focusable = () =>
      [
        ...element.querySelectorAll(
          'button, input, select, textarea, [tabindex="0"]',
        ),
      ].filter((node) => !node.disabled && node.getClientRects().length);
    const trap = (event) => {
      if (event.key !== "Tab") return;
      const nodes = focusable(),
        first = nodes[0],
        last = nodes.at(-1);
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last?.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first?.focus();
      }
    };
    element.addEventListener("keydown", trap);
    queueMicrotask(() => focusable()[0]?.focus());
    element.cleanupModal = () => {
      element.removeEventListener("keydown", trap);
      previous?.focus();
    };
  },
  unmounted(element) {
    element.cleanupModal?.();
  },
};
