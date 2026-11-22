(() => {
  function initAnimation() {
    document.querySelectorAll(".magnetic-hover-effect").forEach((el) => {
      el.addEventListener("mousemove", (e) => {
        const rect = el.getBoundingClientRect();
        const x = ((e.clientX - rect.left) / rect.width) * 100;
        const y = ((e.clientY - rect.top) / rect.height) * 100;

        const percentX = Math.min(60, Math.max(40, x));
        const percentY = Math.min(60, Math.max(40, y));

        el.style.setProperty("--glow-position-x", `${percentX}%`);
        el.style.setProperty("--glow-position-y", `${percentY}%`);
      });
    });
  }

  document.addEventListener("DOMContentLoaded", () => initAnimation());
  document.addEventListener("shopify:section:load", (e) =>
    initAnimation(e.target)
  );
  if (document.readyState !== "loading") initAnimation();
})();
