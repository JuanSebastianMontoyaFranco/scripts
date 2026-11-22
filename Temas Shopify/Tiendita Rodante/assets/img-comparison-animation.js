(function () {
  function initAnimation() {
    const beforeAfterBlocks = document.querySelectorAll(
      ".before-after__slide--move-in"
    );

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            const imgComparisonSlider = entry.target.querySelector(
              "img-comparison-slider"
            );
            const shadowRoot = imgComparisonSlider?.shadowRoot;
            const firstElement = shadowRoot?.querySelector(".first");

            if (firstElement) {
              firstElement.style.setProperty("--transition-time", "1.5s");
              firstElement.style.setProperty("--exposure", "50%");
            }

            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.5 }
    );

    beforeAfterBlocks.forEach((block) => observer.observe(block));
  }

  document.addEventListener("DOMContentLoaded", initAnimation);
  document.addEventListener("shopify:section:load", function (section) {
    initAnimation(section.target);
  });
})();
