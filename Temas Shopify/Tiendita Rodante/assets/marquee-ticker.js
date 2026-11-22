(function () {
  const changeDirectionTicker = (className = "") => {
    const marqueeContents = document.querySelectorAll(className);
    let lastScrollTop =
      window.pageYOffset || document.documentElement.scrollTop;

    window.addEventListener("scroll", () => {
      const scrollTop =
        window.pageYOffset || document.documentElement.scrollTop;
      const directionAnimationName =
        scrollTop > lastScrollTop ? "slide" : "slideReverse";

      marqueeContents.forEach((element) => {
        element.style.animationName = directionAnimationName;
      });

      lastScrollTop = scrollTop <= 0 ? 0 : scrollTop;
    });
  };

  changeDirectionTicker(".promo-banner .marquee-tag.marquee-animate");
  changeDirectionTicker(".image-banner .marquee-tag.marquee-animate");
})();
