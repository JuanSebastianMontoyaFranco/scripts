(function () {
  const initSection = (section) => {
    if (
      !section ||
      !section.classList.contains("section-featured-testimonials-highlight")
    ) {
      return;
    }

    const sliderEl = section.querySelector(".testimonials__highlight__posts");

    if (sliderEl) {
      if (!sliderEl.classList.contains("swiper")) return;
      if (sliderEl.classList.contains("swiper-initialized")) return;

      const wrapperSection = sliderEl.closest("[id]");
      if (!wrapperSection) return;

      const nextButton = wrapperSection.querySelector(".swiper-button-next");
      const prevButton = wrapperSection.querySelector(".swiper-button-prev");
      const paginationEl = sliderEl.querySelector(".swiper-pagination");

      const swiper = new Swiper(sliderEl, {
        slidesPerView: 1,
        spaceBetween: 8,
        speed: 800,
        centeredSlides: false,
        loop: false,
        initialSlide: 0,
        pagination: {
          el: paginationEl,
          clickable: true,
        },
        mousewheel: {
          forceToAxis: true,
        },
        breakpoints: {
          //750: {
          //	slidesPerView: 1.05,
          //	spaceBetween: 16,
          //},
          990: {
            slidesPerView: 1.1,
            spaceBetween: 16,
          },
          1100: {
            slidesPerView: 1.4,
            spaceBetween: 16,
          },
          1360: {
            slidesPerView: 1.7,
            spaceBetween: 16,
          },
        },
        navigation: {
          nextEl: nextButton,
          prevEl: prevButton,
        },
      });

      const updateShadow = () => {
        const shadowEl = wrapperSection.querySelector(
          ".testimonials__highlight__posts--shadow"
        );

        if (swiper.isEnd) {
          shadowEl.classList.add("shadow-hidden");
          shadowEl.classList.add(
            "testimonials__highlight__posts--shadow__left"
          );
        } else {
          shadowEl.classList.remove("shadow-hidden");
          shadowEl.classList.remove(
            "testimonials__highlight__posts--shadow__left"
          );
        }
      };

      swiper.on("slideChange", updateShadow);
      swiper.on("reachEnd", updateShadow);
      swiper.on("fromEdge", updateShadow);
    }
  };

  initSection(document.currentScript.parentElement);

  document.addEventListener("shopify:section:load", function (event) {
    initSection(event.target);
  });
})();
