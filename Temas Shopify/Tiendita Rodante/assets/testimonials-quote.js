(function () {
  const initSwiper = (sliderEl, columnPerRow = 1) => {
    if (sliderEl.classList.contains("swiper-initialized")) return;

    const wrapperSection = sliderEl.closest("[id]");
    if (!wrapperSection) return;

    const nextButton = wrapperSection.querySelector(".swiper-button-next");
    const prevButton = wrapperSection.querySelector(".swiper-button-prev");
    const paginationEl = wrapperSection.querySelector(".swiper-pagination");

    new Swiper(sliderEl, {
      slidesPerView: 1,
      spaceBetween: 16,
      centeredSlides: false,
      loop: false,
      initialSlide: 0,
      mousewheel: {
        forceToAxis: true,
      },
      pagination: {
        el: paginationEl,
        clickable: true,
      },
      navigation: {
        nextEl: nextButton,
        prevEl: prevButton,
      },
      breakpoints: {
        576: {
          slidesPerView: 2,
          centeredSlides: false,
        },
        990: {
          slidesPerView: columnPerRow,
        },
      },
    });
  };

  const setupResponsiveSwiper = (scope = document) => {
    const sliderEls = scope?.querySelectorAll(".testimonials-quote__posts");
    if (!sliderEls.length) return;

    sliderEls.forEach((sliderEl) => {
      const columnPerRow = parseFloat(sliderEl.dataset.column_per_row) || 1;
      initSwiper(sliderEl, columnPerRow);
    });
  };

  document.addEventListener("DOMContentLoaded", () => {
    setupResponsiveSwiper(document);
  });

  document.addEventListener("shopify:section:load", (event) => {
    setupResponsiveSwiper(event.target);
  });
})();
