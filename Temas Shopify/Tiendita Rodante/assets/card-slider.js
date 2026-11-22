(() => {
  const initCardSliders = (section) => {
    if (!section || !section.classList.contains("card-slider-section")) return;

    const slider = section.querySelector(".card-slider__wrapper");
    if (slider.classList.contains("slider-initialized")) return;

    const container = slider.querySelector(".swiper");
    if (!container) return;

    const perRow = Number.parseInt(slider.dataset.perRow) || 4;

    slider.style.setProperty("--columns-per-row", perRow);

    const swiperOptions = {
      slidesPerView: 1,
      spaceBetween: 8,
      loop: false,
      speed: 800,
      mousewheel: {
        forceToAxis: true,
      },
      breakpoints: {
        320: {
          slidesPerView: 2,
          //slidesPerGroup: 1,
          spaceBetween: 8,
        },
        750: {
          slidesPerView: 3,
          //slidesPerGroup: 2,
          spaceBetween: 14,
        },
        990: {
          slidesPerView: 4,
          //slidesPerGroup: 3,
          spaceBetween: 14,
        },
        1100: {
          slidesPerView: perRow === 6 ? 5 : perRow,
          //slidesPerGroup: perRow,
          spaceBetween: 14,
        },
        1360: {
          slidesPerView: perRow,
          //slidesPerGroup: perRow,
          spaceBetween: 14,
        },
      },
      pagination: {
        el: slider.querySelector(".card-slider__pagination"),
        clickable: true,
        type: "custom",
        renderCustom: function (swiper, current, total) {
          let out = "";
          for (let i = 1; i < total + 1; i++) {
            if (i == current) {
              out = `${out}<span class="swiper-pagination-bullet swiper-pagination-bullet-active" tabindex="0" role="button" aria-label="Go to slide ${i}"></span>`;
            } else {
              out = `${out}<span class="swiper-pagination-bullet" tabindex="0" role="button" aria-label="Go to slide ${i}"></span>`;
            }
          }
          return out;
        },
      },
    };

    const nextButton = slider.querySelector(".card-slider__button--next");
    const prevButton = slider.querySelector(".card-slider__button--prev");
    if (nextButton && prevButton) {
      swiperOptions.navigation = {
        nextEl: nextButton,
        prevEl: prevButton,
      };
    }

    const swiper = new Swiper(container, swiperOptions);
    slider.classList.add("slider-initialized");
  };

  const initAppearanceAnimation = (section) => {
    if (!section || !section.classList.contains("card-slider-section")) return;

    const slider = section.querySelector(".card-slider__wrapper");
    if (!slider || !slider.classList.contains("js-appearance-animation"))
      return;

    const animateSlides = (slider) => {
      const elements = slider.querySelectorAll(".info-card:not(.animated)");
      elements.forEach((element) => {
        element.classList.add("animated");

        element.addEventListener("animationend", () => {
          element.classList.add("animation-end");
        });
      });
      slider.classList.add("appearance-initialized");
    };

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            animateSlides(entry.target);
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.2 }
    );

    observer.observe(slider);
  };

  const initShadowHoverEffect = (section) => {
    if (!section || !section.classList.contains("card-slider-section")) return;

    const wrapper = section.querySelector(".card-slider__wrapper");
    if (!wrapper) return;

    const hoverEffectEl = wrapper.querySelector(".js-info-card-hover-effect");
    if (!hoverEffectEl) return;

    wrapper.addEventListener("mousemove", (e) => {
      const targetCard = e.target.closest(".card-slider__slide");
      const targetContent = targetCard?.querySelector(".info-card");

      if (
        targetCard &&
        targetContent &&
        targetContent.classList.contains("animation-end")
      ) {
        const cardRect = targetCard.getBoundingClientRect();
        const wrapperRect = wrapper.getBoundingClientRect();

        const width = cardRect.width;
        const height = cardRect.height * 1.2;

        const x = ((e.clientX - cardRect.left) / cardRect.width) * 100;
        const y = ((e.clientY - cardRect.top) / cardRect.height) * 100;

        const percentX = Math.min(60, Math.max(40, x));
        const percentY = Math.min(60, Math.max(40, y));

        const offsetX = ((percentX - 50) / 100) * cardRect.width;
        const offsetY = ((percentY - 50) / 100) * cardRect.height;

        const cardCenterX =
          cardRect.left - wrapperRect.left + cardRect.width / 2;
        const cardCenterY =
          cardRect.top - wrapperRect.top + cardRect.height / 2;

        const wrapperCenterX = wrapperRect.width / 2;
        const wrapperCenterY = wrapperRect.height / 2;

        hoverEffectEl.style.setProperty(
          "--glow-translate-x",
          `${cardCenterX + offsetX - wrapperCenterX}px`
        );
        hoverEffectEl.style.setProperty(
          "--glow-translate-y",
          `${cardCenterY + offsetY - wrapperCenterY}px`
        );
        hoverEffectEl.style.width = `${width}px`;
        hoverEffectEl.style.height = `${height}px`;
        hoverEffectEl.style.opacity = "0.3";
      } else {
        hoverEffectEl.style.opacity = "0";
      }
    });

    wrapper.addEventListener("mouseleave", () => {
      hoverEffectEl.style.opacity = "0";
      hoverEffectEl.style.width = "0";
      hoverEffectEl.style.height = "0";
    });
  };

  initCardSliders(document.currentScript.parentElement);
  initAppearanceAnimation(document.currentScript.parentElement);
  initShadowHoverEffect(document.currentScript.parentElement);

  document.addEventListener("shopify:section:load", (event) => {
    initCardSliders(event.target);
    initAppearanceAnimation(event.target);
    initShadowHoverEffect(event.target);
  });
})();
