(() => {
  const initSlider = (section) => {
    const slider = section.querySelector(".swiper--products");

    if (!slider) return;

    const numberColumns = slider.dataset.columnsMobile || 1;
    const productsSize = parseInt(slider.dataset.productsSize) || 0;
    const showProductCount =
      slider.dataset.showProductCount === "true" ||
      slider.dataset.showProductCount === true;

    const swiper = new Swiper(slider, {
      loop: false,
      speed: 800,
      spaceBetween: 8,
      allowTouchMove: true,
      slidesPerView: 1,
      slidesPerGroup: 1,
      mousewheel: {
        forceToAxis: true,
      },
      breakpoints: {
        320: {
          slidesPerView: 2,
        },
        750: {
          slidesPerView: 3,
          spaceBetween: 16,
        },
        990: {
          slidesPerView: showProductCount ? 3 : 4,
          spaceBetween: 16,
        },
        1100: {
          slidesPerView: 4,
          spaceBetween: 16,
        },
        1360: {
          slidesPerView: showProductCount ? 4.4 : 4,
          spaceBetween: 16,
        },
      },
      navigation: {
        nextEl: `#${section.id} .swiper-button-next`,
        prevEl: `#${section.id} .swiper-button-prev`,
      },
      pagination: {
        el: slider.querySelector(".popular-products__pagination"),
        clickable: true,
        type: "custom",
        renderCustom: function (swiper, current, total) {
          let out = "";
          for (let i = 1; i < total + 1; i++) {
            out += `<span class="swiper-pagination-bullet${
              i === current ? " swiper-pagination-bullet-active" : ""
            }" tabindex="0" role="button" aria-label="Go to slide ${i}"></span>`;
          }
          return out;
        },
      },
      on: {
        beforeInit: function (swiper) {
          const slides = slider.querySelectorAll(".swiper-slide");
          const screenWidth = window.innerWidth;

          let requiredSlides = 1;
          if (screenWidth >= 1200) {
            requiredSlides = showProductCount ? 4.3 : 4;
          } else if (screenWidth >= 990) {
            requiredSlides = 3.3;
          } else if (screenWidth >= 750) {
            requiredSlides = 3;
          } else {
            requiredSlides = 2;
          }

          if (slides.length < requiredSlides) {
            swiper.disable();
            const arrows = section.querySelectorAll(
              ".swiper-button-prev, .swiper-button-next"
            );
            arrows.forEach((el) => (el.style.display = "none"));
          }
        },
        transitionEnd: function (swiper) {
          swiper.navigation.update();
        },
      },
    });

    const alignArrowsToImageHeight = () => {
      const firstImage = section.querySelector(".media img");
      if (!firstImage) return;

      const updateArrowPosition = () => {
        const imageHeight = firstImage.offsetHeight;
        if (imageHeight === 0) return;

        const arrows = section.querySelectorAll(
          ".swiper-button-prev, .swiper-button-next"
        );
        arrows.forEach((arrow) => {
          arrow.style.top = `${imageHeight / 2}px`;
          arrow.style.transform = `translateY(-50%)`;
        });
      };

      if (firstImage.complete) {
        updateArrowPosition();
      } else {
        firstImage.addEventListener("load", updateArrowPosition);
      }

      window.addEventListener("resize", updateArrowPosition);
    };

    alignArrowsToImageHeight();

    const shadowEl = section.querySelector(".popular-products__content");

    const updateShadow = () => {
      if (!swiper.isEnd) {
        shadowEl.classList.add("slider-shadow-enable");
        shadowEl.classList.remove("slider-shadow-enable__left");
      } else {
        shadowEl.classList.remove("slider-shadow-enable");
        shadowEl.classList.add("slider-shadow-enable__left");
      }
    };

    if (shadowEl?.dataset?.enableSliderShadow === "true") {
      swiper.on("slideChange", updateShadow);
      swiper.on("reachEnd", updateShadow);
      swiper.on("fromEdge", updateShadow);
    }
  };

  const initSection = (section) => {
    if (section && section.classList.contains("section-product-slider")) {
      const resizeObserver = new ResizeObserver((entries) => {
        entries.forEach((entry) => {
          if (entry.contentRect.width > 1100) {
            const productItems = entry.target.querySelectorAll(
              ".popular-products__item"
            );

            productItems.forEach((item) => {
              item.addEventListener("mouseover", () => {
                productItems.forEach((otherItem) => {
                  if (otherItem !== item) {
                    otherItem.classList.add("popular-products__item--inactive");
                  }
                });
              });

              item.addEventListener("mouseout", () => {
                productItems.forEach((item) => {
                  item.classList.remove("popular-products__item--inactive");
                });
              });
            });
          }

          initSlider(entry.target);
        });
      });

      resizeObserver.observe(section);
    }
  };

  initSection(document.currentScript.parentElement);

  document.addEventListener("shopify:section:load", function (event) {
    initSection(event.target);
  });
})();
