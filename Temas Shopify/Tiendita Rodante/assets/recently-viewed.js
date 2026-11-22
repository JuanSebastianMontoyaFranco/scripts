(function () {
  const addSliderClasses = (slider) => {
    const sliderWrapper = slider.querySelector(".products-grid__wrapper");
    const slides = slider.querySelectorAll(".collection-product-card");

    slider.classList.add("swiper");
    if (sliderWrapper) sliderWrapper.classList.add("swiper-wrapper");

    slides.forEach((slide) => {
      slide.classList.add("swiper-slide");
    });
  };

  const removeSliderClasses = (slider) => {
    const sliderWrapper = slider.querySelector(".products-grid__wrapper");
    const slides = slider.querySelectorAll(".collection-product-card");

    slider.classList.remove("swiper");
    if (sliderWrapper) sliderWrapper.classList.remove("swiper-wrapper");

    slides.forEach((slide) => {
      slide.removeAttribute("style");
      slide.classList.remove("swiper-slide");
    });
  };

  const initSlider = (section) => {
    const slider = section.querySelector(".swiper--products-grid");

    if (slider) {
      addSliderClasses(slider);
      const numberColumns = slider.dataset.columnsMobile || 1;

      new Swiper(slider, {
        loop: false,
        speed: 800,
        mousewheel: {
          forceToAxis: true,
        },
        breakpoints: {
          320: {
            slidesPerView: Number(numberColumns),
            slidesPerGroup: Number(numberColumns),
            spaceBetween: 8,
          },
          750: {
            slidesPerView: 2,
            slidesPerGroup: 2,
            spaceBetween: 8,
          },
        },
        pagination: {
          el: slider.querySelector(".products-grid__pagination"),
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
      });
    }
  };

  const destroySlider = (section) => {
    const slider = section.querySelector(".swiper--products-grid");

    if (slider) {
      removeSliderClasses(slider);
    }
  };

  const initSliderObserver = (section) => {
    const resizeObserver = new ResizeObserver((entries) => {
      entries.forEach((entry) => {
        if (entry.contentRect.width < 990) {
          initSlider(entry.target);
        } else {
          destroySlider(entry.target);
        }
      });
    });

    resizeObserver.observe(section);
  };

  const initSection = async (section) => {
    if (!section || !section?.classList.contains("recently-viewed-section")) {
      return;
    }

    const box = section.querySelector(".recently-viewed");
    if (!box) return;

    const STORAGE_KEY = "__sf_theme_recently";
    const EXPIRATION_DAYS = box.dataset.expirationDays
      ? Number(box.dataset.expirationDays)
      : 30;
    const dateNow = Date.now();

    const baseUrl = box.dataset.baseUrl;
    const productsLimit = Number(box.dataset.productsLimit) || 6;
    const currentPageProductId = box.dataset.currentPageProductId;
    const hasMobileSlider = box.dataset.sliderMobile === "true";

    // get recent products from local storage
    let recentProducts = [];
    try {
      recentProducts = JSON.parse(localStorage.getItem(STORAGE_KEY)) || [];
    } catch (e) {
      console.error(`Incorrect value in local storage for "${STORAGE_KEY}"`);
    }

    if (currentPageProductId) {
      recentProducts = recentProducts.filter(
        (item) => item.productId !== currentPageProductId
      );
    }

    if (recentProducts.length === 0) {
      box.classList.remove("recently-viewed--loading");
      box.classList.add("recently-viewed--empty");
      return;
    }

    // filter by expiration time
    const expirationTime = EXPIRATION_DAYS * 24 * 60 * 60 * 1000;
    const validProducts = recentProducts.filter(
      (item) => dateNow - item.timestamp < expirationTime
    );

    // limit by section setting
    const limitedProducts = validProducts.slice(0, productsLimit);

    // get url with query
    const query = limitedProducts
      .filter((item) => item.productId)
      .map((item) => `id:${item.productId}`)
      .join("%20OR%20");
    const url = `${baseUrl}&q=${query}`;

    try {
      const response = await fetch(url);
      const html = await response.text();
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");
      const sourceBox = doc?.querySelector(".recently-viewed");
      if (!sourceBox?.classList.contains("recently-viewed--search-perfomed")) {
        box.classList.add("recently-viewed--empty");
        return;
      }
      box.innerHTML = sourceBox.innerHTML;

      try {
        colorSwatches();
        if (hasMobileSlider) initSliderObserver(section);
      } catch (err) {}
    } catch (error) {
      console.error("Failed to fetch recently viewed products:", error);
      box.classList.add("recently-viewed--empty");
    } finally {
      box.classList.remove("recently-viewed--loading");
    }
  };

  initSection(document.currentScript.parentElement);

  document.addEventListener("shopify:section:load", function (event) {
    initSection(event.target);
  });
})();
