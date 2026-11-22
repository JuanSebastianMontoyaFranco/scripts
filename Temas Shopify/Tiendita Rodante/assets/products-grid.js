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

  function getCollectionUrl(baseUrl, sortType) {
    const sortTypes = {
      alphabetical: "title-ascending",
      alphabetical_reversed: "title-descending",
      products_low: "price-ascending",
      products_high: "price-descending",
      date: "created-ascending",
      date_reversed: "created-descending",
    };

    const sortBy = sortTypes[sortType] || sortTypes.alphabetical;
    return `${baseUrl}?sort_by=${sortBy}`;
  }

  async function handleResponse(response, section, limitProducts) {
    const html = await response.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, "text/html");

    const targetGrid = section.querySelector(".products-grid__wrapper");
    const sourceGrid = doc?.getElementById("product-grid");
    if (!targetGrid || !sourceGrid || !sourceGrid?.children) return;

    const sourceItems = Array.from(sourceGrid.children);
    if (sourceItems.length > limitProducts) {
      const firstItems = sourceItems.slice(0, limitProducts);
      targetGrid.innerHTML = "";
      firstItems.forEach((item) => {
        targetGrid.appendChild(item);
      });
    } else {
      targetGrid.innerHTML = sourceGrid.innerHTML;
    }
  }

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
    if (!section || !section.classList.contains("products-grid-section"))
      return;

    const box = section.querySelector(".products-grid");
    if (!box) return;

    const baseUrl = box.dataset.baseUrl;
    const totalShopProducts = Number(box.dataset.totalShopProducts);
    const sortType = box.dataset.sortType;
    const limitProducts = Number(box.dataset.limitProducts);
    const hasMobileSlider = box.dataset.sliderMobile === "true";

    if (baseUrl && baseUrl !== "none" && sortType && totalShopProducts > 50) {
      const url = getCollectionUrl(baseUrl, sortType);

      box.classList.add("products-grid--loading");
      try {
        const response = await fetch(url);
        await handleResponse(response, section, limitProducts);
        initButtonsAnimation(section);
        if (hasMobileSlider) initSliderObserver(section);
      } catch (error) {
        console.error(
          "Products grid: Error when fetching the collection page:",
          error
        );
      } finally {
        box.classList.remove("products-grid--loading");
      }
    } else {
      box.classList.remove("products-grid--loading");
      if (hasMobileSlider) initSliderObserver(section);
    }
  };

  initSection(document.currentScript.parentElement);

  document.addEventListener("shopify:section:load", function (event) {
    initSection(event.target);
  });
})();
