(function () {
  const setImageClasses = (section, activeId) => {
    const images = section.querySelectorAll(
      ".image-tabs__media--desktop .image-tabs__image"
    );

    images.forEach((image) => {
      image.classList.remove("image-tabs__image--active");
    });

    const activeImage = section.querySelector(
      `.image-tabs__media--desktop [data-image-tabs="${activeId}"].image-tabs__image`
    );

    activeImage.classList.add("image-tabs__image--active");
  };

  const initClassesDesktop = (section) => {
    const blocks = section.querySelectorAll(".image-tabs__block");
    blocks.forEach((block) => {
      block.classList.remove("swiper-slide");
    });
    blocks[0].classList.add("image-tabs__block--active");
    blocks[0].classList.add("active");

    setImageClasses(section, "#image-tabs-1");
  };

  const initClassesMobile = (section) => {
    const slides = section.querySelectorAll(".image-tabs__block");
    slides.forEach((slide) => {
      slide.classList.remove("image-tabs__block--active");
      slide.classList.add("swiper-slide");
    });
  };

  const blockToggleDesktop = (section) => {
    const blocks = section.querySelectorAll(".image-tabs__block");

    const setBlockClasses = (activeBlock) => {
      blocks.forEach((block) => {
        block.classList.remove("image-tabs__block--active");
        block.classList.remove("active");
      });
      activeBlock.classList.add("image-tabs__block--active");
      activeBlock.classList.add("active");
    };

    blocks.forEach((block) => {
      block.addEventListener("click", () => {
        const activeId = block.dataset.imageTabs;
        setBlockClasses(block);
        setImageClasses(section, activeId);
      });

      block.addEventListener("mouseenter", () => {
        const hoverId = block.dataset.imageTabs;
        block.classList.add("image-tabs__block--active");
        setImageClasses(section, hoverId);
      });

      block.addEventListener("mouseleave", () => {
        block.classList.remove("image-tabs__block--active");

        let activeId;
        blocks.forEach((block) => {
          if (block.classList.contains("active")) {
            activeId = block.dataset.imageTabs;

            setBlockClasses(block);
            setImageClasses(section, activeId);
          }
        });
      });
    });
  };

  const initSliderMobile = (section) => {
    const sliderTabs = section.querySelector(".swiper--image-tabs");
    const sliderImage = section.querySelector(".image-tabs__media--mobile");

    if (sliderTabs && sliderImage) {
      const pagination = sliderTabs.querySelector(".image-tabs__pagination");

      const swiperTabs = new Swiper(sliderTabs, {
        loop: false,
        speed: 800,
        spaceBetween: 8,
        mousewheel: {
          forceToAxis: true,
        },
        pagination: {
          el: pagination,
          clickable: true,
          type: "custom",
          renderCustom: function (swiper, current, total) {
            let out = "";
            for (i = 1; i < total + 1; i++) {
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

      const swiperImage = new Swiper(sliderImage, {
        loop: false,
        speed: 800,
        spaceBetween: 8,
      });

      swiperTabs.controller.control = swiperImage;
      swiperImage.controller.control = swiperTabs;
    }
  };

  const destroySlider = (section) => {
    const slides = section.querySelectorAll(".image-tabs__block");

    slides.forEach((slide) => {
      slide.removeAttribute("style");
      slide.classList.remove("swiper-slide");
    });
  };

  const initSection = (section) => {
    if (section && section.classList.contains("image-tabs-section")) {
      const sectionResizeObserver = new ResizeObserver((entries) => {
        entries.forEach((entry) => {
          if (entry.contentRect.width >= 990) {
            initClassesDesktop(entry.target);
            blockToggleDesktop(entry.target);
          }

          if (entry.contentRect.width < 990) {
            initClassesMobile(entry.target);
            initSliderMobile(entry.target);
          } else {
            destroySlider(entry.target);
          }
        });
      });

      sectionResizeObserver.observe(section);
    }
  };

  initSection(document.currentScript.parentElement);

  document.addEventListener("shopify:section:load", function (section) {
    initSection(section.target);
  });
})();
