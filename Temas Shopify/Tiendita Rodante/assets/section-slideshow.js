(function () {
  const splitReadyPromises = new WeakMap();
  const swiperInstances = new WeakMap();
  const ANIMATION_DELAY = 0;

  const animateLines = (lines, delay = ANIMATION_DELAY) => {
    if (!lines?.length) return;

    lines.forEach((line, i) => {
      line.classList.add("animated");
      Array.from(line.children).forEach((element) => {
        element.style.animationDelay = `${delay + i * 0.15}s`;
        setTimeout(() => {
          line.style.overflow = "visible";
        }, (delay + i * 0.15) * 1000 + 600);
      });
    });
  };

  const resetLines = (lines) => {
    if (!lines?.length) return;

    lines.forEach((line) => {
      line.classList.remove("animated");
      line.style.overflow = "hidden";

      Array.from(line.children).forEach((el) => {
        el.style.animation = "none";
        el.offsetHeight;
        el.style.animation = "";
        el.style.opacity = "0";
        el.style.transform = "translate3d(0, 100%, 0)";
      });
    });
  };

  const showSplitTextFallback = (heading) => {
    if (!heading) return;
    heading.classList.add("visible");
    heading.style.opacity = "1";
    Array.from(heading.querySelectorAll(".line > *")).forEach((el) => {
      el.style.opacity = "1";
      el.style.transform = "none";
    });
  };

  const getActiveSlide = (section) => {
    return (
      section.querySelector(".slideshow__swiper .swiper-slide-active") ||
      section.querySelector(".slideshow__swiper .swiper-slide")
    );
  };

  const getSlideLines = (slide) => {
    if (!slide) return [];
    return Array.from(slide.querySelectorAll(".js-split-text .line"));
  };

  const resetInactiveSlideLines = (section) => {
    section
      .querySelectorAll(
        ".slideshow__swiper .swiper-slide:not(.swiper-slide-active) .js-split-text .line"
      )
      .forEach((line) => resetLines([line]));
  };

  const playActiveSlideHeadingAnimation = (
    section,
    delay = ANIMATION_DELAY
  ) => {
    const activeSlide = getActiveSlide(section);
    const lines = getSlideLines(activeSlide);

    if (!lines.length) {
      const heading = activeSlide?.querySelector(".js-split-text");
      if (heading) showSplitTextFallback(heading);
      return;
    }

    resetInactiveSlideLines(section);

    const wasAnimated = lines.some((line) =>
      line.classList.contains("animated")
    );
    if (wasAnimated) resetLines(lines);

    requestAnimationFrame(() => {
      animateLines(lines, delay);
    });
  };

  const ensureSplitTypeReady = () => {
    if (window.SplitType) return Promise.resolve();

    return new Promise((resolve) => {
      let attempts = 0;
      const timer = setInterval(() => {
        attempts += 1;
        if (window.SplitType || attempts >= 40) {
          clearInterval(timer);
          resolve();
        }
      }, 16);
    });
  };

  const splitHeadingsInSection = (section) => {
    const headings = section.querySelectorAll(".js-split-text");
    if (!headings.length) return [];

    const instances = [];

    headings.forEach((heading) => {
      try {
        if (window.SplitType?.revert) {
          window.SplitType.revert(heading);
        }

        if (!window.SplitType) {
          showSplitTextFallback(heading);
          return;
        }

        const split = new SplitType(heading, { types: "lines, words" });
        instances.push(split);
        heading.classList.add("visible");
      } catch (error) {
        console.error(error);
        showSplitTextFallback(heading);
      }
    });

    return instances;
  };

  const headingAnimation = (section) => {
    if (!section || !section.classList.contains("slideshow-section")) {
      return Promise.resolve();
    }

    const wrapper = section.querySelector(".slideshow-slide__content");
    if (!wrapper || !wrapper.getAttribute("data-text-animation")) {
      return Promise.resolve();
    }

    if (splitReadyPromises.has(section)) {
      return splitReadyPromises.get(section);
    }

    let splitInstances = [];
    let resizeTimer;

    const runSplit = () => {
      if (splitInstances.length > 0) {
        splitInstances.forEach((instance) => {
          try {
            instance.split({ types: "lines, words" });
          } catch (error) {
            console.error(error);
          }
        });
        section.querySelectorAll(".js-split-text").forEach((heading) => {
          heading.classList.add("visible");
        });
      } else {
        splitInstances = splitHeadingsInSection(section);
      }

      playActiveSlideHeadingAnimation(section, ANIMATION_DELAY);
      return splitInstances;
    };

    const readyPromise = ensureSplitTypeReady().then(() => {
      runSplit();
    });

    splitReadyPromises.set(section, readyPromise);

    if (document.fonts?.ready) {
      document.fonts.ready
        .then(() => {
          runSplit();
        })
        .catch(() => {});
    }

    const resizeObserver = new ResizeObserver(() => {
      clearTimeout(resizeTimer);
      resizeTimer = setTimeout(() => {
        runSplit();
      }, 100);
    });

    readyPromise.then(() => {
      section
        .querySelectorAll(".testimonials__content-text--animation")
        .forEach((text) => resizeObserver.observe(text));
    });

    return readyPromise;
  };

  const initSwiperForSection = (section) => {
    const id = section.getAttribute("id");
    const box = section.querySelector(".slideshow");
    if (!box || !id) return null;

    const autoplay = box.dataset.autoplay;
    const isPrereview = box.dataset.isPrereview;
    const isLoop = box.dataset.loop;
    const stopAutoplay = box.dataset.stopAutoplay;
    const delay = Number(box.dataset.delay) * 1000;
    const slideCount = Number(box.dataset.slideCount);
    const speed = Number(box.dataset.speed) * 1000;

    document.documentElement.style.setProperty(
      "--bullet-duration",
      `${delay + speed}ms`
    );

    let autoplayParam;
    if (autoplay && slideCount > 1) {
      autoplayParam = {
        autoplay: {
          delay: delay,
          pauseOnMouseEnter: stopAutoplay,
          disableOnInteraction: false,
        },
      };
    } else {
      autoplayParam = { autoplay: false };
    }

    const preriewParams = isPrereview ? {} : {};

    const commonParams = {
      speed,
      effect: box.dataset.effect,
      loop: isLoop && slideCount > 1,
      autoHeight: false,
      calculateHeight: false,
      keyboard: true,
      mousewheel: { forceToAxis: true },
      ...autoplayParam,
    };

    const changeColorScheme = (context) => {
      const activeSlide = context.slides[context.activeIndex];
      const changeItems = [
        context.navigation?.nextEl,
        context.navigation?.prevEl,
        context.pagination?.el,
      ].filter(Boolean);
      const colorScheme = activeSlide?.dataset?.colorScheme;

      changeItems.forEach((item) => {
        let classNames = item.getAttribute("class");
        classNames = classNames.replace(/color-background-\d+/g, "");
        item.setAttribute("class", classNames);
        if (colorScheme) item.classList.add(colorScheme);
      });
    };

    const sliderEl = section.querySelector(".slideshow__swiper");
    if (!sliderEl) return null;

    const swiperOverlay = new Swiper(sliderEl, {
      ...commonParams,
      parallax: box.dataset.parallax,
      centeredSlides: false,
      ...preriewParams,
      spaceBetween: 0,
      navigation: {
        nextEl: `#${id} .swiper-button-next`,
        prevEl: `#${id} .swiper-button-prev`,
      },
      pagination: {
        el: `#${id} .swiper-pagination`,
        clickable: true,
      },
      breakpoints: {
        1100: { spaceBetween: 32 },
      },
      on: {
        init: function () {
          playActiveSlideHeadingAnimation(section, ANIMATION_DELAY);
        },
      },
    });

    if (isPrereview) {
      document.addEventListener("focusin", (e) => {
        if (e.target.closest(".slideshow__swiper.swiper.swiper-initialized")) {
          const slide = e.target.closest(".swiper-slide");
          if (!slide) return;
          const index = [...swiperOverlay.slides].indexOf(slide);
          if (swiperOverlay.activeIndex !== index) {
            swiperOverlay.slideTo(index, 0);
          }
        }
      });
    }

    changeColorScheme(swiperOverlay);
    swiperOverlay.on("beforeTransitionStart", function () {
      changeColorScheme(this);
    });

    swiperOverlay.on("slideChangeTransitionStart", function () {
      resetInactiveSlideLines(section);
    });

    swiperOverlay.on("slideChangeTransitionEnd", function () {
      const activeSlide = this.slides[this.activeIndex];
      initButtonsAnimation(activeSlide);
      playActiveSlideHeadingAnimation(section, ANIMATION_DELAY);
    });

    return swiperOverlay;
  };

  const slideshow = (targetSection) => {
    const sections = targetSection
      ? [targetSection].filter((section) =>
          section.classList.contains("slideshow-section")
        )
      : Array.from(document.querySelectorAll(".slideshow-section"));

    sections.forEach((section) => {
      if (section.classList.contains("slider_started")) return;

      section.classList.add("slider_started");

      const splitPromise = headingAnimation(section);

      splitPromise.then(() => {
        const swiperOverlay = initSwiperForSection(section);
        if (swiperOverlay) swiperInstances.set(section, swiperOverlay);
      });
    });
  };

  const onBlockSelect = (event) => {
    const section = event.target.closest(".slideshow-section");
    if (!section) return;

    const swiperOverlay = swiperInstances.get(section);
    if (!swiperOverlay) return;

    const slide = event.target.closest(".swiper-slide");
    if (!slide) return;

    const index = [...swiperOverlay.slides].indexOf(slide);
    if (index === -1 || swiperOverlay.activeIndex === index) return;

    if (swiperOverlay.autoplay?.running) swiperOverlay.autoplay.stop();
    swiperOverlay.slideTo(index, 0);
  };

  const onBlockDeselect = (event) => {
    const section = event.target.closest(".slideshow-section");
    if (!section) return;

    const swiperOverlay = swiperInstances.get(section);
    if (!swiperOverlay) return;

    if (swiperOverlay.params.autoplay && !swiperOverlay.autoplay?.running) {
      swiperOverlay.autoplay.start();
    }
  };

  const initSlideshow = (event) => {
    const section = event?.target?.classList?.contains("slideshow-section")
      ? event.target
      : event?.target?.closest?.(".slideshow-section");

    slideshow(section || null);
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initSlideshow);
  } else {
    initSlideshow();
  }

  document.addEventListener("shopify:section:load", initSlideshow);
  document.addEventListener("shopify:block:select", onBlockSelect);
  document.addEventListener("shopify:block:deselect", onBlockDeselect);
})();
