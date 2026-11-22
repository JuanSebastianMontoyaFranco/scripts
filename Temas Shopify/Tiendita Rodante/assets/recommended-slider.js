(function () {
  const recommendedSlider = () => {
    $(".recommended-slider-section").each(function () {
      if ($(this).hasClass("slider_started")) {
        return "";
      }
      $(this).addClass("slider_started");

      const id = $(this).attr("id");
      const box = $(this).find(".recommended-slider");
      const autoplay = box.data("autoplay");
      const stopAutoplay = box.data("stop-autoplay");
      const delay = box.data("delay") * 1000;
      const speed = box.data("speed") * 1000;
      const perRow = box.data("per-row");
      let smallPerRow = false;

      if (+perRow <= 2) {
        smallPerRow = true;
      }

      let autoplayParm = {};
      if (autoplay) {
        autoplayParm = {
          autoplay: {
            delay: delay,
            pauseOnMouseEnter: stopAutoplay,
            disableOnInteraction: false,
            waitForTransition: true,
          },
        };
      }

      let swiperParms = {
        speed: speed,
        loop: false,
        centeredSlides: false,
        autoHeight: false,
        calculateHeight: false,
        keyboard: true,
        slidesPerView: 1,
        spaceBetween: 8,
        mousewheel: {
          forceToAxis: true,
        },
        breakpoints: {
          320: {
            slidesPerView: smallPerRow ? perRow : 2,
            spaceBetween: 8,
          },
          750: {
            slidesPerView: smallPerRow ? perRow : 2,
            spaceBetween: 16,
          },
          990: {
            slidesPerView: smallPerRow ? perRow : 3,
            spaceBetween: 16,
          },
          //1100: {
          //	slidesPerView: smallPerRow ? perRow : 3,
          //	spaceBetween: 16,
          //},
          1360: {
            slidesPerView: perRow,
            spaceBetween: 16,
          },
        },
        navigation: {
          nextEl: `#${id} .swiper-button-next`,
          prevEl: `#${id} .swiper-button-prev`,
        },
        pagination: {
          el: `#${id} .recommended-slider__pagination`,
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
        ...autoplayParm,
      };

      const swiper = new Swiper(`#${id} .swiper`, swiperParms);

      setTimeout(() => {
        swiper.update();
        swiper.navigation.update();
      }, 300);
    });
  };

  document.addEventListener("DOMContentLoaded", function () {
    recommendedSlider();
    document.addEventListener("shopify:section:load", function () {
      recommendedSlider();
    });
  });
})();
