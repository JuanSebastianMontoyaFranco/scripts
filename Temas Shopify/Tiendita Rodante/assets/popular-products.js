(() => {
  const initSlider = (section) => {
    const popularProducts = section.querySelectorAll(".popular-products");

    for (sectionItem of popularProducts) {
      const isCollectionProductsSelected =
        sectionItem.dataset.isCollectionProductsSelected;

      const alignArrowsToImageHeight = (currentSection) => {
        const firstImage = currentSection.querySelector(".media img");
        if (!firstImage) return;

        const updateArrowPosition = () => {
          const imageHeight = firstImage.offsetHeight;
          if (imageHeight === 0) return;

          const arrows = currentSection.querySelectorAll(
            ".swiper-button-prev, .swiper-button-next"
          );
          arrows.forEach((arrow) => {
            arrow.style.top = `${imageHeight / 2}px`;
            arrow.style.transform = `translateY(-33%)`;
          });
        };

        if (firstImage.complete) {
          updateArrowPosition();
        } else {
          firstImage.addEventListener("load", updateArrowPosition);
        }

        window.addEventListener("resize", updateArrowPosition);
      };

      if (
        isCollectionProductsSelected === "true" ||
        isCollectionProductsSelected === true
      ) {
        // Find all product carousels within this specific sectionItem
        const productCarousels = sectionItem.querySelectorAll(
          '[id^="product-carousel-main-"]'
        );

        productCarousels.forEach((carousel) => {
          // Find slider within the current carousel being iterated
          const slider = carousel.querySelector(".swiper--products");

          if (slider) {
            const numberColumns = slider.dataset.columnsMobile || 1;
            const productCardType = slider.dataset.productCardType || "simple";

            const prodSwiperParams = {
              loop: false,
              slidesPerView: 1,
              allowTouchMove: true,
              centeredSlides: false,
              lazy: true,
              preloadImages: false,
              spaceBetween: 8,
              speed: 800,
              mousewheel: {
                forceToAxis: true,
              },
              navigation: {
                nextEl: carousel.querySelector(".swiper-button-next"),
                prevEl: carousel.querySelector(".swiper-button-prev"),
              },
              pagination: {
                el: carousel.querySelector(".swiper-pagination"),
                clickable: true,
              },
              on: {
                init: function () {
                  this.pagination.render();
                  this.pagination.update();
                },
                slideChange: function () {
                  this.pagination.render();
                  this.pagination.update();
                },
              },
              breakpoints: {
                320: {
                  slidesPerView:
                    productCardType === "simple" ? 2 : Number(numberColumns),
                },
                750: {
                  slidesPerView: productCardType === "simple" ? 3 : 1,
                  spaceBetween: 16,
                },
                990: {
                  slidesPerView: productCardType === "simple" ? 4 : 1.2,
                  spaceBetween: 16,
                },
                1360: {
                  slidesPerView: productCardType === "simple" ? 4 : 1.7,
                  spaceBetween: 16,
                },
              },
            };

            const swiper = new Swiper(
              carousel.querySelector(".product-carousel-swiper"),
              prodSwiperParams
            );

            const tabTargets =
              sectionItem.querySelectorAll("[data-tab-target]");
            tabTargets.forEach((tabTarget) => {
              tabTarget.addEventListener("click", function () {
                setTimeout(() => {
                  swiper?.update();
                  swiper.pagination.render();
                  swiper.pagination.update();
                }, 0);
              });
            });

            alignArrowsToImageHeight(sectionItem);

            const shadowEl = sectionItem.querySelector(
              ".popular-products__content"
            );
            const updateShadow = () => {
              if (!swiper.isEnd) {
                shadowEl.classList.add("slider-shadow-enabled");
                shadowEl.classList.remove("slider-shadow-enable__left");
              } else {
                shadowEl.classList.remove("slider-shadow-enabled");
                shadowEl.classList.add("slider-shadow-enable__left");
              }
            };

            if (shadowEl?.dataset?.sliderShadowEnabled === "true") {
              swiper.on("slideChange", updateShadow);
              swiper.on("reachEnd", updateShadow);
              swiper.on("fromEdge", updateShadow);
            }
          }
        });
      } else {
        // Find slider within the current sectionItem being iterated
        const slider = sectionItem.querySelector(".swiper--products");

        if (slider) {
          const numberColumns = slider.dataset.columnsMobile || 1;
          const productCardType = slider.dataset.productCardType || "simple";

          const swiper = new Swiper(slider, {
            loop: false,
            speed: 800,
            spaceBetween: 8,
            centeredSlides: false,
            mousewheel: {
              forceToAxis: true,
            },
            allowTouchMove: true,
            breakpoints: {
              320: {
                slidesPerView: productCardType === "simple" ? 2 : 1,
              },
              750: {
                slidesPerView: productCardType === "simple" ? 3 : 1,
                spaceBetween: 16,
              },
              990: {
                slidesPerView: productCardType === "simple" ? 4 : 1.2,
                spaceBetween: 16,
              },
              1360: {
                slidesPerView: productCardType === "simple" ? 4 : 1.7,
                spaceBetween: 16,
              },
            },
            navigation: {
              nextEl: sectionItem.querySelector(".swiper-button-next"),
              prevEl: sectionItem.querySelector(".swiper-button-prev"),
            },
            pagination: {
              el: slider.querySelector(".popular-products__pagination"),
              clickable: true,
              type: "custom",
              renderCustom: (swiper, current, total) => {
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

          alignArrowsToImageHeight(sectionItem);

          const shadowEl = sectionItem.querySelector(
            ".popular-products__content"
          );
          const updateShadow = () => {
            if (!swiper.isEnd) {
              shadowEl.classList.add("slider-shadow-enabled");
              shadowEl.classList.remove("slider-shadow-enable__left");
            } else {
              shadowEl.classList.remove("slider-shadow-enabled");
              shadowEl.classList.add("slider-shadow-enable__left");
            }
          };

          if (shadowEl?.dataset?.sliderShadowEnabled === "true") {
            swiper.on("slideChange", updateShadow);
            swiper.on("reachEnd", updateShadow);
            swiper.on("fromEdge", updateShadow);
          }
        }
      }
    }
  };

  const initSection = (section) => {
    if (section && section.classList.contains("popular-products-section")) {
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

  const collectionTabs = (section) => {
    if (section && section.classList.contains("popular-products-section")) {
      const tabs = section.querySelectorAll("[data-tab-target]");
      const tabContents = section.querySelectorAll("[data-tab-content]");

      tabs.forEach((tab, index) => {
        tab.addEventListener("click", () => {
          const target = section.querySelector(tab.dataset.tabTarget);
          tabContents.forEach((tabContent) => {
            tabContent.classList.remove("active");
          });

          tabs.forEach((t) => {
            t.classList.remove("active");
            t.setAttribute("aria-selected", "false");
            t.setAttribute("tabindex", "-1");
          });

          tab.classList.add("active");
          tab.setAttribute("aria-selected", "true");
          tab.setAttribute("tabindex", "0");
          target.classList.add("active");
        });

        tab.addEventListener("keydown", (e) => {
          let newIndex = index;

          if (e.key === "Enter" || e.key === " ") {
            e.preventDefault();
            tab.click();
          } else if (e.key === "ArrowRight" || e.key === "ArrowDown") {
            e.preventDefault();
            newIndex = (index + 1) % tabs.length;
          } else if (e.key === "ArrowLeft" || e.key === "ArrowUp") {
            e.preventDefault();
            newIndex = (index - 1 + tabs.length) % tabs.length;
          }

          if (newIndex !== index) {
            tabs[newIndex].focus();
            tabs[newIndex].click();
          }
        });
      });
    }
  };

  initSection(document.currentScript.parentElement);
  collectionTabs(document.currentScript.parentElement);
  document.addEventListener("DOMContentLoaded", () => {
    document.addEventListener("shopify:section:load", (event) => {
      initSection(event.target);
      collectionTabs(event.target);
    });
  });
})();
