(function () {
  const initProductAccordion = () => {
    const aboutToggleItems = $(".about__accordion-toggle");

    aboutToggleItems.each(function () {
      const currentToggle = $(this);
      const siblingToggles = currentToggle.siblings(
        ".about__accordion-description"
      );

      currentToggle.click(function () {
        if (!currentToggle.hasClass("active")) {
          aboutToggleItems.each(function () {
            const siblingToggle = $(this);
            siblingToggle.removeClass("active");
            siblingToggle
              .siblings(".about__accordion-description")
              .stop()
              .slideUp(300);
          });

          currentToggle.addClass("active");

          siblingToggles.stop().slideDown(300);
        } else {
          currentToggle.removeClass("active");
          siblingToggles.stop().slideUp(300);
        }
      });
    });
  };

  const initProductCollapsible = (section) => {
    if (!section || !section.classList.contains("product-section")) {
      return;
    }

    const accordions = section.querySelectorAll(".product-accordion");

    accordions.forEach((accordion) => {
      accordion.addEventListener("click", (event) => {
        const toggleEl = event.target.closest(".product-accordion__toggle");
        if (!toggleEl) return;

        const contentEl = toggleEl.nextElementSibling;
        if (
          !contentEl ||
          !contentEl.classList.contains("product-accordion__content")
        ) {
          return;
        }

        const isActive = toggleEl.classList.contains("active");
        if (!isActive) {
          slideDown(toggleEl, contentEl, 300); // func in global.js
        } else {
          slideUp(toggleEl, contentEl, 300); // func in global.js
        }
      });
    });
  };

  const initZoomImage = () => {
    const imagesWrapper = document.querySelector(
      ".product-media-modal__content"
    );
    const images = imagesWrapper?.querySelectorAll(".js-image-zoom") || [];

    images.forEach((el) => {
      el.addEventListener("click", (e) => {
        imagesWrapper.classList.toggle("zoom");
      });
    });
  };

  const formatFreeShippingAmount = (value) => {
    const placeholderRegex = /\{\{\s*(\w+)\s*\}\}/;
    const formatString = theme.moneyFormat;

    return formatString.replace(placeholderRegex, value);
  };

  const setTotalFreeShipping = () => {
    if (document.querySelector(".js-free-shipping")) {
      const freeShippingTotal = document.querySelector(".free-shipping-total");
      if (freeShippingTotal) {
        const minSpend = Number(freeShippingTotal.dataset.minSpend);
        if (minSpend && minSpend > 0) {
          freeShippingTotal.innerText = `${formatFreeShippingAmount(
            Math.round(minSpend * (Shopify.currency.rate || 1))
          )}`;
        }
      }
    }
  };

  const revealStickyAddToCart = (section) => {
    if (!section && section.classList.contains(".product-section")) return;
    const buyButtons = section.querySelector(
      ".product__buy-buttons > product-form"
    );
    const stickyBar = section.querySelector(".sticky-add-to-cart");
    const footerBottom = document.querySelector(
      ".shopify-section-group-footer-group .footer__content-bottom"
    );

    if (!stickyBar) return;

    const observerOptions = {
      root: null,
      rootMargin: "0px",
      threshold: 0.01,
    };

    const isElementBelowScroll = (element) => {
      return element ? element.getBoundingClientRect().top > 0 : false;
    };

    const toggleStickyBar = (isVisible) => {
      if (isVisible || isElementBelowScroll(buyButtons)) {
        stickyBar.classList.remove("active");
      } else {
        stickyBar.classList.add("active");
      }
    };

    const handleIntersect = (entries) => {
      const isVisible = entries.some((entry) => entry.isIntersecting);
      toggleStickyBar(isVisible);
    };

    const observer = new IntersectionObserver(handleIntersect, observerOptions);

    if (buyButtons) observer.observe(buyButtons);
    if (footerBottom) observer.observe(footerBottom);
  };

  const setRadioValue = () => {
    document.querySelectorAll(".product-form__controls").forEach((fieldset) => {
      const legendSpan = fieldset.querySelector(".selected-value");
      if (!legendSpan) return;

      const radios = fieldset.querySelectorAll('input[type="radio"]');

      radios.forEach((radio) => {
        radio.addEventListener("change", function () {
          if (this.checked) {
            legendSpan.textContent = this.value;
          }
        });
      });
    });
  };

  const setTabsText = () => {
    const reviewsTab = document.getElementById("testimonials-tab");
    const allFeaturesTab = document.getElementById("all-features-tab");
    const reviewsSection = document.getElementById("testimonials-quote");
    const mediaSection = document.getElementById("media-section");
    const descriptionGridSection = document.getElementById(
      "description-grid-section"
    );
    const imageWithTextSection = document.getElementById(
      "image-with-text-section"
    );
    const simpleCardsSection = document.getElementById("simple-cards-section");
    const instructionsTab = document.getElementById("instructions-tab");
    const instructionsBlock = document.getElementById("instructions");

    if (!reviewsSection && reviewsTab) {
      reviewsTab.remove();
    }

    if (!instructionsBlock && instructionsTab) {
      instructionsTab.remove();
    }

    if (
      !mediaSection &&
      !descriptionGridSection &&
      !imageWithTextSection &&
      !simpleCardsSection &&
      allFeaturesTab
    ) {
      allFeaturesTab.remove();
    }
  };

  function hideTabsIfSingle() {
    const track = document.querySelector(".tabs-lava-lamp");
    const navItems = document.querySelectorAll(
      ".tabs-lava-lamp li:not(#tabs-lava-lamp-blob)"
    );

    if (track && navItems.length <= 1) {
      track.style.display = "none";
      return true;
    }
    return false;
  }

  const initStickyTabs = () => {
    const tabs = document.querySelector(".sticky-tabs");
    if (!tabs) return;

    const sentinel = document.createElement("div");
    sentinel.style.height = "1px";
    tabs.parentNode.insertBefore(sentinel, tabs);

    const spacer = document.createElement("div");
    spacer.style.display = "none";
    tabs.parentNode.insertBefore(spacer, tabs.nextSibling);

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (!entry.isIntersecting && entry.boundingClientRect.top < 0) {
            spacer.style.height = `${tabs.offsetHeight}px`;
            spacer.style.display = "block";
            tabs.classList.add("is-fixed");
            tabs.classList.add("container");
          } else {
            spacer.style.display = "none";
            tabs.classList.remove("is-fixed");
            tabs.classList.remove("container");
          }
        });
      },
      { threshold: 0 }
    );

    observer.observe(sentinel);
  };

  document.addEventListener("shopify:section:load", function (event) {
    initProductAccordion();
    initProductCollapsible(event.target);
    initZoomImage();
    setTotalFreeShipping();
    revealStickyAddToCart(event.target);
    initStickyTabs();
    setRadioValue();
    setTabsText();
    hideTabsIfSingle();
  });

  initProductAccordion();
  initProductCollapsible(document.currentScript.parentElement);
  initZoomImage();
  setTotalFreeShipping();
  revealStickyAddToCart(document.currentScript.parentElement);
  initStickyTabs();
  setRadioValue();
  setTabsText();
  hideTabsIfSingle();
})();
