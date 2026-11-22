function getProductMediaModalRoot(modalRoot) {
  if (modalRoot?.matches?.("product-modal")) return modalRoot;
  if (modalRoot?.closest) {
    const fromClosest = modalRoot.closest("product-modal");
    if (fromClosest) return fromClosest;
  }
  return (
    document.querySelector("product-modal.product-media-modal[open]") ||
    document.querySelector("product-modal.product-media-modal")
  );
}

function updatePopupSwiperPagination(swiper) {
  if (!swiper?.pagination) return;
  if (typeof swiper.pagination.render === "function") {
    swiper.pagination.render();
  }
  if (typeof swiper.pagination.update === "function") {
    swiper.pagination.update();
  }
}

/**
 * Main gallery lists featured media first; the modal gallery uses product.media order.
 * Syncing popup Swiper by slide index shows the wrong frame — match by numeric media id.
 */
function syncPopupSwiperToProductMediaId(numericMediaId, speed = 0, modalRoot) {
  if (numericMediaId == null || numericMediaId === "") return;
  const modal = getProductMediaModalRoot(modalRoot);
  const popupEl = modal?.querySelector(".js-popup-slider");
  const swiper = popupEl?.swiper;
  if (!swiper?.slides?.length) return;

  const safe = CSS.escape(String(numericMediaId));
  const slides = swiper.slides;
  for (let i = 0; i < slides.length; i++) {
    if (slides[i].querySelector(`[data-media-id="${safe}"]`)) {
      if (swiper.activeIndex !== i) {
        swiper.slideTo(i, speed);
      }
      updatePopupSwiperPagination(swiper);
      return;
    }
  }
}

function syncPopupSwiperFromMainGallery(swiperInstance, speed = 0) {
  if (!swiperInstance?.slides?.length) return;
  const activeSlide = swiperInstance.slides[swiperInstance.activeIndex];
  const toggle = activeSlide?.querySelector(".product__media-toggle");
  const id = toggle?.getAttribute("data-media-id");
  syncPopupSwiperToProductMediaId(id, speed);
}

window.syncProductPopupToMediaId = syncPopupSwiperToProductMediaId;

const sliderInit = (isUpdate) => {
  if (
    document.querySelectorAll(".js-media-list") &&
    document.querySelectorAll(".js-media-list").length > 0
  ) {
    document.querySelectorAll(".js-media-list").forEach((elem) => {
      const mediaListId = elem.dataset?.jsMediaListId;
      const mediaSublist = Array.from(
        document.querySelectorAll(".js-media-sublist")
      ).find(
        (subElem) =>
          subElem.dataset?.jsMediaListId === mediaListId && subElem.swiper
      );
      let slider = new Swiper(elem, {
        slidesPerView: 1,
        spaceBetween: 1,
        autoHeight: true,
        speed: 800,
        navigation: {
          nextEl:
            ".product .product__slider-nav .swiper-button-next, .product .product__gallery-arrows .swiper-button-next",
          prevEl:
            ".product .product__slider-nav .swiper-button-prev, .product .product__gallery-arrows .swiper-button-prev",
        },
        pagination: {
          el: ".product .product__pagination",
          type: "bullets",
          clickable: true,
        },
        thumbs: {
          swiper: mediaListId && mediaSublist ? mediaSublist.swiper : "",
        },
        on: {
          slideChangeTransitionStart: function () {
            if (mediaListId && mediaSublist) {
              mediaSublist.swiper.slideTo(this.activeIndex);
            }
          },
          slideChange: function () {
            window.pauseAllMedia();
            this.params.noSwiping = false;

            syncPopupSwiperFromMainGallery(this, 0);
          },
          slideChangeTransitionEnd: function () {
            if (this.slides[this.activeIndex].querySelector("model-viewer")) {
              this.slides[this.activeIndex]
                .querySelector(".shopify-model-viewer-ui__button--poster")
                .removeAttribute("hidden");
            }
          },
          touchStart: function () {
            if (this.slides[this.activeIndex].querySelector("model-viewer")) {
              if (
                !this.slides[this.activeIndex]
                  .querySelector("model-viewer")
                  .classList.contains("shopify-model-viewer-ui__disabled")
              ) {
                this.params.noSwiping = true;
                this.params.noSwipingClass = "swiper-slide";
              } else {
                this.params.noSwiping = false;
              }
            }
          },
        },
      });

      if (isUpdate) {
        setTimeout(function () {
          slider.update();
        }, 800);
      }
    });
  }
};

const subSliderInit = (isUpdate) => {
  if (
    document.querySelectorAll(".js-media-sublist") &&
    document.querySelectorAll(".js-media-sublist").length > 0
  ) {
    document.querySelectorAll(".js-media-sublist").forEach((elem) => {
      let subSlider = new Swiper(elem, {
        slidesPerView: 5,
        spaceBetween: 8,
        direction: "horizontal",
        freeMode: false,
        watchSlidesProgress: true,
        autoHeight: false,
        on: {
          touchEnd: function (s, e) {
            let range = 5;
            let diff = (s.touches.diff = s.isHorizontal()
              ? s.touches.currentX - s.touches.startX
              : s.touches.currentY - s.touches.startY);
            if (diff < range || diff > -range) s.allowClick = true;
          },
        },
        breakpoints: {
          990: {
            spaceBetween: 8,
            direction: "vertical",
            autoHeight: false,
          },
        },
      });

      if (isUpdate) {
        setTimeout(function () {
          subSlider.update();
        }, 800);
      }
    });
  }
};

function destroyPopupSlider(modalRoot) {
  const modal = getProductMediaModalRoot(modalRoot);
  if (!modal) return;

  const popupEl = modal.querySelector(".js-popup-slider");
  const paginationEl = modal.querySelector(".product__pagination");

  if (popupEl?.swiper) {
    popupEl.swiper.destroy(true, true);
  }
  if (paginationEl) {
    paginationEl.innerHTML = "";
  }
}

const popupSliderInit = (isUpdate, modalRoot) => {
  const modal = getProductMediaModalRoot(modalRoot);
  const popupEl = modal?.querySelector(".js-popup-slider");
  if (!popupEl) return null;

  destroyPopupSlider(modal);

  const slideCount = popupEl.querySelectorAll(".swiper-slide").length;
  const paginationEl = modal.querySelector(".product__pagination");
  const nextEl = modal.querySelector(
    ".product__slider-nav .swiper-button-next"
  );
  const prevEl = modal.querySelector(
    ".product__slider-nav .swiper-button-prev"
  );

  const config = {
    slidesPerView: 1,
    on: {
      afterInit: function () {
        updatePopupSwiperPagination(this);
        if (document.querySelector(".product__outer--static-column-aside")) {
          document
            .querySelectorAll(".product__media-list .product__media-toggle")
            .forEach((elem) => {
              elem.addEventListener("click", () => {
                const mediaId = elem.getAttribute("data-media-id");
                if (mediaId) {
                  syncPopupSwiperToProductMediaId(mediaId, 0, modal);
                }
              });
            });
        }
      },
      slideChange: function () {
        window.pauseAllMedia();
        this.params.noSwiping = false;
        modal
          .querySelector(".product-media-modal__content")
          ?.classList.remove("zoom");
        updatePopupSwiperPagination(this);
      },
      touchMove: function () {
        modal
          .querySelector(".product-media-modal__content")
          ?.classList.remove("zoom");
      },
      slideChangeTransitionEnd: function () {
        updatePopupSwiperPagination(this);
        if (this.slides[this.activeIndex].querySelector("model-viewer")) {
          this.slides[this.activeIndex]
            .querySelector(".shopify-model-viewer-ui__button--poster")
            ?.removeAttribute("hidden");
        }
      },
      touchStart: function () {
        if (this.slides[this.activeIndex].querySelector("model-viewer")) {
          if (
            !this.slides[this.activeIndex]
              .querySelector("model-viewer")
              .classList.contains("shopify-model-viewer-ui__disabled")
          ) {
            this.params.noSwiping = true;
            this.params.noSwipingClass = "swiper-slide";
          } else {
            this.params.noSwiping = false;
          }
        }
      },
    },
  };

  if (nextEl && prevEl) {
    config.navigation = { nextEl, prevEl };
  }

  if (paginationEl && slideCount > 1) {
    config.pagination = {
      el: paginationEl,
      type: "bullets",
      clickable: true,
    };
  }

  const popupSlider = new Swiper(popupEl, config);

  if (isUpdate) {
    setTimeout(() => {
      popupSlider.update();
      updatePopupSwiperPagination(popupSlider);
    }, 800);
  }

  return popupSlider;
};

function slideUp(toggleEl, contentEl, duration = 300) {
  // Get the current height of the element
  const height = contentEl.clientHeight;

  // Set animation styles
  contentEl.style.transitionProperty = "height, margin, padding";
  contentEl.style.transitionDuration = duration + "ms";
  contentEl.style.overflow = "hidden";

  // Set initial values
  contentEl.style.height = height + "px";
  contentEl.style.padding = "0";
  contentEl.style.margin = "0";

  // Delay to start animation
  setTimeout(function () {
    // Set the values ​​for hiding the element
    contentEl.style.height = "0";
    contentEl.style.padding = "0";
    contentEl.style.margin = "0";
  }, 10);

  // Delay for animation to complete
  setTimeout(function () {
    // Remove installed styles after animation
    toggleEl.classList.remove("active");
    contentEl.classList.remove("active");
    contentEl.style.removeProperty("height");
    contentEl.style.removeProperty("padding");
    contentEl.style.removeProperty("margin");
    contentEl.style.removeProperty("overflow");
    contentEl.style.removeProperty("transition-duration");
    contentEl.style.removeProperty("transition-property");
  }, duration);
}

function slideDown(toggleEl, contentEl, duration = 300) {
  toggleEl.classList.add("active");
  contentEl.classList.add("active");
  contentEl.style.overflow = "hidden";
  contentEl.style.height = "0";

  const height = contentEl.scrollHeight;

  setTimeout(function () {
    contentEl.style.height = height + "px";
  }, 10);

  setTimeout(function () {
    contentEl.style.removeProperty("overflow");
    contentEl.style.removeProperty("height");
  }, duration);
}

if (navigator.userAgent.indexOf("iPhone") > -1) {
  document
    .querySelector("[name=viewport]")
    .setAttribute(
      "content",
      "width=device-width, initial-scale=1, maximum-scale=1"
    );
}

function getFocusableElements(container) {
  if (!container) return [];
  return Array.from(
    container.querySelectorAll(
      "summary, a[href], button:enabled, [tabindex]:not([tabindex^='-']), [draggable], area, input:not([type=hidden]):enabled, select:enabled, textarea:enabled, object, iframe"
    )
  );
}

document.querySelectorAll('[id^="Details-"] summary').forEach((summary) => {
  summary.setAttribute("role", "button");
  summary.setAttribute("aria-expanded", "false");

  if (summary.nextElementSibling.getAttribute("id")) {
    summary.setAttribute("aria-controls", summary.nextElementSibling.id);
  }

  summary.addEventListener("click", (event) => {
    event.currentTarget.setAttribute(
      "aria-expanded",
      !event.currentTarget.closest("details").hasAttribute("open")
    );
  });

  if (summary.closest("header-drawer")) return;
  summary.parentElement.addEventListener("keyup", onKeyUpEscape);
});

function onKeyUpEscape(event) {
  if (event.code.toUpperCase() !== "ESCAPE") return;

  const openDetailsElement = event.target.closest("details[open]");
  if (!openDetailsElement) return;

  const summaryElement = openDetailsElement.querySelector("summary");
  openDetailsElement.removeAttribute("open");
  summaryElement.setAttribute("aria-expanded", false);
  summaryElement.focus();
}

const trapFocusHandlers = {};

function trapFocus(container, elementToFocus = container) {
  var elements = getFocusableElements(container);
  var first = elements[0];
  var last = elements[elements.length - 1];

  removeTrapFocus();

  trapFocusHandlers.focusin = (event) => {
    if (
      event.target !== container &&
      event.target !== last &&
      event.target !== first
    )
      return;

    document.addEventListener("keydown", trapFocusHandlers.keydown);
  };

  trapFocusHandlers.focusout = function () {
    document.removeEventListener("keydown", trapFocusHandlers.keydown);
  };

  trapFocusHandlers.keydown = function (event) {
    if (event.code.toUpperCase() !== "TAB") return; // If not TAB key
    // On the last focusable element and tab forward, focus the first element.
    if (event.target === last && !event.shiftKey) {
      event.preventDefault();
      first.focus();
    }

    //  On the first focusable element and tab backward, focus the last element.
    if (
      (event.target === container || event.target === first) &&
      event.shiftKey
    ) {
      event.preventDefault();
      last.focus();
    }
  };

  document.addEventListener("focusout", trapFocusHandlers.focusout);
  document.addEventListener("focusin", trapFocusHandlers.focusin);

  if (elementToFocus) elementToFocus.focus();
}

function pauseAllMedia() {
  document.querySelectorAll(".js-youtube").forEach((video) => {
    video.contentWindow.postMessage(
      '{"event":"command","func":"' + "pauseVideo" + '","args":""}',
      "*"
    );
  });
  document.querySelectorAll(".js-vimeo").forEach((video) => {
    video.contentWindow.postMessage('{"method":"pause"}', "*");
  });
  document.querySelectorAll("video").forEach((video) => video.pause());
  document.querySelectorAll("product-model").forEach((model) => {
    if (model.modelViewerUI) model.modelViewerUI.pause();
  });
}

function removeTrapFocus(elementToFocus = null) {
  document.removeEventListener("focusin", trapFocusHandlers.focusin);
  document.removeEventListener("focusout", trapFocusHandlers.focusout);
  document.removeEventListener("keydown", trapFocusHandlers.keydown);

  if (elementToFocus && !elementToFocus.classList.contains("card-focused"))
    elementToFocus.focus();
}

class QuantityInput extends HTMLElement {
  constructor() {
    super();
    this.input = this.querySelector("input");
    this.changeEvent = new Event("change", { bubbles: true });
    this.isCartItem = Boolean(this.closest(".cart-quantity"));

    this.querySelectorAll("button").forEach((button) => {
      this.setMinimumDisable();

      button.addEventListener("click", this.onButtonClick.bind(this));
    });

    var eventList = ["paste", "input"];

    for (event of eventList) {
      this.input.addEventListener(event, (e) => {
        const value = e.currentTarget.value;
        const numberRegex = this.isCartItem ? /^\d*$/ : /^0*?[1-9]\d*$/;

        if (numberRegex.test(value) || value === "") {
        } else {
          e.currentTarget.value = 1;
        }

        if (e.currentTarget.value === 1 || e.currentTarget.value === "") {
          this.previousElementSibling.classList.add("disabled");
        } else {
          this.previousElementSibling?.classList.remove("disabled");
        }
      });
    }

    this.input.addEventListener("focusout", function (e) {
      if (e.currentTarget.value === "") {
        e.currentTarget.value = 1;
      }
    });
  }

  setMinimumDisable() {
    if (this.input.value == 1) {
      this.querySelector('button[name="minus"]').classList.add("disabled");
    } else {
      this.querySelector('button[name="minus"]').classList.remove("disabled");
    }
  }

  onButtonClick(event) {
    event.preventDefault();
    const previousValue = this.input.value;

    event.target.name === "plus" ? this.input.stepUp() : this.input.stepDown();
    if (previousValue !== this.input.value)
      this.input.dispatchEvent(this.changeEvent);

    this.setMinimumDisable();
  }
}

customElements.define("quantity-input", QuantityInput);

class PricePerItem extends HTMLElement {
  connectedCallback() {
    this.tiers = [];
    try {
      this.tiers = JSON.parse(this.dataset.tiers || "[]");
    } catch (e) {
      this.tiers = [];
    }
    this.input = document.getElementById(this.dataset.quantityInput);
    this.cartQuantity = Number(this.dataset.cartQuantity || 0);

    if (!this.input || !this.tiers.length) return;

    this.onQuantityChange = this.onQuantityChange.bind(this);
    this.input.addEventListener("input", this.onQuantityChange);
    this.input.addEventListener("change", this.onQuantityChange);
    this.render();
  }

  disconnectedCallback() {
    if (!this.input) return;
    this.input.removeEventListener("input", this.onQuantityChange);
    this.input.removeEventListener("change", this.onQuantityChange);
  }

  onQuantityChange() {
    this.render();
  }

  currentTierPrice() {
    const quantity = this.cartQuantity + Number(this.input.value || 0);
    let price = this.tiers[0].price;

    this.tiers.forEach((tier) => {
      if (quantity >= tier.min) price = tier.price;
    });

    return price;
  }

  render() {
    const at = this.dataset.atLabel || "at";
    const each = this.dataset.eachLabel || "each";
    const price = formatMoney(
      this.currentTierPrice(),
      this.dataset.moneyFormat
    );
    this.textContent = `${at} ${price}/${each}`;
  }
}

if (!customElements.get("price-per-item")) {
  customElements.define("price-per-item", PricePerItem);
}

function syncStickyQuantity(sectionId) {
  if (!sectionId) return;
  const main = document.getElementById(`Quantity-${sectionId}`);
  const sticky = document.getElementById(`Quantity-Sticky-${sectionId}`);
  if (!main || !sticky) return;
  if (main.dataset.stickySyncBound === "true") return;
  main.dataset.stickySyncBound = "true";
  sticky.dataset.stickySyncBound = "true";

  const sync = (from, to) => {
    if (to.value === from.value) return;
    to.value = from.value;
    to.dispatchEvent(new Event("change", { bubbles: true }));
  };

  main.addEventListener("input", () => sync(main, sticky));
  main.addEventListener("change", () => sync(main, sticky));
  sticky.addEventListener("input", () => sync(sticky, main));
  sticky.addEventListener("change", () => sync(sticky, main));
}

function initAllStickyQuantitySync(root = document) {
  root.querySelectorAll('[id^="product-form-quantity-"]').forEach((el) => {
    syncStickyQuantity(el.id.replace("product-form-quantity-", ""));
  });
}

document.addEventListener("DOMContentLoaded", () =>
  initAllStickyQuantitySync()
);
document.addEventListener("shopify:section:load", (event) =>
  initAllStickyQuantitySync(event.target)
);

function debounce(fn, wait) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn.apply(this, args), wait);
  };
}

const serializeForm = (form) => {
  const obj = {};
  const formData = new FormData(form);
  for (const key of formData.keys()) {
    obj[key] = formData.get(key);
  }
  return JSON.stringify(obj);
};

function fetchConfig(type = "json") {
  return {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Accept: `application/${type}`,
    },
  };
}

/*
 * Shopify Common JS
 *
 */
if (typeof window.Shopify == "undefined") {
  window.Shopify = {};
}

Shopify.bind = function (fn, scope) {
  return function () {
    return fn.apply(scope, arguments);
  };
};

Shopify.setSelectorByValue = function (selector, value) {
  for (var i = 0, count = selector.options.length; i < count; i++) {
    var option = selector.options[i];
    if (value == option.value || value == option.innerHTML) {
      selector.selectedIndex = i;
      return i;
    }
  }
};

Shopify.addListener = function (target, eventName, callback) {
  target.addEventListener
    ? target.addEventListener(eventName, callback, false)
    : target.attachEvent("on" + eventName, callback);
};

Shopify.postLink = function (path, options) {
  options = options || {};
  var method = options["method"] || "post";
  var params = options["parameters"] || {};

  var form = document.createElement("form");
  form.setAttribute("method", method);
  form.setAttribute("action", path);

  for (var key in params) {
    var hiddenField = document.createElement("input");
    hiddenField.setAttribute("type", "hidden");
    hiddenField.setAttribute("name", key);
    hiddenField.setAttribute("value", params[key]);
    form.appendChild(hiddenField);
  }
  document.body.appendChild(form);
  form.submit();
  document.body.removeChild(form);
};

Shopify.CountryProvinceSelector = function (
  country_domid,
  province_domid,
  options
) {
  this.countryEl = document.getElementById(country_domid);
  this.provinceEl = document.getElementById(province_domid);
  this.provinceContainer = document.getElementById(
    options["hideElement"] || province_domid
  );

  Shopify.addListener(
    this.countryEl,
    "change",
    Shopify.bind(this.countryHandler, this)
  );

  this.initCountry();
  this.initProvince();
};

Shopify.CountryProvinceSelector.prototype = {
  initCountry: function () {
    var value = this.countryEl.getAttribute("data-default");
    Shopify.setSelectorByValue(this.countryEl, value);
    this.countryHandler();
  },

  initProvince: function () {
    var value = this.provinceEl.getAttribute("data-default");
    if (value && this.provinceEl.options.length > 0) {
      Shopify.setSelectorByValue(this.provinceEl, value);
    }
  },

  countryHandler: function (e) {
    var opt = this.countryEl.options[this.countryEl.selectedIndex];
    var raw = opt.getAttribute("data-provinces");
    var provinces = JSON.parse(raw);

    this.clearOptions(this.provinceEl);
    if (provinces && provinces.length == 0) {
      this.provinceContainer.style.display = "none";
    } else {
      for (let i = 0; i < provinces.length; i++) {
        var opt = document.createElement("option");
        opt.value = provinces[i][0];
        opt.innerHTML = provinces[i][1];
        this.provinceEl.appendChild(opt);
      }

      this.provinceContainer.style.display = "";
    }
  },

  clearOptions: function (selector) {
    while (selector.firstChild) {
      selector.removeChild(selector.firstChild);
    }
  },

  setOptions: function (selector, values) {
    for (var i = 0, count = values.length; i < values.length; i++) {
      var opt = document.createElement("option");
      opt.value = values[i];
      opt.innerHTML = values[i];
      selector.appendChild(opt);
    }
  },
};

class MenuDrawer extends HTMLElement {
  constructor() {
    super();

    this.mainDetailsToggle = this.querySelector("details");
    const summaryElements = this.querySelectorAll("summary");
    this.addAccessibilityAttributes(summaryElements);

    this.headerWrapper = document.querySelector(".header-wrapper");
    if (this.headerWrapper) this.headerWrapper.preventHide = false;

    if (navigator.platform === "iPhone")
      document.documentElement.style.setProperty(
        "--viewport-height",
        `${window.innerHeight}px`
      );

    this.addEventListener("keyup", this.onKeyUp.bind(this));
    this.addEventListener("focusout", this.onFocusOut.bind(this));
    this.bindEvents();
  }

  bindEvents() {
    this.querySelectorAll("summary").forEach((summary) =>
      summary.addEventListener("click", this.onSummaryClick.bind(this))
    );
    this.querySelectorAll("button").forEach((button) => {
      if (this.querySelector(".toggle-scheme-button") === button) return;
      if (this.querySelector(".header__localization-button") === button) return;
      if (this.querySelector(".header__localization-lang-button") === button)
        return;
      button.addEventListener("click", this.onCloseButtonClick.bind(this));
    });
  }

  addAccessibilityAttributes(summaryElements) {
    summaryElements.forEach((element) => {
      element.setAttribute("role", "button");
      element.setAttribute("aria-expanded", false);
      element.setAttribute("aria-controls", element.nextElementSibling.id);
    });
  }

  onKeyUp(event) {
    if (event.code.toUpperCase() !== "ESCAPE") return;

    const openDetailsElement = event.target.closest("details[open]");
    if (!openDetailsElement) return;

    openDetailsElement === this.mainDetailsToggle
      ? this.closeMenuDrawer(this.mainDetailsToggle.querySelector("summary"))
      : this.closeSubmenu(openDetailsElement);
  }

  onSummaryClick(event) {
    const summaryElement = event.currentTarget;
    const detailsElement = summaryElement.parentNode;
    const isOpen = detailsElement.hasAttribute("open");

    if (detailsElement === this.mainDetailsToggle) {
      if (isOpen) event.preventDefault();
      isOpen
        ? this.closeMenuDrawer(summaryElement)
        : this.openMenuDrawer(summaryElement);
    } else {
      trapFocus(
        summaryElement.nextElementSibling,
        detailsElement.querySelector("button")
      );

      setTimeout(() => {
        detailsElement.classList.add("menu-opening");
      });
    }
  }

  openMenuDrawer(summaryElement) {
    if (this.headerWrapper) this.headerWrapper.preventHide = true;
    setTimeout(() => {
      this.mainDetailsToggle.classList.add("menu-opening");
    });
    summaryElement.setAttribute("aria-expanded", true);
    trapFocus(this.mainDetailsToggle, summaryElement);
    document.body.classList.add(`overflow-hidden-${this.dataset.breakpoint}`);
  }

  closeMenuDrawer(event, elementToFocus = false) {
    if (event !== undefined) {
      this.mainDetailsToggle.classList.remove("menu-opening");
      this.mainDetailsToggle.querySelectorAll("details").forEach((details) => {
        details.removeAttribute("open");
        details.classList.remove("menu-opening");
      });
      this.mainDetailsToggle
        .querySelector("summary")
        .setAttribute("aria-expanded", false);
      document.body.classList.remove(
        `overflow-hidden-${this.dataset.breakpoint}`
      );
      removeTrapFocus(elementToFocus);
      this.closeAnimation(this.mainDetailsToggle);
      this.header =
        this.header || document.querySelector(".shopify-section-header");
      const main = document.querySelector("main");
      if (
        main
          ?.querySelectorAll(".shopify-section")[0]
          ?.classList.contains("section--has-overlay") &&
        !this.header.classList.contains("animate")
      ) {
        this.header.classList.remove("color-background-overlay-hidden");
        this.header.classList.add("color-background-overlay");
      }

      if (this.headerWrapper) this.headerWrapper.preventHide = false;
    }
  }

  onFocusOut(event) {
    setTimeout(() => {
      if (
        this.mainDetailsToggle.hasAttribute("open") &&
        !this.mainDetailsToggle.contains(document.activeElement)
      )
        this.closeMenuDrawer();
    });
  }

  onCloseButtonClick(event) {
    const detailsElement = event.currentTarget.closest("details");
    if (detailsElement === this.mainDetailsToggle) {
      this.closeMenuDrawer(this.mainDetailsToggle.querySelector("summary"));
      return;
    }
    this.closeSubmenu(detailsElement);
  }

  closeSubmenu(detailsElement) {
    detailsElement.classList.remove("menu-opening");
    removeTrapFocus();
    this.closeAnimation(detailsElement);
  }

  closeAnimation(detailsElement) {
    let animationStart;

    const handleAnimation = (time) => {
      if (animationStart === undefined) {
        animationStart = time;
      }

      const elapsedTime = time - animationStart;

      if (elapsedTime < 400) {
        window.requestAnimationFrame(handleAnimation);
      } else {
        detailsElement.removeAttribute("open");
        if (detailsElement.closest("details[open]")) {
          trapFocus(
            detailsElement.closest("details[open]"),
            detailsElement.querySelector("summary")
          );
        }
      }
    };

    window.requestAnimationFrame(handleAnimation);
  }
}

customElements.define("menu-drawer", MenuDrawer);

class HeaderDrawer extends MenuDrawer {
  constructor() {
    super();
    this.headerWrapper = document.querySelector(".header-wrapper");
    if (this.headerWrapper) this.headerWrapper.preventHide = false;
  }

  openMenuDrawer(summaryElement) {
    if (this.headerWrapper) this.headerWrapper.preventHide = true;
    this.header =
      this.header || document.querySelector(".shopify-section-header");
    this.borderOffset =
      this.borderOffset ||
      this.closest(".header-wrapper").classList.contains(
        "header-wrapper--border-bottom"
      )
        ? 1
        : 0;

    const main = document.querySelector("main");
    if (
      main
        ?.querySelectorAll(".shopify-section")[0]
        ?.classList.contains("section--has-overlay")
    ) {
      this.header.classList.remove("color-background-overlay");
      this.header.classList.add("color-background-overlay-hidden");
    }

    setTimeout(() => {
      this.mainDetailsToggle.classList.add("menu-opening");
    });

    summaryElement.setAttribute("aria-expanded", true);
    trapFocus(this.mainDetailsToggle, summaryElement);
    document.body.classList.add(`overflow-hidden-${this.dataset.breakpoint}`);
  }
}

customElements.define("header-drawer", HeaderDrawer);

class ModalDialog extends HTMLElement {
  constructor() {
    super();
    this.addEventListener("click", (event) => {
      if (event.target.closest('[id^="ModalClose-"]')) {
        event.preventDefault();
        this.hide();
        return;
      }
      if (!this.classList.contains("media-modal") && event.target === this) {
        this.hide();
      }
    });
    this.addEventListener("keyup", (event) => {
      if (event.code.toUpperCase() === "ESCAPE") this.hide();
    });
    if (this.classList.contains("media-modal")) {
      this.addEventListener("pointerup", (event) => {
        if (
          event.pointerType === "mouse" &&
          !event.target.closest("deferred-media, product-model")
        )
          this.hide();
      });
    }
  }

  connectedCallback() {
    if (this.moved) return;
    this.moved = true;
    document.body.appendChild(this);
  }

  show(opener) {
    this.openedBy = opener;
    const popup = this.querySelector(".template-popup");
    document.body.classList.add("overflow-hidden-modal");
    this.setAttribute("open", "");
    if (popup) popup.loadContent();
    trapFocus(this, this.querySelector('[role="dialog"]'));
    window.pauseAllMedia();
  }

  hide() {
    let isOpen = false;

    this.removeAttribute("open");
    removeTrapFocus(this.openedBy);
    window.pauseAllMedia();

    document.querySelectorAll("body > quick-add-modal").forEach((el) => {
      if (el.hasAttribute("open")) {
        isOpen = true;
      }
    });

    if (!isOpen) {
      document.body.classList.remove("overflow-hidden-modal");
      document.body.dispatchEvent(new CustomEvent("modalClosed"));
    }

    const images = document.querySelector(".product-media-modal__content");

    if (images) {
      images.classList.remove("zoom");
    }
  }
}

customElements.define("modal-dialog", ModalDialog);

class ModalOpener extends HTMLElement {
  constructor() {
    super();

    const button = this.querySelector("button");

    if (!button) return;
    button.addEventListener("click", () => {
      const modal = document.querySelector(this.getAttribute("data-modal"));
      if (modal) modal.show(button);
    });
  }
}

customElements.define("modal-opener", ModalOpener);

class DeferredMedia extends HTMLElement {
  constructor() {
    super();
    this.querySelector('[id^="Deferred-Poster-"]')?.addEventListener(
      "click",
      this.loadContent.bind(this)
    );
  }

  loadContent() {
    if (!this.getAttribute("loaded")) {
      const content = document.createElement("div");
      content.appendChild(
        this.querySelector("template").content.firstElementChild.cloneNode(true)
      );

      this.setAttribute("loaded", true);
      window.pauseAllMedia();
      this.appendChild(
        content.querySelector("video, model-viewer, iframe")
      ).focus();

      if (
        this.closest(".swiper")?.swiper.slides[
          this.closest(".swiper").swiper.activeIndex
        ].querySelector("model-viewer")
      ) {
        if (
          !this.closest(".swiper")
            .swiper.slides[
              this.closest(".swiper").swiper.activeIndex
            ].querySelector("model-viewer")
            .classList.contains("shopify-model-viewer-ui__disabled")
        ) {
          this.closest(".swiper").swiper.params.noSwiping = true;
          this.closest(".swiper").swiper.params.noSwipingClass = "swiper-slide";
        }
      }
    }
  }
}

customElements.define("deferred-media", DeferredMedia);

class VariantSelects extends HTMLElement {
  static historyListenerBound = false;

  static initHistoryListener() {
    if (VariantSelects.historyListenerBound) return;
    VariantSelects.historyListenerBound = true;
    window.addEventListener("popstate", () => {
      document
        .querySelectorAll(
          'variant-selects[data-update-url="true"], variant-radios[data-update-url="true"]'
        )
        .forEach((picker) => {
          picker.onHistoryNavigation();
        });
    });
  }

  constructor() {
    super();
    this.addEventListener("change", this.onVariantChange);

    this.isHighVariantNeedUpdate = false;
    this.isCombinedListingsNeedUpdate = false;
    this._skipHistoryUpdate = false;
  }

  connectedCallback() {
    VariantSelects.initHistoryListener();
    if (this.dataset.updateUrl === "false") return;
    if (this._initialUrlSyncScheduled) return;
    this._initialUrlSyncScheduled = true;
    requestAnimationFrame(() => {
      this.syncProductFromWindowLocation({ onLoad: true });
    });
  }

  getProductPathFromUrl(url) {
    try {
      return new URL(url, window.location.origin).pathname;
    } catch {
      return url;
    }
  }

  shouldSyncFromWindowLocation() {
    if (this.dataset.updateUrl === "false") return false;

    const locationPath = window.location.pathname;
    const pickerPath = this.getProductPathFromUrl(this.dataset.url);
    if (locationPath !== pickerPath) return true;

    const urlVariantId = new URLSearchParams(window.location.search).get(
      "variant"
    );
    if (!urlVariantId) return false;

    const selectedScript = this.querySelector("[data-selected-variant]");
    if (!selectedScript) return false;

    try {
      const variant = JSON.parse(selectedScript.textContent.trim());
      return String(variant.id) !== String(urlVariantId);
    } catch {
      return true;
    }
  }

  createRequestUrlFromWindowLocation() {
    const sectionId = this.dataset.originalSection
      ? this.dataset.originalSection
      : this.dataset.section;
    const productPath = window.location.pathname;
    const search = new URLSearchParams(window.location.search);
    const variantId = search.get("variant");
    const optionValues = search.get("option_values");

    if (variantId) {
      return `${productPath}?variant=${variantId}&section_id=${sectionId}`;
    }
    if (optionValues) {
      return `${productPath}?option_values=${optionValues}&section_id=${sectionId}`;
    }
    return `${productPath}?section_id=${sectionId}`;
  }

  syncProductFromWindowLocation({ fromHistory = false, onLoad = false } = {}) {
    if (this.dataset.updateUrl === "false") return;
    if (!fromHistory && !onLoad) return;
    if (!this.shouldSyncFromWindowLocation()) return;

    const locationPath = window.location.pathname;
    const pickerPath = this.getProductPathFromUrl(this.dataset.url);

    this._skipHistoryUpdate = true;
    this.isCombinedListingsNeedUpdate = locationPath !== pickerPath;
    this.isHighVariantNeedUpdate =
      this.isCombinedListingsNeedUpdate ||
      this.dataset.isHighVariantProduct === "true" ||
      new URLSearchParams(window.location.search).has("option_values") ||
      new URLSearchParams(window.location.search).has("variant");

    this.classList.add("high-variant-loading");
    this.renderProductInfo(this.createRequestUrlFromWindowLocation());
  }

  onHistoryNavigation() {
    this.syncProductFromWindowLocation({ fromHistory: true });
  }

  onVariantChange(event) {
    if (!this.contains(event.target)) return;

    const combinedProductURL = event.target.dataset?.productUrl;
    if (combinedProductURL) {
      this._preservedCombinedListingValues =
        this.getPreservedCombinedListingValues();
    }

    this.updateOptions();
    // updateMasterId method updates currentVariant from liquid <script data-all-variations-no-high>
    this.updateMasterId();

    this.isHighVariantNeedUpdate = false;
    this.isCombinedListingsNeedUpdate = false;
    // -----
    // checking for high-variant and combined products
    // if variant not found in liquid <script data-all-variants-no-high> and product is high-variant
    if (
      (!this.currentVariant && this.dataset.isHighVariantProduct === "true") ||
      (!this.currentVariant && combinedProductURL)
    ) {
      const selectedValuesIds = this.getSelectedValuesIds();
      this.highVariantRequestUrl = this.createRequestUrl({
        selectedValuesIds: selectedValuesIds,
        combinedProductURL: combinedProductURL,
      });
      if (this.highVariantRequestUrl) {
        this.isHighVariantNeedUpdate = true;
        if (combinedProductURL) {
          this.isCombinedListingsNeedUpdate = true;
        }
      }
    }
    // -----

    if (this.isHighVariantNeedUpdate === false) {
      this.updatePickupAvailability();
      this.updateVariantStatuses();
    }
    this.resetErrorMessage();

    if (!this.currentVariant) {
      // -----
      // for high-variant products
      if (this.isHighVariantNeedUpdate) {
        this.toggleAddButton(true, window.variantStrings.unavailable);
        this.classList.add("high-variant-loading");
        this.renderProductInfo(this.highVariantRequestUrl);
        return;
      }
      // -----

      this.setUnavailable();
    } else {
      if (!this.isHighVariantNeedUpdate) {
        this.syncAddToCartButtonsFromVariant();
      }
      if (
        this.currentVariant?.featured_media &&
        this.dataset?.variantMediaDisplay === "show_all"
      ) {
        // If variant display != "show_all", the media gallery element is fully replaced inside updateElementsAfterFetch
        const mediaId = `${this.dataset.section}-${this.currentVariant.featured_media.id}`;
        this.updateMedia(mediaId);
      }
      this.updateURL();
      this.updateVariantInput();
      const requestUrl = this.createRequestUrl({
        currentVariantId: this.currentVariant.id,
      });
      this.renderProductInfo(requestUrl);
    }
  }

  updateOptions() {
    const fieldsets = Array.from(
      this.querySelectorAll(".product-form__controls--dropdown")
    );

    this.options = Array.from(
      this.querySelectorAll("select"),
      (select) => select.value
    ).concat(
      fieldsets.map((fieldset) => {
        return Array.from(fieldset.querySelectorAll("input")).find(
          (radio) => radio.checked
        ).value;
      })
    );
  }

  updateMasterId() {
    if (this.variantData || this.querySelector("[data-all-variants-no-high]")) {
      this.currentVariant = this.getVariantData().find((variant) => {
        return !variant.options
          .map((option, index) => {
            return this.options[index] === option;
          })
          .includes(false);
      });
      this.updateSelectedVariantJson();
    }
  }

  extractVariantIdFromRequestUrl(requestUrl) {
    if (!requestUrl) return null;
    try {
      const url = requestUrl.startsWith("http")
        ? new URL(requestUrl)
        : new URL(requestUrl, window.location.origin);
      return url.searchParams.get("variant");
    } catch {
      return null;
    }
  }

  resolveVariantFromFetchedHtml(html, sourceSectionId, variantId) {
    if (!html || !variantId) return null;
    const picker = html.getElementById(`variant-picker-${sourceSectionId}`);
    const allEl = picker?.querySelector("[data-all-variants-no-high]");
    if (!allEl) return null;
    try {
      const list = JSON.parse(allEl.textContent.trim());
      return (
        list.find((variant) => String(variant.id) === String(variantId)) || null
      );
    } catch {
      return null;
    }
  }

  updateSelectedVariantJson() {
    const el = this.querySelector("[data-selected-variant]");
    if (!el) return;
    if (this.currentVariant) {
      el.textContent = JSON.stringify(this.currentVariant);
    } else {
      el.textContent = "null";
    }
  }

  getProductAddToCartForms() {
    const scope = this.getProductScopeRoot();
    const sectionId = this.dataset.section;
    return [
      scope.querySelector(`#product-form-${sectionId}`),
      scope.querySelector(`#product-form-installment-${sectionId}`),
      scope.querySelector(`#sticky-bar-product-form-${sectionId}`),
    ].filter(Boolean);
  }

  ensureAddButtonLabelElement(addButton) {
    if (!addButton) return null;

    let content = addButton.querySelector(".button__content");
    if (!content) {
      content = document.createElement("div");
      content.className = "button__content";
      const spinner = addButton.querySelector(".loading-overlay__spinner");
      const circle = addButton.querySelector(".button_circle");
      if (spinner) {
        addButton.insertBefore(content, spinner);
      } else if (circle) {
        addButton.insertBefore(content, circle);
      } else {
        addButton.prepend(content);
      }
    }

    let label = content.querySelector(".button__label");
    if (!label) {
      label = content.querySelector(":scope > span:not(.visually-hidden)");
    }
    if (!label) {
      label = document.createElement("span");
      label.className = "button__label";
      content.appendChild(label);
    } else if (!label.classList.contains("button__label")) {
      label.classList.add("button__label");
    }

    return label;
  }

  applyAddButtonState(addButton, productForm, { disabled, labelText, status }) {
    if (!addButton) return;

    const labelEl = this.ensureAddButtonLabelElement(addButton);
    if (labelEl && labelText != null && labelText !== "") {
      labelEl.textContent = labelText;
      labelEl.setAttribute("data-label", labelText);
    }

    if (disabled) {
      addButton.setAttribute("disabled", "");
      addButton.setAttribute("aria-disabled", "true");
      if (status) addButton.dataset.status = status;
    } else {
      addButton.removeAttribute("disabled");
      addButton.removeAttribute("aria-disabled");
      addButton.dataset.status = "available";
      if (labelText) {
        addButton.setAttribute("data-button-text", labelText);
      } else {
        const fromForm = productForm?.getAttribute?.("id") || "";
        if (fromForm.includes("sticky-bar-product-form")) {
          addButton.setAttribute(
            "data-button-text",
            this.getAddButtonEnableLabel(addButton, productForm)
          );
        }
      }
    }
  }

  isHidden(elem) {
    const styles = window.getComputedStyle(elem);
    return styles.display === "none" || styles.visibility === "hidden";
  }

  getProductScopeRoot() {
    return (
      this.closest("quick-add-modal") ||
      this.closest("quick-view-modal") ||
      this.closest("#quick-view") ||
      this.closest('[id^="shopify-section-"]') ||
      document.documentElement
    );
  }

  destroyMediaSwipersWithin(root) {
    if (!root || !root.querySelectorAll) return;
    root.querySelectorAll(".js-media-list, .js-media-sublist").forEach((el) => {
      if (el.swiper && typeof el.swiper.destroy === "function") {
        el.swiper.destroy(true, true);
      }
    });
  }

  getMediaMountRoot() {
    const sectionId = this.dataset.section;
    const scope = this.getProductScopeRoot();
    return (
      scope.querySelector(
        `.product__media-wrapper[data-section="product-media-${sectionId}"]`
      ) ||
      scope.querySelector(`[data-section="product-media-${sectionId}"]`) ||
      scope
    );
  }

  initProductMediaSwipersInRoot(root, { animate = false } = {}) {
    if (!root?.querySelectorAll) return;

    const speed = animate ? 800 : 0;

    root.querySelectorAll(".js-media-sublist").forEach((elem) => {
      if (elem.swiper?.destroy) {
        elem.swiper.destroy(true, true);
      }

      const subSlider = new Swiper(elem, {
        slidesPerView: 5,
        spaceBetween: 8,
        direction: "horizontal",
        freeMode: false,
        watchSlidesProgress: true,
        autoHeight: false,
        speed,
        initialSlide: 0,
        on: {
          touchEnd(s, e) {
            const range = 5;
            const diff = (s.touches.diff = s.isHorizontal()
              ? s.touches.currentX - s.touches.startX
              : s.touches.currentY - s.touches.startY);
            if (diff < range || diff > -range) s.allowClick = true;
          },
        },
        breakpoints: {
          990: {
            spaceBetween: 8,
            direction: "vertical",
            autoHeight: false,
          },
        },
      });

      elem._variantSelectsSubSwiper = subSlider;
    });

    root.querySelectorAll(".js-media-list").forEach((elem) => {
      if (elem.swiper?.destroy) {
        elem.swiper.destroy(true, true);
      }

      const mediaListId = elem.dataset?.jsMediaListId;
      const mediaSublist = mediaListId
        ? root.querySelector(
            `.js-media-sublist[data-js-media-list-id="${mediaListId}"]`
          )
        : null;
      const thumbsSwiper =
        mediaSublist?._variantSelectsSubSwiper || mediaSublist?.swiper || null;
      const productRoot = elem.closest(".product") || root;

      new Swiper(elem, {
        slidesPerView: 1,
        spaceBetween: 1,
        autoHeight: false,
        speed,
        initialSlide: 0,
        navigation: {
          nextEl: [
            productRoot.querySelector(
              ".product__slider-nav .swiper-button-next"
            ),
            productRoot.querySelector(
              ".product__gallery-arrows .swiper-button-next"
            ),
          ].filter(Boolean),
          prevEl: [
            productRoot.querySelector(
              ".product__slider-nav .swiper-button-prev"
            ),
            productRoot.querySelector(
              ".product__gallery-arrows .swiper-button-prev"
            ),
          ].filter(Boolean),
        },
        pagination: {
          el: productRoot.querySelector(".product__pagination"),
          type: "bullets",
          clickable: true,
        },
        thumbs: {
          swiper: thumbsSwiper,
        },
        on: {
          slideChangeTransitionStart() {
            if (thumbsSwiper && !thumbsSwiper.destroyed) {
              thumbsSwiper.slideTo(this.activeIndex, speed);
            }
          },
          slideChange() {
            window.pauseAllMedia?.();
            this.params.noSwiping = false;

            syncPopupSwiperFromMainGallery(this, 0);
          },
          slideChangeTransitionEnd() {
            if (this.slides[this.activeIndex]?.querySelector("model-viewer")) {
              this.slides[this.activeIndex]
                .querySelector(".shopify-model-viewer-ui__button--poster")
                ?.removeAttribute("hidden");
            }
          },
          touchStart() {
            const modelViewer =
              this.slides[this.activeIndex]?.querySelector("model-viewer");
            if (modelViewer) {
              if (
                !modelViewer.classList.contains(
                  "shopify-model-viewer-ui__disabled"
                )
              ) {
                this.params.noSwiping = true;
                this.params.noSwipingClass = "swiper-slide";
              } else {
                this.params.noSwiping = false;
              }
            }
          },
        },
      });
    });
  }

  reinitProductMediaAfterDomSwap(root) {
    const mountRoot = root || this.getMediaMountRoot();
    this.initProductMediaSwipersInRoot(mountRoot, { animate: false });
  }

  updateMedia(mediaId) {
    if (!mediaId) return;

    const scope = this.getProductScopeRoot();
    const mediaEl =
      scope.querySelector(
        `.product__media-wrapper[data-section="product-media-${this.dataset.section}"]`
      ) ||
      scope.querySelector(
        `[data-section="product-media-${this.dataset.section}"]`
      );
    if (!mediaEl) return;

    // for slider and slider_previews layout
    const sliderEl = mediaEl.querySelector(".js-media-list");
    if (sliderEl && sliderEl.swiper) {
      sliderEl.querySelectorAll(".swiper-slide").forEach((slide, index) => {
        const slideMediaId = slide.dataset?.mediaId;
        if (slideMediaId === mediaId) {
          const findIndex = sliderEl.swiper.slides.findIndex((slideEl) => {
            return slideEl.dataset?.mediaId === slideMediaId;
          });
          const fallbackIndex = Number(
            slide.dataset?.swiperSlideIndex || index
          );
          const swiperSlideIndex = findIndex !== -1 ? findIndex : fallbackIndex;
          sliderEl.swiper.slideTo(swiperSlideIndex, 800);
          const targetSlide = sliderEl.swiper.slides[swiperSlideIndex];
          const numericMediaId = targetSlide
            ?.querySelector(".product__media-toggle")
            ?.getAttribute("data-media-id");
          syncPopupSwiperToProductMediaId(numericMediaId, 0);
        }
      });
    }

    // for stacked_previews layout
    const stackedEl =
      mediaEl.querySelector(
        ".product__media-list[data-desktop-type='stacked_previews']"
      ) ||
      mediaEl.querySelector(
        ".product__media-list[data-desktop-type='stacked']"
      );
    if (stackedEl && window.innerWidth >= 990) {
      stackedEl.querySelectorAll("[data-media-id]").forEach((mediaItem) => {
        const mediaItemId = mediaItem.dataset?.mediaId;
        if (mediaItemId === mediaId) {
          const offset = mediaItem.getBoundingClientRect().top + window.scrollY;
          window.scrollTo({ top: offset - 100, behavior: "smooth" });
        }
      });
    }
  }

  updateURL({ replace = false } = {}) {
    if (this.dataset.updateUrl === "false" || this._skipHistoryUpdate) return;
    const newUrl = this.currentVariant
      ? `${this.dataset.url}?variant=${this.currentVariant.id}`
      : this.dataset.url;

    const state = { variantSelects: this.dataset.section };
    if (replace) {
      window.history.replaceState(state, "", newUrl);
    } else {
      window.history.pushState(state, "", newUrl);
    }
  }

  updateVariantInput() {
    const formSelector = `#product-form-${this.dataset.section}, #product-form-installment-${this.dataset.section}, #sticky-bar-product-form-${this.dataset.section}`;
    const productForms = document.querySelectorAll(formSelector);
    productForms.forEach((productForm) => {
      const input = productForm.querySelector('input[name="id"]');
      if (!input) return;
      input.value = this.currentVariant.id;
      input.disabled = false;
      input.dispatchEvent(new Event("change", { bubbles: true }));
    });

    this.syncStickyAddToCartFeaturedImage();

    publish(PUB_SUB_EVENTS.variantChange, {
      data: {
        sectionId: this.dataset.section,
        variant: this.currentVariant,
      },
    });
  }

  updateVariantStatuses() {
    if (this.dataset.isHighVariantProduct === "true") {
      return;
    }

    const variantData = this.getVariantData();
    if (!variantData.length) return;

    const inputWrappers = [
      ...this.querySelectorAll(".product-form__controls[data-option-position]"),
    ];
    inputWrappers.sort(
      (a, b) =>
        Number(a.dataset.optionPosition) - Number(b.dataset.optionPosition)
    );

    const selectedValues = this.getSelectedValues();
    if (
      !selectedValues.length ||
      selectedValues.length !== inputWrappers.length ||
      selectedValues.some((s) => s == null || s === "")
    ) {
      return;
    }

    inputWrappers.forEach((option, index) => {
      const optionInputs = [
        ...option.querySelectorAll('input[type="radio"], option'),
      ];
      if (!optionInputs.length) return;

      const availableOptionInputsValue = variantData
        .filter(
          (variant) =>
            variant.available &&
            selectedValues.every(
              (selected, j) =>
                j === index || String(variant.options[j]) === String(selected)
            )
        )
        .map((variantOption) => variantOption.options[index]);

      this.setInputAvailability(optionInputs, availableOptionInputsValue);
    });
  }

  setInputAvailability(listOfOptions, listOfAvailableOptions) {
    listOfOptions.forEach((input) => {
      if (input.dataset?.productUrl) {
        return;
      }

      const rawVal = input.getAttribute("value");
      const isAvail = listOfAvailableOptions.some(
        (v) => String(v) === String(rawVal)
      );
      if (isAvail) {
        if (input.tagName === "OPTION") {
          input.innerText = input.getAttribute("value");
        } else if (input.tagName === "INPUT") {
          input.classList.remove("disabled");
        }
      } else {
        if (input.tagName === "OPTION") {
          input.innerText =
            window.variantStrings.unavailable_with_option.replace(
              "[value]",
              input.getAttribute("value")
            );
        } else if (input.tagName === "INPUT") {
          input.classList.add("disabled");
        }
      }
    });
  }

  getSortedOptionControls() {
    const controls = [
      ...this.querySelectorAll(".product-form__controls[data-option-position]"),
    ];

    controls.sort((a, b) => {
      return (
        Number(a.dataset.optionPosition) - Number(b.dataset.optionPosition)
      );
    });

    return controls;
  }

  normalizeOptionValue(value) {
    return String(value ?? "")
      .trim()
      .toLowerCase();
  }

  setCheckedInputsBySelectedValues(selectedValues) {
    const inputWrappers = this.getSortedOptionControls();

    inputWrappers.forEach((groupEl, index) => {
      const selectedValue = selectedValues[index];
      if (!selectedValue) return;

      const inputs = [
        ...groupEl.querySelectorAll('input[type="radio"]'),
      ].filter((input) => !input.dataset.productUrl);
      const normalizedSelectedValue = this.normalizeOptionValue(selectedValue);
      const matchedInput =
        inputs.find(
          (input) =>
            !input.classList.contains("disabled") &&
            !input.disabled &&
            this.normalizeOptionValue(input.value) === normalizedSelectedValue
        ) ||
        inputs.find(
          (input) =>
            this.normalizeOptionValue(input.value) === normalizedSelectedValue
        );
      if (!matchedInput) return;

      inputs.forEach((input) => {
        const shouldBeChecked = input === matchedInput;
        input.checked = shouldBeChecked;
        if (shouldBeChecked) {
          input.setAttribute("checked", "");
        } else {
          input.removeAttribute("checked");
        }
      });
    });
  }

  getSelectedValues() {
    const selectedValues = this.getSortedOptionControls().map((control) => {
      const checkedInput = control.querySelector('input[type="radio"]:checked');
      return checkedInput ? checkedInput.value : null;
    });

    return selectedValues;
  }

  getPreservedCombinedListingValues() {
    return this.getSortedOptionControls().map((control) => {
      const hasCombinedListingInputs = control.querySelector(
        'input[type="radio"][data-product-url]'
      );
      if (!hasCombinedListingInputs) return null;

      const checkedInput = control.querySelector('input[type="radio"]:checked');
      return checkedInput ? checkedInput.value : null;
    });
  }

  restoreCombinedListingSelections() {
    if (!this._preservedCombinedListingValues?.length) return;

    this.setCheckedInputsBySelectedValues(this._preservedCombinedListingValues);
    delete this._preservedCombinedListingValues;

    this.updateOptions();
    this.updateMasterId();
    this.updateSelectedVariantJson();

    try {
      this.updateVariantStatuses();
    } catch (err) {}
  }

  updatePickupAvailability() {
    const pickUpAvailability = document.querySelector("pickup-availability");
    if (!pickUpAvailability) return;

    if (this.currentVariant && this.currentVariant.available) {
      pickUpAvailability.fetchAvailability(this.currentVariant.id);
    } else {
      pickUpAvailability.removeAttribute("available");
      pickUpAvailability.innerHTML = "";
    }
  }

  renderProductInfo(requestUrl) {
    this.abortController?.abort();
    this.abortController = new AbortController();

    fetch(requestUrl, { signal: this.abortController.signal })
      .then((response) => response.text())
      .then((responseText) => {
        const html = new DOMParser().parseFromString(responseText, "text/html");

        try {
          this.setCurrentVariantAfterFetch(html, requestUrl);
        } catch (err) {}

        let refreshedPickerFromSection = false;

        // -----
        // for high-variant products
        // and if variant not found in liquid <script data-all-variants-no-high>
        // but it was found after a request with the option_values parameter
        if (this.isHighVariantNeedUpdate) {
          try {
            this.updatePickerInnerHtml(html);
            if (this.isCombinedListingsNeedUpdate) {
              this.restoreCombinedListingSelections();
            }
            refreshedPickerFromSection = true;
            this.updateURL();
            this.updatePickupAvailability();
            if (this.currentVariant) {
              this.updateVariantInput();
              if (
                this.currentVariant.featured_media &&
                this.dataset?.variantMediaDisplay === "show_all" &&
                !this.isCombinedListingsNeedUpdate
              ) {
                // If variant display != "show_all", the media gallery element is fully replaced inside updateElementsAfterFetch
                const mediaId = `${this.dataset.section}-${this.currentVariant.featured_media.id}`;
                this.updateMedia(mediaId);
              }
            }
          } catch (err) {}
        } else if (
          this.dataset.isHighVariantProduct === "true" &&
          this.currentVariant
        ) {
          // Liquid caps variant JSON in the theme; client-side availability math is wrong for
          // high-variant catalogs. Re-render the picker from the section response (authoritative).
          try {
            this.updatePickerInnerHtml(html);
            refreshedPickerFromSection = true;
          } catch (err) {}
        }
        // -----

        this.updateElementsAfterFetch(html);

        if (
          this.currentVariant &&
          !refreshedPickerFromSection &&
          this.dataset.isHighVariantProduct !== "true" &&
          !this.isCombinedListingsNeedUpdate
        ) {
          try {
            this.updateVariantStatuses();
          } catch (err) {}
        }

        this.syncAddToCartButtonsFromVariant();
      })
      .catch((error) => {
        if (error.name === "AbortError") {
          console.info("Fetch aborted by user");
        } else {
          console.error(error);
        }
      })
      .finally(() => {
        this.classList.remove("high-variant-loading");
        this._skipHistoryUpdate = false;
        if (this.dataset.updateUrl !== "false") {
          try {
            this.dataset.url = window.location.pathname;
          } catch (e) {}
        }
      });
  }

  syncStickyAddToCartFeaturedImage() {
    if (!this.currentVariant) return;
    const scope = this.getProductScopeRoot();
    const wrap = scope.querySelector(
      ".sticky-add-to-cart .product-card-mini__image"
    );
    if (!wrap) return;
    const preview =
      this.currentVariant.featured_media?.preview_image ||
      this.currentVariant.featured_image ||
      this.currentVariant.featured_media;
    if (!preview) return;
    const src = preview.src || preview.url;
    if (!src) return;
    let img = wrap.querySelector("img");
    if (!img) {
      img = document.createElement("img");
      img.loading = "lazy";
      wrap.appendChild(img);
    }
    img.src = src;
    if (preview.alt) img.alt = preview.alt;
    const w = preview.width;
    const h = preview.height;
    if (w) img.setAttribute("width", String(w));
    if (h) img.setAttribute("height", String(h));
  }

  getAddButtonEnableLabel(addButton, productForm) {
    if (this.isCurrentVariantPreOrder()) {
      return window.variantStrings.preOrder;
    }

    // form.id is shadowed by <input name="id"> — use getAttribute
    const formId = productForm?.getAttribute?.("id") || "";
    const isSticky = formId.includes("sticky-bar-product-form");
    if (isSticky) {
      const fromButton = addButton?.getAttribute("data-button-text");
      if (fromButton) return fromButton;

      const existing = this.ensureAddButtonLabelElement(addButton);
      const fromLabel = existing?.getAttribute("data-label");
      if (
        fromLabel &&
        fromLabel !== window.variantStrings.soldOut &&
        fromLabel !== window.variantStrings.unavailable
      ) {
        return fromLabel;
      }
    }

    return window.variantStrings.addToCart;
  }

  isCurrentVariantPreOrder() {
    const variant = this.currentVariant;
    if (!variant || !window.theme?.preOrderEnabled) return false;
    return Boolean(
      variant.available &&
        variant.inventory_management === "shopify" &&
        variant.inventory_policy === "continue" &&
        Number(variant.inventory_quantity) <= 0
    );
  }

  syncAddToCartButtonsFromVariant() {
    if (!this.currentVariant) {
      this.setUnavailable();
      return;
    }

    const scope = this.getProductScopeRoot();
    const price = scope.querySelector(`#price-${this.dataset.section}`);
    const priceSticky = scope.querySelector(
      `#price-sticky-${this.dataset.section}`
    );
    const inventory = scope.querySelector(`#Inventory-${this.dataset.section}`);
    const stockBar = this.getStockBarElement(scope);
    const sku = scope.querySelector(`#Sku-${this.dataset.section}`);
    price?.classList.remove("visibility-hidden");
    priceSticky?.classList.remove("visibility-hidden");
    inventory?.classList.remove("visibility-hidden");
    stockBar?.classList.remove("visibility-hidden");
    sku?.classList.remove("visibility-hidden");
    scope
      .querySelectorAll(`[id^="ColorName-${this.dataset.section}"]`)
      .forEach((el) => el.classList.remove("visibility-hidden"));

    const available = Boolean(this.currentVariant.available);
    const isPreOrder = this.isCurrentVariantPreOrder();

    this.getProductAddToCartForms().forEach((productForm) => {
      const addButton = productForm.querySelector('[name="add"]');
      if (!addButton) return;

      if (isPreOrder) {
        this.applyAddButtonState(addButton, productForm, {
          disabled: false,
          labelText: window.variantStrings.preOrder,
          status: "pre-order",
        });
      } else if (available) {
        this.applyAddButtonState(addButton, productForm, {
          disabled: false,
          labelText: this.getAddButtonEnableLabel(addButton, productForm),
          status: "available",
        });
      } else {
        this.applyAddButtonState(addButton, productForm, {
          disabled: true,
          labelText: window.variantStrings.soldOut,
          status: "sold-out",
        });
      }
    });

    this.updateStockBarFromVariant();
  }

  syncAddButtonsWithCurrentVariant() {
    this.syncAddToCartButtonsFromVariant();
  }

  toggleAddButton(disable = true, text) {
    const labelText =
      text != null && text !== ""
        ? text
        : disable
        ? this.currentVariant
          ? window.variantStrings.soldOut
          : window.variantStrings.unavailable
        : window.variantStrings.addToCart;

    const status = disable
      ? labelText === window.variantStrings.unavailable
        ? "unavailable"
        : "sold-out"
      : "available";

    this.getProductAddToCartForms().forEach((productForm) => {
      const addButton = productForm.querySelector('[name="add"]');
      if (!addButton) return;

      if (disable) {
        this.applyAddButtonState(addButton, productForm, {
          disabled: true,
          labelText,
          status,
        });
      } else {
        this.applyAddButtonState(addButton, productForm, {
          disabled: false,
          labelText: this.getAddButtonEnableLabel(addButton, productForm),
          status,
        });
      }
    });
  }

  resetErrorMessage() {
    const productForms = document.querySelectorAll(
      `#product-form-${this.dataset.section}, #product-form-installment-${this.dataset.section}, #sticky-bar-product-form-${this.dataset.section}`
    );
    productForms.forEach((productForm) => {
      const parentEl = productForm.closest("product-form");
      if (parentEl) {
        const errorWrapperEl = parentEl.querySelector(
          ".product-form__error-message-wrapper"
        );
        const errorTextEl = errorWrapperEl?.querySelector(
          ".product-form__error-message"
        );
        if (!errorWrapperEl || !errorTextEl) return;
        errorWrapperEl.setAttribute("hidden", true);
        errorTextEl.textContent = "";
      }
    });
  }

  setUnavailable() {
    const price = document.getElementById(`price-${this.dataset.section}`);
    const priceSticky = document.getElementById(
      `price-sticky-${this.dataset.section}`
    );
    const scope = this.getProductScopeRoot();
    const inventory = scope.querySelector(`#Inventory-${this.dataset.section}`);
    const stockBar = this.getStockBarElement(scope);
    const sku = scope.querySelector(`#Sku-${this.dataset.section}`);
    const colorNameDestinations = scope.querySelectorAll(
      `[id^="ColorName-${this.dataset.section}"]`
    );

    this.toggleAddButton(true, window.variantStrings.unavailable);
    if (price) price.classList.add("visibility-hidden");
    if (priceSticky) priceSticky.classList.add("visibility-hidden");
    if (inventory) inventory.classList.add("visibility-hidden");
    if (stockBar) stockBar.classList.add("visibility-hidden");
    if (sku) sku.classList.add("visibility-hidden");
    colorNameDestinations.forEach((colorNameDestination) => {
      colorNameDestination.classList.add("visibility-hidden");
    });
  }

  getStockBarElement(scope = this.getProductScopeRoot()) {
    return scope?.querySelector(`#StockBar-${this.dataset.section}`);
  }

  updateStockBarFromVariant() {
    const stockBar = this.getStockBarElement();
    const variant = this.currentVariant;
    if (!stockBar || !variant) return;

    if (variant.inventory_management !== "shopify") {
      stockBar.classList.add("visibility-hidden");
      return;
    }

    stockBar.classList.remove("visibility-hidden");

    const qty = Number(variant.inventory_quantity) || 0;
    const maxStock = Number(stockBar.dataset.maxStock) || 100;
    let percent = Math.round((qty * 100) / maxStock);
    percent = Math.max(0, Math.min(100, percent));

    stockBar.classList.remove("green", "yellow", "red");
    if (qty <= 0) {
      stockBar.classList.add("red");
    } else if (percent > 66) {
      stockBar.classList.add("green");
    } else if (percent > 33) {
      stockBar.classList.add("yellow");
    } else {
      stockBar.classList.add("red");
    }

    const bar = stockBar.querySelector(".stock-bar");
    const fill = stockBar.querySelector(".stock-bar-fill");
    const oosMessage = stockBar.querySelector(".message__out-of-stock");

    // In-stock vs out-of-stock use different markup; partial update only when structure matches.
    if (qty > 0) {
      if (!bar) return;
      if (oosMessage) oosMessage.remove();
      if (fill) fill.style.width = `${percent}%`;
    } else if (!oosMessage) {
      return;
    }
  }

  getVariantData() {
    const el = this.querySelector("[data-all-variants-no-high]");
    if (!el) {
      return this.variantData || [];
    }
    try {
      this.variantData = JSON.parse(el.textContent.trim());
    } catch {
      this.variantData = [];
    }
    return this.variantData;
  }

  cloneModalSubtreeForSection(el, currentSectionId, sourceSectionId) {
    if (!el) return null;
    const wrap = document.createElement("div");
    let markup = el.outerHTML;
    if (currentSectionId !== sourceSectionId) {
      markup = markup.replaceAll(sourceSectionId, currentSectionId);
    }
    wrap.innerHTML = markup;
    return wrap.firstElementChild;
  }

  replaceProductModalChromePart(
    srcModal,
    destModal,
    selector,
    insertAfter,
    currentSectionId,
    sourceSectionId
  ) {
    const srcPart = srcModal.querySelector(selector);
    const destPart = destModal.querySelector(selector);
    const nextPart = this.cloneModalSubtreeForSection(
      srcPart,
      currentSectionId,
      sourceSectionId
    );

    if (destPart) {
      if (nextPart) destPart.replaceWith(nextPart);
      else destPart.remove();
    } else if (nextPart && insertAfter) {
      insertAfter.insertAdjacentElement("afterend", nextPart);
    }
  }

  syncProductMediaModalFromFetchedHtml(
    html,
    currentSectionId,
    sourceSectionId
  ) {
    const srcModal = html.getElementById(`ProductModal-${sourceSectionId}`);
    const destModal = document.getElementById(
      `ProductModal-${currentSectionId}`
    );
    if (!srcModal || !destModal) return;

    const srcDialog = srcModal.querySelector(".product-media-modal__dialog");
    const destDialog = destModal.querySelector(".product-media-modal__dialog");
    if (!srcDialog || !destDialog) return;

    destroyPopupSlider(destModal);

    let markup = srcDialog.innerHTML;
    if (currentSectionId !== sourceSectionId) {
      markup = markup.replaceAll(sourceSectionId, currentSectionId);
    }
    destDialog.innerHTML = markup;

    this.replaceProductModalChromePart(
      srcModal,
      destModal,
      ".product__pagination-wrapper",
      destDialog,
      currentSectionId,
      sourceSectionId
    );

    const paginationAnchor =
      destModal.querySelector(".product__pagination-wrapper") || destDialog;

    this.replaceProductModalChromePart(
      srcModal,
      destModal,
      ".swipper-arrows-parent",
      paginationAnchor,
      currentSectionId,
      sourceSectionId
    );

    if (typeof popupSliderInit === "function") {
      popupSliderInit(false, destModal);
    }

    destModal.querySelectorAll(".js-image-zoom").forEach((img) => {
      if (img.dataset.zoomBound === "true") return;
      img.dataset.zoomBound = "true";
      img.addEventListener("click", () => {
        destDialog
          .querySelector(".product-media-modal__content")
          ?.classList.toggle("zoom");
      });
    });
  }

  updateElementsAfterFetch(html) {
    const currentSectionId = this.dataset.section;
    const sourceSectionId = this.dataset.originalSection
      ? this.dataset.originalSection
      : this.dataset.section;
    const scope = this.getProductScopeRoot();

    const byId = (root, id) => root.querySelector(`#${CSS.escape(id)}`);

    const cloneSubtreeForCurrentSection = (el) => {
      if (!el) return null;
      const wrap = document.createElement("div");
      let markup = el.outerHTML;
      if (currentSectionId !== sourceSectionId) {
        markup = markup.replaceAll(sourceSectionId, currentSectionId);
      }
      wrap.innerHTML = markup;
      return wrap.firstElementChild;
    };

    // price
    const priceDestination = byId(scope, `price-${currentSectionId}`);
    const priceSource = html.getElementById(`price-${sourceSectionId}`);

    if (priceSource && priceDestination) {
      priceDestination.innerHTML = priceSource.innerHTML;
      priceDestination.classList.remove("visibility-hidden");
    }

    const priceStickyDestination = byId(
      scope,
      `price-sticky-${currentSectionId}`
    );
    const priceStickySource = html.getElementById(
      `price-sticky-${sourceSectionId}`
    );
    if (priceStickySource && priceStickyDestination) {
      priceStickyDestination.innerHTML = priceStickySource.innerHTML;
      priceStickyDestination.classList.remove("visibility-hidden");
    }

    // B2B elements (volume pricing table + quantity rules)
    const b2bDestination = byId(scope, `product-b2b-${currentSectionId}`);
    const b2bSource = html.getElementById(`product-b2b-${sourceSectionId}`);
    if (b2bSource && b2bDestination) {
      b2bDestination.innerHTML = b2bSource.innerHTML;
    }

    // Quantity input (rebinds min/max/step/value + price-per-item tiers)
    const qtyDestination = byId(
      scope,
      `product-form-quantity-${currentSectionId}`
    );
    const qtySource = html.getElementById(
      `product-form-quantity-${sourceSectionId}`
    );
    if (qtySource && qtyDestination) {
      qtyDestination.innerHTML = qtySource.innerHTML;
    }

    // Mirror quantity rule attrs (min/max/step) onto the sticky bar's quantity input
    const mainQtyInput = byId(scope, `Quantity-${currentSectionId}`);
    const stickyQtyInput = byId(scope, `Quantity-Sticky-${currentSectionId}`);
    if (mainQtyInput && stickyQtyInput) {
      ["min", "max", "step"].forEach((attr) => {
        if (mainQtyInput.hasAttribute(attr)) {
          stickyQtyInput.setAttribute(attr, mainQtyInput.getAttribute(attr));
        } else {
          stickyQtyInput.removeAttribute(attr);
        }
      });
      if (Number(stickyQtyInput.value) < Number(mainQtyInput.min || 0)) {
        stickyQtyInput.value = mainQtyInput.min;
      }
    }

    const stickyVariantDest = byId(scope, `StickyVariant-${currentSectionId}`);
    const stickyVariantSrc = html.getElementById(
      `StickyVariant-${sourceSectionId}`
    );
    if (stickyVariantSrc && stickyVariantDest) {
      stickyVariantDest.textContent = stickyVariantSrc.textContent;
      stickyVariantDest.className = stickyVariantSrc.className;
    }

    // stock bar + inventory text
    const stockBarSource = html.getElementById(`StockBar-${sourceSectionId}`);
    const stockBarDestination = byId(scope, `StockBar-${currentSectionId}`);
    const inventoryDestination = byId(scope, `Inventory-${currentSectionId}`);

    if (stockBarDestination) {
      if (stockBarSource) {
        const nextStockBar = cloneSubtreeForCurrentSection(stockBarSource);
        if (nextStockBar) stockBarDestination.replaceWith(nextStockBar);
      } else {
        stockBarDestination.remove();
      }
    } else if (stockBarSource && inventoryDestination) {
      const nextStockBar = cloneSubtreeForCurrentSection(stockBarSource);
      if (nextStockBar) {
        inventoryDestination.insertAdjacentElement("beforebegin", nextStockBar);
      }
    }

    const inventorySource = html.getElementById(`Inventory-${sourceSectionId}`);
    if (inventorySource && inventoryDestination) {
      inventoryDestination.innerHTML = inventorySource.innerHTML;
      inventoryDestination.classList.toggle(
        "visibility-hidden",
        inventorySource.innerText === ""
      );
    }

    this.updateStockBarFromVariant();

    // sku
    const skuSource = html.getElementById(`Sku-${sourceSectionId}`);
    const skuDestination = byId(scope, `Sku-${currentSectionId}`);
    if (skuDestination) {
      if (skuSource) {
        skuDestination.innerHTML = skuSource.innerHTML;
        const hideSku =
          skuSource.classList.contains("visibility-hidden") ||
          skuSource.innerText.trim() === "";
        skuDestination.classList.toggle("visibility-hidden", hideSku);
      } else {
        skuDestination.innerHTML = "";
        skuDestination.classList.add("visibility-hidden");
      }
    }

    // color swatches label
    const colorNameSources = html.querySelectorAll(
      `[id^="ColorName-${sourceSectionId}"]`
    );
    const colorNameDestinations = scope.querySelectorAll(
      `[id^="ColorName-${currentSectionId}"]`
    );
    if (colorNameSources?.length === colorNameDestinations?.length) {
      colorNameDestinations.forEach((colorNameDestination, index) => {
        colorNameDestination.classList.remove("visibility-hidden");
        colorNameDestination.innerHTML = colorNameSources[index].innerHTML;
      });
    }

    // variant image swatches
    if (this.isHighVariantNeedUpdate !== true) {
      const variantSwatchesSource = html.querySelector(
        `#${CSS.escape(
          `variant-picker-${sourceSectionId}`
        )} [data-is-variant-image-swatch="true"]`
      );
      const variantSwatchesDestination = scope.querySelector(
        `#${CSS.escape(
          `variant-picker-${currentSectionId}`
        )} [data-is-variant-image-swatch="true"]`
      );
      if (variantSwatchesSource && variantSwatchesDestination) {
        if (currentSectionId !== sourceSectionId) {
          variantSwatchesDestination.innerHTML =
            variantSwatchesSource.innerHTML.replaceAll(
              sourceSectionId,
              currentSectionId
            );
        } else {
          variantSwatchesDestination.innerHTML =
            variantSwatchesSource.innerHTML;
        }
      }
    }

    // product media — target outer gallery wrapper (several nodes share data-section in markup)
    if (
      this.dataset?.variantMediaDisplay !== "show_all" ||
      this.isCombinedListingsNeedUpdate
    ) {
      let mediaSource =
        html.querySelector(
          `.product__media-wrapper[data-section="product-media-${sourceSectionId}"]`
        ) ||
        html.querySelector(`[data-section="product-media-${sourceSectionId}"]`);
      if (!mediaSource && this.isCombinedListingsNeedUpdate) {
        const srcMain = html.querySelector(
          `[data-product-main="${sourceSectionId}"]`
        );
        if (srcMain) {
          mediaSource = srcMain.querySelector(".product__media-wrapper");
        }
      }
      const mediaDestination =
        scope.querySelector(
          `.product__media-wrapper[data-section="product-media-${currentSectionId}"]`
        ) ||
        scope.querySelector(
          `[data-section="product-media-${currentSectionId}"]`
        );

      if (mediaSource && mediaDestination) {
        const mountRoot =
          this.closest("quick-add-modal") ||
          this.closest("#quick-view") ||
          document.getElementById(`shopify-section-${currentSectionId}`) ||
          scope;
        const mediaWrapper =
          mediaDestination.closest(".product__media-wrapper") ||
          mediaDestination;

        mediaWrapper.classList.add("product__media-wrapper--updating");
        this.destroyMediaSwipersWithin(mountRoot);

        mediaDestination.innerHTML =
          currentSectionId !== sourceSectionId
            ? mediaSource.innerHTML.replaceAll(
                sourceSectionId,
                currentSectionId
              )
            : mediaSource.innerHTML;

        const parentQuickViewModal = this.closest("quick-add-modal");
        const parentFeaturedProduct = this.closest(".featured-product-section");

        if (
          parentQuickViewModal &&
          typeof parentQuickViewModal.removeDOMElements === "function"
        ) {
          parentQuickViewModal.removeDOMElements(mediaDestination);
        }

        if (
          parentFeaturedProduct &&
          typeof window.initFeaturedProduct === "function"
        ) {
          const section = document.getElementById(
            `shopify-section-${currentSectionId}`
          );
          if (section) window.initFeaturedProduct(section);
        } else {
          this.reinitProductMediaAfterDomSwap(
            mediaDestination.closest(".product__media-wrapper") || mountRoot
          );
        }

        requestAnimationFrame(() => {
          mediaWrapper.classList.remove("product__media-wrapper--updating");
        });
      }
    }

    // combined listings — title row (badges), subtitle, buy box, long description tabs
    if (this.isCombinedListingsNeedUpdate) {
      this.syncProductMediaModalFromFetchedHtml(
        html,
        currentSectionId,
        sourceSectionId
      );
      const titleSrcHeading = html.getElementById(
        `ProductTitle-${sourceSectionId}`
      );
      const wrapSrc = titleSrcHeading?.closest(".product__title__wrapper");
      const titleDstHeading = byId(scope, `ProductTitle-${currentSectionId}`);
      const wrapDest = titleDstHeading?.closest(".product__title__wrapper");
      if (wrapSrc && wrapDest) {
        const nextWrap = cloneSubtreeForCurrentSection(wrapSrc);
        if (nextWrap) wrapDest.replaceWith(nextWrap);
      }
      const textDest = scope.querySelector(
        ".product__info-container .product__text"
      );
      const textSrc = html.querySelector(
        ".product__info-container .product__text"
      );
      if (textDest && textSrc) {
        textDest.innerHTML = textSrc.innerHTML;
      }

      const buySrc = html.querySelector(".buy-box__wrapper");
      const buyDest = scope.querySelector(".buy-box__wrapper");
      if (buySrc && buyDest) {
        const nextBuy = cloneSubtreeForCurrentSection(buySrc);
        if (nextBuy) buyDest.replaceWith(nextBuy);
      }

      const tabsSrc = html.querySelector(".main-product__tabs-blocks--wrapper");
      const tabsDest = scope.querySelector(
        ".main-product__tabs-blocks--wrapper"
      );
      if (tabsSrc && tabsDest) {
        tabsDest.innerHTML =
          currentSectionId !== sourceSectionId
            ? tabsSrc.innerHTML.replaceAll(sourceSectionId, currentSectionId)
            : tabsSrc.innerHTML;
      }

      const sectionEl = document.getElementById(
        `shopify-section-${currentSectionId}`
      );
      const mainProductEl = document.getElementById(
        `MainProduct-${currentSectionId}`
      );
      const dispatchRoot = mainProductEl || sectionEl;
      if (dispatchRoot) {
        dispatchRoot.dispatchEvent(
          new CustomEvent("shopify:section:load", { bubbles: true })
        );
      }
    }

    const resolveProductInfo = (root, sectionIdForInfoId) => {
      return (
        byId(root, `ProductInfo-${sectionIdForInfoId}`) ||
        root.querySelector('[id^="ProductInfo-"]')
      );
    };

    const infoSrc = resolveProductInfo(html, sourceSectionId);
    const infoDest = resolveProductInfo(scope, currentSectionId);
    if (infoSrc && infoDest) {
      const srcTags = infoSrc.querySelector("ul.product-tags");
      const dstTags = infoDest.querySelector("ul.product-tags");
      if (srcTags && dstTags) {
        const nextTags = cloneSubtreeForCurrentSection(srcTags);
        if (nextTags) dstTags.replaceWith(nextTags);
      } else if (dstTags && !srcTags) {
        dstTags.remove();
      } else if (srcTags && !dstTags) {
        const nextTags = cloneSubtreeForCurrentSection(srcTags);
        const titleWrap = infoDest.querySelector(".product__title__wrapper");
        if (nextTags) {
          if (titleWrap) {
            titleWrap.insertAdjacentElement("afterend", nextTags);
          } else {
            infoDest.prepend(nextTags);
          }
        }
      }

      const getDescriptionSyncRoot = (root) => {
        const fixed =
          root.querySelector(
            ".product__description-fixed.main-product-description"
          ) || root.querySelector(".product__description-fixed");
        if (fixed) return fixed;
        const md = root.querySelector(
          ".product-about__accordion-content.main-product-description"
        );
        if (md) return md.closest(".product-about");
        const md2 = root.querySelector(".main-product-description");
        if (md2) return md2.closest(".product-about");
        return null;
      };
      const srcDescRoot = getDescriptionSyncRoot(infoSrc);
      const dstDescRoot = getDescriptionSyncRoot(infoDest);
      if (srcDescRoot && dstDescRoot) {
        const nextDesc = cloneSubtreeForCurrentSection(srcDescRoot);
        if (nextDesc) dstDescRoot.replaceWith(nextDesc);
      } else if (dstDescRoot && !srcDescRoot) {
        dstDescRoot.remove();
      } else if (srcDescRoot && !dstDescRoot) {
        const nextDesc = cloneSubtreeForCurrentSection(srcDescRoot);
        const anchor =
          infoDest.querySelector("ul.product-tags") ||
          infoDest.querySelector(".product__title__wrapper");
        if (nextDesc) {
          if (anchor) {
            anchor.insertAdjacentElement("afterend", nextDesc);
          } else {
            infoDest.appendChild(nextDesc);
          }
        }
      }
    }

    syncStickyQuantity(currentSectionId);
  }

  // methods for high variant products
  getSelectedValuesIds() {
    return this.getSortedOptionControls().map((control) => {
      const checkedInput = control.querySelector('input[type="radio"]:checked');
      return checkedInput?.dataset?.optionValueId
        ? checkedInput.dataset.optionValueId
        : null;
    });
  }

  createRequestUrl({
    currentVariantId = "",
    selectedValuesIds = [],
    combinedProductURL = "",
  }) {
    const productUrl = combinedProductURL || `${this.dataset.url}`;
    const sectionId = this.dataset.originalSection
      ? this.dataset.originalSection
      : this.dataset.section;

    if (currentVariantId) {
      return `${productUrl}?variant=${currentVariantId}&section_id=${sectionId}`;
    }

    // -----
    // for high-variant products
    // and if variant not found in liquid <script data-all-variants-no-high>
    if (selectedValuesIds.length) {
      const params = [];
      params.push(`section_id=${sectionId}`);
      params.push(`option_values=${selectedValuesIds.join(",")}`);
      return `${productUrl}?${params.join("&")}`;
    }
    // -----
  }

  /**
   * innerHTML replace does not update the custom element host. After combined listing
   * switches to another product, keep data-url / data-is-high-variant-product in sync
   * with the fetched section HTML.
   */
  syncVariantPickerHostAttributesFromFetchedSource(sourceHost, destHost) {
    if (!sourceHost || !destHost) return;
    const ATTRS = [
      "data-url",
      "data-is-high-variant-product",
      "data-update-url",
      "data-variant-media-display",
    ];
    ATTRS.forEach((attr) => {
      if (sourceHost.hasAttribute(attr)) {
        destHost.setAttribute(attr, sourceHost.getAttribute(attr));
      }
    });
    if (destHost.closest?.("quick-add-modal")) {
      destHost.setAttribute("data-update-url", "false");
    }
  }

  setCurrentVariantAfterFetch(html, requestUrl = "") {
    // attr data-original-section use for Quick view modal
    const sourceSectionId = this.dataset.originalSection
      ? this.dataset.originalSection
      : this.dataset.section;

    const variantPickerSource = html.getElementById(
      `variant-picker-${sourceSectionId}`
    );
    const variantPickerDestination = document.getElementById(
      `variant-picker-${this.dataset.section}`
    );
    if (!variantPickerSource || !variantPickerDestination) return;

    const newVariantDataEl = variantPickerSource.querySelector(
      "[data-selected-variant]"
    );
    if (!newVariantDataEl) return;

    const raw = newVariantDataEl.textContent?.trim();
    let selectedVariant = null;
    try {
      selectedVariant = raw ? JSON.parse(raw) : null;
    } catch {
      selectedVariant = null;
    }

    const variantId = this.extractVariantIdFromRequestUrl(requestUrl);
    const variantFromList = this.resolveVariantFromFetchedHtml(
      html,
      sourceSectionId,
      variantId
    );
    if (variantFromList) {
      selectedVariant = variantFromList;
    } else if (
      this.currentVariant?.id &&
      variantId &&
      String(this.currentVariant.id) === String(variantId)
    ) {
      selectedVariant = this.currentVariant;
    }

    this.currentVariant = selectedVariant;
    this.updateSelectedVariantJson();

    const oldEl = variantPickerDestination.querySelector(
      "[data-selected-variant]"
    );
    if (oldEl && this.currentVariant) {
      oldEl.textContent = JSON.stringify(this.currentVariant);
    }
  }

  updatePickerInnerHtml(html) {
    // attr data-original-section use for Quick view modal
    const currentSectionId = this.dataset.section;
    const sourceSectionId = this.dataset.originalSection
      ? this.dataset.originalSection
      : this.dataset.section;

    const variantPickerSource = html.getElementById(
      `variant-picker-${sourceSectionId}`
    );
    const variantPickerDestination = document.getElementById(
      `variant-picker-${currentSectionId}`
    );

    if (variantPickerSource && variantPickerDestination) {
      variantPickerDestination.innerHTML =
        currentSectionId !== sourceSectionId
          ? variantPickerSource.innerHTML.replaceAll(
              sourceSectionId,
              currentSectionId
            )
          : variantPickerSource.innerHTML;

      this.syncVariantPickerHostAttributesFromFetchedSource(
        variantPickerSource,
        variantPickerDestination
      );

      this.variantData = null;
    }
  }
}

if (!customElements.get("variant-selects")) {
  customElements.define("variant-selects", VariantSelects);
}

class VariantRadios extends VariantSelects {
  constructor() {
    super();
  }

  setInputAvailability(listOfOptions, listOfAvailableOptions) {
    listOfOptions.forEach((input) => {
      if (input.dataset?.productUrl) {
        return;
      }

      const rawVal = input.getAttribute("value");
      const isAvail = listOfAvailableOptions.some(
        (v) => String(v) === String(rawVal)
      );
      if (isAvail) {
        input.classList.remove("disabled");
        input.disabled = false;
      } else {
        input.classList.add("disabled");
        //input.disabled = true;
      }
    });
  }

  updateOptions() {
    const fieldsets = Array.from(this.querySelectorAll("fieldset"));
    this.options = fieldsets.map((fieldset) => {
      return Array.from(fieldset.querySelectorAll("input")).find(
        (radio) => radio.checked
      ).value;
    });
  }
}

if (!customElements.get("variant-radios")) {
  customElements.define("variant-radios", VariantRadios);
}

class VariantDropdownSelect extends HTMLElement {
  constructor() {
    super();

    this.currentEl = this.querySelector(".dropdown-select__current");
    this.optionsWrapperEl = this.querySelector(".dropdown-select__options");
    this.hiddenInput = this.querySelector("input[type='hidden']");
    if (!this.currentEl || !this.optionsWrapperEl || !this.hiddenInput) return;

    this.optionsEls = Array.from(this.optionsWrapperEl.querySelectorAll("li"));
    this.inputs = Array.from(this.optionsWrapperEl.querySelectorAll("input"));
    this.hasColorSwatch = this.currentEl?.classList?.contains(
      "dropdown-select__current--with-color"
    );
    this.isActive = false;

    this.optionPosition = this.getOptionPosition();
    this._onCurrentClick = this.onCurrentClick.bind(this);
    this._onKeyUp = this.onKeyUp.bind(this);
    this._onOptionsClick = this.onClickOption.bind(this);
    this._onKeyUpOptions = this.onKeyUpOptions.bind(this);
    this._onOutsideClick = this.onOutsideClick.bind(this);
  }

  connectedCallback() {
    this.currentEl.addEventListener("click", this._onCurrentClick);
    this.currentEl.addEventListener("keyup", this._onKeyUp);
    this.currentEl.addEventListener("dblclick", this.preventDoubleClick);
    this.optionsWrapperEl.addEventListener("click", this._onOptionsClick);
    this.optionsWrapperEl.addEventListener("keyup", this._onKeyUpOptions);
    document.addEventListener("click", this._onOutsideClick);
  }

  disconnectedCallback() {
    this.currentEl.removeEventListener("click", this._onCurrentClick);
    this.currentEl.removeEventListener("keyup", this._onKeyUp);
    this.currentEl.removeEventListener("dblclick", this.preventDoubleClick);
    this.optionsWrapperEl.removeEventListener("click", this._onOptionsClick);
    this.optionsWrapperEl.removeEventListener("keyup", this._onKeyUpOptions);
    document.removeEventListener("click", this._onOutsideClick);
  }

  onCurrentClick(event) {
    event.stopPropagation();
    if (this.isActive) {
      this.onClose();
    } else {
      this.onOpen();
    }
  }

  onKeyUp(event) {
    if (event.code?.toUpperCase() === "ENTER") {
      event.preventDefault();
      event.stopPropagation();
      if (this.isActive) {
        this.onClose();
      } else {
        this.onOpen();
      }
    }
  }

  preventDoubleClick(event) {
    event.preventDefault();
    event.stopPropagation();
  }

  onClickOption(event) {
    const optionEl = event.target.closest("li");
    const currentValueEl = this.currentEl.querySelector(
      "[data-dropdown-current-value]"
    );

    if (
      !optionEl ||
      !currentValueEl ||
      !this.optionsWrapperEl.contains(optionEl)
    ) {
      event.preventDefault();
      return;
    }

    const newValue = optionEl.dataset.value;

    currentValueEl.textContent = newValue;
    this.hiddenInput.value = newValue;

    if (this.hasColorSwatch) {
      const newColor = optionEl.dataset.color;
      this.hiddenInput.dataset.colorSwatch = newColor;
      this.currentEl.style.setProperty("--swatch-color", newColor);
    }

    this.hiddenInput.dispatchEvent(new Event("change", { bubbles: true }));

    this.inputs.forEach((el) => el.removeAttribute("checked"));
    if (event.target.tagName === "INPUT") {
      event.target.setAttribute("checked", "true");
    }

    this.onClose();
    this.currentEl.focus();
  }

  onKeyUpOptions(event) {
    if (event.code === "Escape" && this.isActive) {
      event.preventDefault();
      this.onClose();
      this.currentEl.focus();
    }
  }

  onOpen() {
    document.querySelectorAll("variant-dropdown-select").forEach((dropdown) => {
      if (dropdown !== this && typeof dropdown.onClose === "function") {
        dropdown.onClose();
      }
    });

    this.optionsWrapperEl.classList.add("active");
    this.currentEl.setAttribute("aria-expanded", "true");
    this.isActive = true;

    const currentIndex = this.inputs.findIndex((inp) =>
      inp.hasAttribute("checked")
    );
    this.inputs[currentIndex >= 0 ? currentIndex : 0]?.focus();
  }

  onClose() {
    this.optionsWrapperEl.classList.remove("active");
    this.currentEl.setAttribute("aria-expanded", "false");
    this.isActive = false;
  }

  onOutsideClick(event) {
    if (!this.isActive || this.contains(event.target)) return;
    this.onClose();
  }

  getOptionPosition() {
    const parentFieldset = this.closest(".product-form__controls");
    return Number(parentFieldset?.dataset?.optionPosition || -1);
  }

  updateCurrentOption(selectedValues = []) {
    // Method is used to synchronize after changes options in sticky bar.
    if (this.optionPosition === -1) return;
    const currentValueEl = this.currentEl.querySelector(
      "[data-dropdown-current-value]"
    );

    if (!currentValueEl || !selectedValues || !selectedValues.length) return;

    const newValue = selectedValues[this.optionPosition - 1];

    currentValueEl.textContent = newValue;
    this.hiddenInput.value = newValue;

    this.optionsEls.forEach((optionEl) => {
      const liDataValue = optionEl.dataset.value;
      if (newValue === liDataValue && this.hasColorSwatch) {
        const newColor = optionEl.dataset.color;
        this.hiddenInput.dataset.colorSwatch = newColor;
        this.currentEl.style.setProperty("--swatch-color", newColor);
      }
    });
  }
}

if (!customElements.get("variant-dropdown-select")) {
  customElements.define("variant-dropdown-select", VariantDropdownSelect);
}

class PasswordViewer {
  constructor() {
    const passwordField = document.querySelectorAll(".field--pass");

    passwordField.forEach((el) => {
      const input = el.querySelector("input");
      const btnWrapper = el.querySelector(".button-pass-visibility");
      const btnOpen = el.querySelector(".icon-eye-close");
      const btnClose = el.querySelector(".icon-eye");

      input.addEventListener("input", () => {
        input.value !== ""
          ? (btnWrapper.style.display = "block")
          : (btnWrapper.style.display = "none");
      });

      btnOpen.addEventListener("click", () => {
        input.type = "text";
        btnOpen.style.display = "none";
        btnClose.style.display = "block";
      });

      btnClose.addEventListener("click", () => {
        input.type = "password";
        btnOpen.style.display = "block";
        btnClose.style.display = "none";
      });
    });
  }
}

class ProductRecommendations extends HTMLElement {
  constructor() {
    super();

    const handleIntersection = (entries, observer) => {
      if (!entries[0].isIntersecting) return;
      observer.unobserve(this);

      if (this.querySelector(".product-recommendations__loading")) {
        this.querySelector(".product-recommendations__loading").classList.add(
          "loading"
        );
        this.querySelector(".product-recommendations__loading").style.display =
          "flex";
      }

      fetch(this.dataset.url)
        .then((response) => response.text())
        .then((text) => {
          const html = document.createElement("div");
          html.innerHTML = text;
          const recommendations = html.querySelector("product-recommendations");
          if (recommendations && recommendations.innerHTML.trim().length) {
            this.innerHTML = recommendations.innerHTML;
          }

          if (this.querySelector(".product-recommendations__empty")) {
            this.querySelector(
              ".product-recommendations__empty"
            ).style.display = "flex";
          }

          /* Color swatches */
          const generateSrcset = (image, widths = []) => {
            const imageUrl = new URL(image["src"]);
            return widths
              .filter((width) => width <= image["width"])
              .map((width) => {
                imageUrl.searchParams.set("width", width.toString());
                return `${imageUrl.href} ${width}w`;
              })
              .join(", ");
          };

          const createImageElement = (image, classes, sizes, productTitle) => {
            const previewImage = image["preview_image"];
            const newImage = new Image(
              previewImage["width"],
              previewImage["height"]
            );
            newImage.className = classes;
            newImage.alt = image["alt"] || productTitle;
            newImage.sizes = sizes;
            newImage.src = previewImage["src"];
            newImage.srcset = generateSrcset(
              previewImage,
              [165, 360, 533, 720, 940, 1066]
            );
            newImage.loading = "lazy";
            return newImage;
          };

          const checkSwatches = () => {
            document
              .querySelectorAll(".js-color-swatches-wrapper")
              .forEach((wrapper) => {
                wrapper
                  .querySelectorAll(".js-color-swatches input")
                  .forEach((input) => {
                    input.addEventListener("click", (event) => {
                      const primaryImage =
                        wrapper.querySelector(".media--first");
                      const secondaryImage =
                        wrapper.querySelector(".media--second");
                      const handleProduct = wrapper.dataset.product;

                      if (event.currentTarget.checked && primaryImage) {
                        wrapper
                          .querySelector(".js-color-swatches-link")
                          .setAttribute(
                            "href",
                            event.currentTarget.dataset.variantLink
                          );
                        if (
                          wrapper.querySelector(
                            '.card__add-to-cart button[name="add"]'
                          )
                        ) {
                          wrapper
                            .querySelector(
                              '.card__add-to-cart button[name="add"]'
                            )
                            .setAttribute("aria-disabled", false);
                          if (
                            wrapper.querySelector(
                              '.card__add-to-cart button[name="add"] > span'
                            )
                          ) {
                            wrapper
                              .querySelector(
                                '.card__add-to-cart button[name="add"] > span'
                              )
                              .classList.remove("hidden");
                            wrapper
                              .querySelector(
                                '.card__add-to-cart button[name="add"] .sold-out-message'
                              )
                              .classList.add("hidden");
                          }
                          wrapper.querySelector(
                            '.card__add-to-cart input[name="id"]'
                          ).value = event.currentTarget.dataset.variantId;
                        }
                        const currentColor = event.currentTarget.value;

                        jQuery.getJSON(
                          window.Shopify.routes.root +
                            `products/${handleProduct}.js`,
                          function (product) {
                            const variant = product.variants.filter(
                              (item) =>
                                item.featured_media != null &&
                                item.options.includes(currentColor)
                            )[0];

                            if (variant) {
                              const newPrimaryImage = createImageElement(
                                variant["featured_media"],
                                primaryImage.className,
                                primaryImage.sizes,
                                product.title
                              );

                              if (newPrimaryImage.src !== primaryImage.src) {
                                let flag = false;
                                if (secondaryImage) {
                                  const secondaryImagePathname = new URL(
                                    secondaryImage.src
                                  ).pathname;
                                  const newPrimaryImagePathname = new URL(
                                    newPrimaryImage.src
                                  ).pathname;

                                  if (
                                    secondaryImagePathname ==
                                    newPrimaryImagePathname
                                  ) {
                                    primaryImage.remove();
                                    secondaryImage.classList.remove(
                                      "media--second"
                                    );
                                    secondaryImage.classList.add(
                                      "media--first"
                                    );
                                    flag = true;
                                  }
                                }
                                if (flag == false) {
                                  primaryImage.animate(
                                    { opacity: [1, 0] },
                                    {
                                      duration: 200,
                                      easing: "ease-in",
                                      fill: "forwards",
                                    }
                                  ).finished;
                                  setTimeout(function () {
                                    primaryImage.replaceWith(newPrimaryImage);
                                    newPrimaryImage.animate(
                                      { opacity: [0, 1] },
                                      { duration: 200, easing: "ease-in" }
                                    );
                                    if (secondaryImage) {
                                      secondaryImage.remove();
                                    }
                                  }, 200);
                                }
                              }
                            }
                          }
                        );
                      }
                    });
                  });
              });
          };

          checkSwatches();
          renderProductCardSpecs();

          const addClasses = (slider) => {
            const sliderWrapper = slider.querySelector(
              ".product-recommendations__wrapper"
            );
            const slides = slider.querySelectorAll(
              ".product-recommendations__item"
            );

            slider.classList.add("swiper");
            if (sliderWrapper) sliderWrapper.classList.add("swiper-wrapper");

            if (slides.length > 1) {
              slides.forEach((slide) => {
                slide.classList.add("swiper-slide");
              });
            }
          };

          const removeClasses = (slider) => {
            const sliderWrapper = slider.querySelector(
              ".product-recommendations__wrapper"
            );
            const slides = slider.querySelectorAll(
              ".product-recommendations__item"
            );

            slider.classList.remove("swiper");
            if (sliderWrapper) sliderWrapper.classList.remove("swiper-wrapper");

            if (slides.length > 0) {
              slides.forEach((slide) => {
                slide.removeAttribute("style");
                slide.classList.remove("swiper-slide");
              });
            }
          };

          const initSlider = () => {
            const slider = this.querySelector(".swiper--recomend-products");

            if (slider) {
              addClasses(slider);
              const numberColumnsMobile = Number(
                slider.dataset.columnsMobile || 1
              );
              const numberColumnsTablet = numberColumnsMobile + 1;

              new Swiper(slider, {
                loop: false,
                speed: 800,
                breakpoints: {
                  320: {
                    slidesPerView: Number(numberColumnsMobile),
                    spaceBetween: 8,
                  },
                  750: {
                    slidesPerView: Number(numberColumnsTablet),
                    spaceBetween: 16,
                  },
                },
                pagination: {
                  el: slider.querySelector(
                    ".product-recommendations__pagination"
                  ),
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

          const destroySlider = () => {
            const slider = this.querySelector(".swiper--recomend-products");

            if (slider) {
              removeClasses(slider);
            }
          };

          const initSection = () => {
            const resizeObserver = new ResizeObserver((entries) => {
              const [entry] = entries;

              if (entry.contentRect.width < 990) {
                initSlider();
              } else {
                destroySlider();
              }
            });

            resizeObserver.observe(this);
          };

          initSection();

          // Initialize buttons animations (global function initButtonsAnimation in global.js)
          initButtonsAnimation(this);
        })
        .catch((e) => {
          console.error(e);
        })
        .finally(() => {
          if (this.querySelector(".product-recommendations__loading")) {
            this.querySelector(
              ".product-recommendations__loading"
            ).classList.remove("loading");
            this.querySelector(".product-recommendations__loading").remove();
          }
        });
    };

    new IntersectionObserver(handleIntersection.bind(this), {
      rootMargin: "0px 0px 200px 0px",
    }).observe(this);
  }
}

customElements.define("product-recommendations", ProductRecommendations);

class LocalizationForm extends HTMLElement {
  constructor() {
    super();
    this.elements = {
      input: this.querySelector(
        'input[name="locale_code"], input[name="country_code"]'
      ),
      button: this.querySelector(".disclosure__button"),
      panel: this.querySelector(".disclosure__list"),
      overlay: this.querySelector(".disclosure__overlay"),
    };
    this.handleDocumentClick = this.handleDocumentClick.bind(this);
    this.elements.button?.addEventListener(
      "click",
      this.togglePanel.bind(this)
    );
    this.elements.overlay?.addEventListener(
      "click",
      this.togglePanel.bind(this)
    );
    this.addEventListener("keydown", this.onEscapePress.bind(this));

    this.querySelectorAll("a").forEach((item) =>
      item.addEventListener("click", this.onItemClick.bind(this))
    );
  }

  connectedCallback() {
    document.addEventListener("click", this.handleDocumentClick);
  }

  disconnectedCallback() {
    document.removeEventListener("click", this.handleDocumentClick);
  }

  handleDocumentClick(event) {
    if (!this.contains(event.target)) {
      this.hidePanel();
    }
  }

  hidePanel() {
    this.elements.button.setAttribute("aria-expanded", "false");
    this.elements.panel.setAttribute("hidden", true);
  }

  showPanel() {
    this.elements.button.setAttribute("aria-expanded", "true");
    this.elements.panel.removeAttribute("hidden");
  }

  togglePanel() {
    if (this.elements.button.getAttribute("aria-expanded") === "true") {
      this.hidePanel();
    } else {
      this.showPanel();
    }
  }

  onEscapePress(event) {
    if (event.key === "Escape") {
      this.hidePanel();
    }
  }

  onItemClick(event) {
    event.preventDefault();
    this.elements.input.value = event.currentTarget.dataset.value;
    this.querySelector("form")?.submit();
    this.hidePanel();
  }
}

customElements.define("localization-form", LocalizationForm);

(function () {
  const initHeaderOverlay = () => {
    const main = document.getElementById("MainContent");
    const sections = main.querySelectorAll(".shopify-section");

    if (sections.length > 0) {
      const sectionFirstChild = sections[0].querySelector(
        "[data-header-overlay]"
      );
      const headerGroupSections = document.querySelectorAll(
        ".shopify-section-group-header-group"
      );
      const header = document.querySelector(".shopify-section-header");
      const breadcrumbs = document.querySelector("body > .breadcrumbs-wrapper");

      if (sectionFirstChild) {
        if (headerGroupSections[headerGroupSections.length - 1] === header) {
          sections[0].classList.add("section--has-overlay");
          header.classList.add("color-background-overlay");
          if (breadcrumbs) breadcrumbs.classList.add("color-background-3");
        } else {
          sections[0].classList.remove("section--has-overlay");
          header.classList.remove("color-background-overlay");
          if (breadcrumbs) breadcrumbs.classList.remove("color-background-3");
        }
      } else {
        sections[0].classList.remove("section--has-overlay");
        header?.classList?.remove("color-background-overlay");
        if (breadcrumbs) breadcrumbs.classList.remove("color-background-3");
      }
    }
  };

  initHeaderOverlay();

  document.addEventListener("shopify:section:load", initHeaderOverlay);
  document.addEventListener("shopify:section:unload", initHeaderOverlay);
  document.addEventListener("shopify:section:reorder", initHeaderOverlay);
})();

function formatMoney(cents, format = "") {
  if (typeof cents === "string") {
    cents = cents.replace(".", "");
  }

  cents = parseInt(cents, 10);

  let value = "";
  const placeholderRegex = /\{\{\s*(\w+)\s*\}\}/;
  const formatString = format || theme.moneyFormat;

  function formatWithDelimiters(
    number,
    precision = 2,
    thousands = ",",
    decimal = "."
  ) {
    if (isNaN(number) || number == null) {
      return "0";
    }

    number = (number / 100.0).toFixed(precision);

    const parts = number.split(".");
    const dollarsAmount = parts[0].replace(
      /(\d)(?=(\d{3})+(?!\d))/g,
      `$1${thousands}`
    );
    const centsAmount = precision > 0 ? decimal + parts[1] : "";

    return dollarsAmount + centsAmount;
  }

  const match = formatString.match(placeholderRegex);
  const formatType = match ? match[1] : "amount";

  switch (formatType) {
    case "amount":
      value = formatWithDelimiters(cents, 2, ",", ".");
      break;
    case "amount_no_decimals":
      value = formatWithDelimiters(cents, 0, ",", ".");
      break;
    case "amount_with_comma_separator":
      value = formatWithDelimiters(cents, 2, ".", ",");
      break;
    case "amount_no_decimals_with_comma_separator":
      value = formatWithDelimiters(cents, 0, ".", ",");
      break;
    case "amount_with_apostrophe_separator":
      value = formatWithDelimiters(cents, 2, "'", ".");
      break;
    case "amount_no_decimals_with_space_separator":
      value = formatWithDelimiters(cents, 0, " ", ".");
      break;
    case "amount_with_space_separator":
      value = formatWithDelimiters(cents, 2, " ", ",");
      break;
    case "amount_with_period_and_space_separator":
      value = formatWithDelimiters(cents, 2, " ", ".");
      break;
    default:
      value = formatWithDelimiters(cents, 2, ",", ".");
  }

  return formatString.replace(placeholderRegex, value);
}

// BUTTON ANIMATION CODE START
// -----------------------------------------------------------------------------

function initButtonsAnimation(parent) {
  if (!parent) return;
  const initButton = (button) => {
    const circle = button.querySelector(".button_circle");

    const fillEffect = (e) => {
      const rect = button.getBoundingClientRect(),
        i = e.clientY - rect.top,
        s = e.clientX - rect.left,
        r = i / rect.height,
        n = s / rect.width;

      circle.style.transformOrigin = 100 * n + "% " + 100 * r + "%";
    };

    if (circle) {
      button.addEventListener("mouseenter", fillEffect);
      button.addEventListener("mouseleave", fillEffect);
    }
  };

  const buttons = parent.querySelectorAll(
    ".button--primary, .button--secondary, .button--transparent, .customer .button--primary, .customer .button--secondary, .customer .button--transparent, .products-tabs__tab"
  );

  if (buttons.length) {
    buttons.forEach((button) => {
      initButton(button);
    });
  }
}

initButtonsAnimation(document);

// BUTTON ANIMATION CODE END
// -----------------------------------------------------------------------------
