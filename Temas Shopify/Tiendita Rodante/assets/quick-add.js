if (!customElements.get("quick-add-modal")) {
  customElements.define(
    "quick-add-modal",
    class QuickAddModal extends ModalDialog {
      constructor() {
        super();
        this.modalContent = this.querySelector('[id^="QuickAddInfo-"]');

        this.addEventListener("keyup", (event) => {
          if (event.code.toUpperCase() === "ESCAPE") this.hide();
        });

        this.querySelector(".quick-add-modal__overlay")?.addEventListener(
          "click",
          this.hide.bind(this, false)
        );
      }

      hide(preventFocus = false) {
        const cartDrawer = document.querySelector("cart-drawer");
        if (cartDrawer) cartDrawer.setActiveElement(this.openedBy);
        this.modalContent.innerHTML = "";

        $(".js-media-list").each(function () {
          this.swiper?.destroy();
        });
        $(".js-media-sublist").each(function () {
          this.swiper?.destroy();
        });

        subSliderInit(true);
        sliderInit(true);

        if (preventFocus) this.openedBy = null;

        removeTrapFocus();

        super.hide();

        document.body.classList.remove("overflow-hidden-modal");
      }

      show(opener) {
        opener.setAttribute("aria-disabled", true);
        opener.classList.add("loading");

        if (opener.querySelector(".loading-overlay__spinner")) {
          opener
            .querySelector(".loading-overlay__spinner")
            .classList.remove("hidden");
        }

        fetch(opener.getAttribute("data-product-url"))
          .then((response) => response.text())
          .then((responseText) => {
            const responseHTML = new DOMParser().parseFromString(
              responseText,
              "text/html"
            );
            this.productElement = responseHTML.querySelector(
              'section[id^="MainProduct-"]'
            );
            this.preventDuplicatedIDs();
            this.removeDOMElements();
            this.setInnerHTML(
              this.modalContent,
              this.productElement.innerHTML,
              opener
            );

            if (window.Shopify && window.Shopify.PaymentButton) {
              window.Shopify.PaymentButton.init();
            }

            if (window.ProductModel) window.ProductModel.loadShopifyXR();

            this.removeGalleryListSemantic();
            this.updateImageSizes();
            this.preventVariantURLSwitching();
            super.show(opener);

            this.addEventListener(
              "transitionend",
              () => {
                const containerToTrapFocusOn = this;
                const focusElement = this.querySelector(
                  ".quick-add-modal__content"
                );
                trapFocus(containerToTrapFocusOn, focusElement);
              },
              { once: true }
            );

            document.body.classList.add("overflow-hidden-modal");
          })
          .finally(() => {
            opener.removeAttribute("aria-disabled");
            opener.classList.remove("loading");

            if (opener.querySelector(".loading-overlay__spinner")) {
              opener
                .querySelector(".loading-overlay__spinner")
                .classList.add("hidden");
            }

            subSliderInit(true);
            sliderInit(true);
          });
      }

      setInnerHTML(element, html, opener) {
        element.innerHTML = html;

        // Reinjects the script tags to allow execution. By default, scripts are disabled when using element.innerHTML.
        element.querySelectorAll("script").forEach((oldScriptTag) => {
          const newScriptTag = document.createElement("script");
          Array.from(oldScriptTag.attributes).forEach((attribute) => {
            newScriptTag.setAttribute(attribute.name, attribute.value);
          });
          newScriptTag.appendChild(
            document.createTextNode(oldScriptTag.innerHTML)
          );
          oldScriptTag.parentNode.replaceChild(newScriptTag, oldScriptTag);
        });

        const moreBtn = document.createElement("a");
        moreBtn.innerHTML = `<span>${theme.quickviewMore}</span>`;
        moreBtn.setAttribute("href", opener.getAttribute("data-product-url"));
        moreBtn.setAttribute(
          "class",
          "product__full-details animated-link__underline--reverse"
        );
        if (
          element.querySelectorAll(".product__info-column") &&
          element.querySelectorAll(".product__info-column").length > 0
        ) {
          element.querySelectorAll(".product__info-column").forEach((el, i) => {
            if (i === 1) {
              el.appendChild(moreBtn);
            }
          });
        } else {
          element
            .querySelector(".product__info-container")
            .appendChild(moreBtn);
        }
        // Initialize buttons animations (global function initButtonsAnimation in global.js)
        initButtonsAnimation(element);
      }

      removeDOMElements() {
        const modal = this.productElement.querySelector("product-modal");
        if (modal) modal.remove();

        const breadcrumb = this.productElement.querySelectorAll(".breadcrumb");
        breadcrumb.forEach((el) => {
          const container = el.closest(".container");
          if (container) {
            container.remove();
          }
        });

        const popup = this.productElement.querySelectorAll(".product-popup");
        popup.forEach((el) => {
          el.remove();
        });

        const sku = this.productElement.querySelector(".product__sku");
        if (sku) sku.remove();

        const about = this.productElement.querySelectorAll(".about");
        about.forEach((el) => {
          el.remove();
        });

        const customLiquid =
          this.productElement.querySelectorAll(".custom-liquid");
        customLiquid.forEach((el) => {
          el.remove();
        });

        const shareButtons =
          this.productElement.querySelector(".share-buttons");
        if (shareButtons) shareButtons.remove();

        const tags = this.productElement.querySelector(".product-tags");
        if (tags) tags.remove();

        const pickupAvailability = this.productElement.querySelectorAll(
          ".pickup-availability"
        );
        if (pickupAvailability) pickupAvailability.forEach((el) => el.remove());

        const badgeTags = this.productElement.querySelectorAll(".badge");
        if (badgeTags) badgeTags.forEach((el) => el.remove());

        const description = this.productElement.querySelector(
          ".product__description"
        );
        if (description) description.remove();

        const stickyAddToCart = this.productElement.querySelector(
          ".sticky-add-to-cart"
        );
        if (stickyAddToCart) stickyAddToCart.remove();

        const mainProductCountdown = this.productElement.querySelector(
          ".main-product__countdown"
        );
        if (mainProductCountdown) mainProductCountdown.remove();

        const stockBarWrapper =
          this.productElement.querySelector(".stock-bar-wrapper");
        if (stockBarWrapper) stockBarWrapper.remove();

        const productInventory = this.productElement.querySelector(
          ".product__inventory"
        );
        if (productInventory) productInventory.remove();

        const productPickupAvailability = this.productElement.querySelector(
          ".product__pickup-availabilities"
        );
        if (productPickupAvailability) productPickupAvailability.remove();

        const expressShipping =
          this.productElement.querySelector(".express-shipping");
        if (expressShipping) expressShipping.remove();

        const productSpecsShort = this.productElement.querySelector(
          ".product-specs__short"
        );
        if (productSpecsShort) productSpecsShort.remove();

        const productInfonavWrapper = this.productElement.querySelector(
          ".product__info-navigations--wrapper"
        );
        if (productInfonavWrapper) productInfonavWrapper.remove();

        const productTabsBlocksWrapper = this.productElement.querySelector(
          ".main-product__tabs-blocks--wrapper"
        );
        if (productTabsBlocksWrapper) productTabsBlocksWrapper.remove();
      }

      preventDuplicatedIDs() {
        const sectionId = this.productElement.dataset.section;
        this.productElement.innerHTML =
          this.productElement.innerHTML.replaceAll(
            sectionId,
            `quickadd-${sectionId}`
          );
        this.productElement
          .querySelectorAll("variant-selects, variant-radios")
          .forEach((variantSelect) => {
            variantSelect.dataset.originalSection = sectionId;
          });
      }

      preventVariantURLSwitching() {
        if (this.modalContent.querySelector("variant-radios,variant-selects")) {
          this.modalContent
            .querySelector("variant-radios,variant-selects")
            .setAttribute("data-update-url", "false");
        }
      }

      removeGalleryListSemantic() {
        const galleryList = this.modalContent.querySelector(
          '[id^="Slider-Gallery"]'
        );
        if (!galleryList) return;

        galleryList.setAttribute("role", "presentation");
        galleryList
          .querySelectorAll('[id^="Slide-"]')
          .forEach((li) => li.setAttribute("role", "presentation"));
      }

      updateImageSizes() {
        const product = this.modalContent.querySelector(".product");
        const desktopColumns = product.classList.contains("product--columns");
        if (!desktopColumns) return;

        const mediaImages = product.querySelectorAll(".product__media img");
        if (!mediaImages.length) return;

        let mediaImageSizes =
          "(min-width: 1000px) 715px, (min-width: 750px) calc((100vw - 11.5rem) / 2), calc(100vw - 4rem)";

        if (product.classList.contains("product--medium")) {
          mediaImageSizes = mediaImageSizes.replace("715px", "605px");
        } else if (product.classList.contains("product--small")) {
          mediaImageSizes = mediaImageSizes.replace("715px", "495px");
        }

        mediaImages.forEach((img) =>
          img.setAttribute("sizes", mediaImageSizes)
        );
      }
    }
  );
}
