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
  const newImage = new Image(previewImage["width"], previewImage["height"]);
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
  document.querySelectorAll(".js-color-swatches-wrapper").forEach((wrapper) => {
    wrapper.querySelectorAll(".js-color-swatches input").forEach((input) => {
      input.addEventListener("click", (event) => {
        const primaryImage = wrapper.querySelector(".media--first");
        const secondaryImage = wrapper.querySelector(".media--second");
        const handleProduct = wrapper.dataset.product;

        if (event.currentTarget.checked && primaryImage) {
          const variantLink = event.currentTarget.dataset.variantLink;
          const swatchesLink = wrapper.querySelector(".js-color-swatches-link");
          if (variantLink && variantLink !== "undefined") {
            swatchesLink.setAttribute("href", variantLink);
          }

          if (wrapper.querySelector('.card__add-to-cart button[name="add"]')) {
            wrapper
              .querySelector('.card__add-to-cart button[name="add"]')
              .setAttribute("aria-disabled", false);
            if (
              wrapper.querySelector(
                '.card__add-to-cart button[name="add"] > span'
              )
            ) {
              wrapper
                .querySelector('.card__add-to-cart button[name="add"] > span')
                .classList.remove("hidden");
              wrapper
                .querySelector(
                  '.card__add-to-cart button[name="add"] .sold-out-message'
                )
                .classList.add("hidden");
            }
            wrapper.querySelector('.card__add-to-cart input[name="id"]').value =
              event.currentTarget.dataset.variantId;
          }

          const currentColor = event.currentTarget.value;

          jQuery.getJSON(
            window.Shopify.routes.root + `products/${handleProduct}.js`,
            function (product) {
              const variant = product.variants.filter(
                (item) =>
                  item.featured_media != null &&
                  item.options.includes(currentColor)
              )[0];

              if (variant) {
                const resolvedLink = `${window.Shopify.routes.root}products/${product.handle}?variant=${variant.id}`;
                swatchesLink.setAttribute("href", resolvedLink);

                const newPrimaryImage = createImageElement(
                  variant["featured_media"],
                  primaryImage.className,
                  primaryImage.sizes,
                  product.title
                );

                if (newPrimaryImage.src !== primaryImage.src) {
                  let flag = false;
                  if (secondaryImage) {
                    const secondaryImagePathname = new URL(secondaryImage.src)
                      .pathname;
                    const newPrimaryImagePathname = new URL(newPrimaryImage.src)
                      .pathname;

                    if (secondaryImagePathname == newPrimaryImagePathname) {
                      primaryImage.remove();
                      secondaryImage.classList.remove("media--second");
                      secondaryImage.classList.add("media--first");
                      flag = true;
                    }
                  }
                  if (flag == false) {
                    primaryImage.animate(
                      { opacity: [1, 0] },
                      { duration: 200, easing: "ease-in", fill: "forwards" }
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

function colorSwatches() {
  checkSwatches();

  document.addEventListener("shopify:section:load", function () {
    checkSwatches();
  });
}

function productCardSpecs() {
  const toggles = document.querySelectorAll(".js-toggle-specs");

  toggles.forEach((wrapper) => {
    const openBtn = wrapper.querySelector(".specs_closed");
    const closeBtn = wrapper.querySelector(".specs_opened");

    openBtn.addEventListener("mouseenter", () => {
      wrapper.classList.add("is-active");
    });

    openBtn?.addEventListener("click", (e) => {
      e.stopPropagation();
      wrapper.classList.add("is-active");
    });

    closeBtn?.addEventListener("click", (e) => {
      e.stopPropagation();
      wrapper.classList.remove("is-active");
    });
  });

  const cardWrappers = document.querySelectorAll(".product_card_specs");

  cardWrappers.forEach((wrapper) => {
    const jsToggleSpecs = wrapper.querySelector(".js-toggle-specs");

    wrapper.addEventListener("mouseleave", (e) => {
      const toElement = e.relatedTarget;

      if (toElement && toElement.closest(".swiper-buttons__wrapper")) {
        return;
      }

      if (jsToggleSpecs) {
        jsToggleSpecs.classList.remove("is-active");
      }
    });
  });
}

function renderProductCardSpecs() {
  productCardSpecs();

  document.addEventListener("shopify:section:load", function () {
    productCardSpecs();
  });
}

(function () {
  colorSwatches();
  renderProductCardSpecs();
})();
