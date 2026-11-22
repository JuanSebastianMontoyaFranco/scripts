(() => {
  const productMarkers = () => {
    const $markers = $(".js-product-markers__item");
    const $cards = $(".product-markers .card-wrapper");

    if ($markers.length === 0) return;

    let isDesktop = window.innerWidth >= 750;
    let eventsAttached = false;

    const setupDesktop = () => {
      if (eventsAttached === "desktop") return;

      $markers.off();

      $markers.on("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          $(this).trigger("mouseenter");
        }
      });

      $markers.each(function () {
        const $marker = $(this);
        const $card = $marker.find(".card-wrapper");
        const box = $(".product-markers__img-box")[0];
        if (!box) return;

        $marker.removeClass(
          "product-markers__card-goToLeft product-markers__card-goToTop product-markers__card-goToBottom"
        );
        $card.css({
          display: "block",
          visibility: "hidden",
          opacity: 0,
          transition: "none",
        });

        const cardRect = $card[0].getBoundingClientRect();
        const boxRect = box.getBoundingClientRect();

        if (cardRect.right > boxRect.right)
          $marker.addClass("product-markers__card-goToLeft");
        if (cardRect.bottom > boxRect.bottom)
          $marker.addClass("product-markers__card-goToTop");
        if (cardRect.top < boxRect.top)
          $marker.addClass("product-markers__card-goToBottom");

        $card.css({
          visibility: "hidden",
          opacity: 0,
          pointerEvents: "none",
          transition: "opacity 0.3s ease-in-out, visibility 0.3s ease-in-out",
        });
      });

      setTimeout(() => {
        $markers
          .first()
          .find(".card-wrapper")
          .css({ visibility: "visible", opacity: 1, pointerEvents: "auto" });
      }, 100);

      let hideTimeout;
      $markers.on("mouseenter", function () {
        const $currentMarker = $(this);
        clearTimeout(hideTimeout);
        $cards
          .not($currentMarker.find(".card-wrapper"))
          .css({ visibility: "hidden", opacity: 0, pointerEvents: "none" });
        $currentMarker
          .find(".card-wrapper")
          .css({ visibility: "visible", opacity: 1, pointerEvents: "auto" });
      });

      $markers.on("mouseleave", () => {
        hideTimeout = setTimeout(() => {
          $cards.css({
            visibility: "hidden",
            opacity: 0,
            pointerEvents: "none",
          });
        }, 100);
      });

      eventsAttached = "desktop";
    };

    const setupMobile = () => {
      if (eventsAttached === "mobile") return;

      $markers.off();
      $cards.css({ visibility: "hidden", opacity: 0, pointerEvents: "none" });

      $markers.on("click", function () {
        const index = $(this).data("index");
        $(".product-markers-for-mobile .js-product-markers__item-inner").hide();
        $(
          `.product-markers-for-mobile .js-product-markers__item-inner[data-index='${index}']`
        ).show();
      });

      // Show first product on load
      $(
        ".product-markers-for-mobile .js-product-markers__item-inner[data-index='1']"
      ).show();

      eventsAttached = "mobile";
    };

    const handleResize = () => {
      const newIsDesktop = window.innerWidth >= 750;
      if (newIsDesktop !== isDesktop) {
        isDesktop = newIsDesktop;
        isDesktop ? setupDesktop() : setupMobile();
      }
    };

    isDesktop ? setupDesktop() : setupMobile();
    window.addEventListener("resize", handleResize);
  };

  document.addEventListener("DOMContentLoaded", productMarkers);
  document.addEventListener("shopify:section:load", productMarkers);
})();
