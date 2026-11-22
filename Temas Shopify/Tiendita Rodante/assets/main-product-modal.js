class ProductModal extends ModalDialog {
  constructor() {
    super();
  }

  hide() {
    removeTrapFocus();
    super.hide();
    window.pauseAllMedia();
  }

  show(opener) {
    super.show(opener);

    this.addEventListener(
      "transitionend",
      () => {
        const containerToTrapFocusOn = this;
        const focusElement = this.querySelector(".product-media-modal__dialog");
        trapFocus(containerToTrapFocusOn, focusElement);
      },
      { once: true }
    );

    this.showActiveMedia();
  }

  showActiveMedia() {
    const mediaId = this.openedBy?.getAttribute("data-media-id");
    if (!mediaId || !window.syncProductPopupToMediaId) return;

    const modal = this;
    const syncToMedia = (attempt = 0) => {
      const popupEl = modal.querySelector(".js-popup-slider");
      if (popupEl?.swiper?.slides?.length) {
        window.syncProductPopupToMediaId(mediaId, 0, modal);
        return;
      }
      if (attempt < 20) {
        requestAnimationFrame(() => syncToMedia(attempt + 1));
      }
    };
    syncToMedia();

    const safeId = mediaId != null ? CSS.escape(String(mediaId)) : "";
    const activeMedia = safeId
      ? this.querySelector(`[data-media-id="${safeId}"]`)
      : null;

    if (
      activeMedia &&
      activeMedia.nodeName == "DEFERRED-MEDIA" &&
      activeMedia
        .querySelector("template")
        ?.content?.querySelector(".js-youtube")
    ) {
      activeMedia.loadContent();
    }
  }
}

customElements.define("product-modal", ProductModal);
