class CartDrawer extends HTMLElement {
  constructor() {
    super();

    this.header = document.querySelector(".header-wrapper");
    if (this.header) this.header.preventHide = false;

    this.trapFocusSetup = false;

    this.addEventListener(
      "keyup",
      (evt) => evt.code === "Escape" && this.close()
    );
    this.querySelector("#CartDrawer-Overlay").addEventListener(
      "click",
      this.close.bind(this)
    );
    this.setHeaderCartIconAccessibility();
  }

  addCloseButtonHandlers() {
    const closeDetails = (detailsElement) => {
      detailsElement?.removeAttribute("open");
    };

    const discountDetails = this.querySelector("#Details-CartDrawer-Discount");
    const discountCloseButton = this.querySelector("#cart_discount-close");

    if (discountDetails && discountCloseButton) {
      discountCloseButton.addEventListener("click", () =>
        closeDetails(discountDetails)
      );
    }
  }

  setHeaderCartIconAccessibility() {
    const cartLink = document.querySelector("#cart-icon-bubble");
    cartLink.setAttribute("role", "button");
    cartLink.setAttribute("aria-haspopup", "dialog");
    cartLink.addEventListener("click", (event) => {
      event.preventDefault();
      this.open(cartLink);
    });
    cartLink.addEventListener("keydown", (event) => {
      if (event.code.toUpperCase() === "SPACE") {
        event.preventDefault();
        this.open(cartLink);
      }
    });
  }

  open(triggeredBy) {
    if (this.header) this.header.preventHide = true;
    if (triggeredBy) this.setActiveElement(triggeredBy);

    const cartDrawerSummaries = this.querySelectorAll('[id^="Details-"] summary');

    cartDrawerSummaries.forEach((summary) => {
      if (!summary.hasAttribute("role")) {
        this.setSummaryAccessibility(summary);
      }
    });

    // <CHANGE> Reset trap focus flag
    this.trapFocusSetup = false;

    setTimeout(() => {
      this.classList.add("animate", "active");
      this.addCloseButtonHandlers();
    });

    // <CHANGE> Added fallback timeout for trap focus
    const setupTrapFocus = () => {
      if (this.trapFocusSetup) return; // Prevent duplicate setup
      this.trapFocusSetup = true;

      const containerToTrapFocusOn = this.querySelector(".drawer__inner");
      const focusElement =
        containerToTrapFocusOn?.querySelector(".drawer__close");

      if (containerToTrapFocusOn && focusElement) {
        trapFocus(containerToTrapFocusOn, focusElement);
      }
    };

    // Try with transitionend
    this.addEventListener("transitionend", setupTrapFocus, { once: true });

    // Fallback if transitionend doesn't fire
    setTimeout(setupTrapFocus, 400);

    document.body.classList.add("overflow-hidden-drawer");
  }

  close() {
    this.classList.remove("active");

    // <CHANGE> Only remove trap focus if it was set up
    if (this.trapFocusSetup && this.activeElement) {
      removeTrapFocus(this.activeElement);
      this.trapFocusSetup = false;
    }

    document.body.classList.remove("overflow-hidden-drawer");
    if (this.header) this.header.preventHide = false;
  }

  setSummaryAccessibility(cartDrawerNote) {
    cartDrawerNote.setAttribute("role", "button");
    cartDrawerNote.setAttribute("aria-expanded", "false");

    if (cartDrawerNote.nextElementSibling.getAttribute("id")) {
      cartDrawerNote.setAttribute(
        "aria-controls",
        cartDrawerNote.nextElementSibling.id
      );
    }

    cartDrawerNote.addEventListener("click", (event) => {
      event.currentTarget.setAttribute(
        "aria-expanded",
        !event.currentTarget.closest("details").hasAttribute("open")
      );
    });

    cartDrawerNote.parentElement.addEventListener("keyup", onKeyUpEscape);
  }

  renderContents(parsedState) {
    this.querySelector(".drawer__inner").classList.contains("is-empty") &&
      this.querySelector(".drawer__inner").classList.remove("is-empty");
    this.productId = parsedState.id;
    this.getSectionsToRender().forEach((section) => {
      const sectionElement = section.selector
        ? document.querySelector(section.selector)
        : document.getElementById(section.id);
      sectionElement.innerHTML = this.getSectionInnerHTML(
        parsedState.sections[section.id],
        section.selector
      );
    });

    setTimeout(() => {
      this.querySelector("#CartDrawer-Overlay").addEventListener(
        "click",
        this.close.bind(this)
      );
      this.open();
    });
  }

  getSectionInnerHTML(html, selector = ".shopify-section") {
    return new DOMParser()
      .parseFromString(html, "text/html")
      .querySelector(selector).innerHTML;
  }

  getSectionsToRender() {
    return [
      {
        id: "cart-drawer",
        selector: "#CartDrawer",
      },
      {
        id: "cart-icon-bubble",
      },
    ];
  }

  getSectionDOM(html, selector = ".shopify-section") {
    return new DOMParser()
      .parseFromString(html, "text/html")
      .querySelector(selector);
  }

  setActiveElement(element) {
    this.activeElement = element;
  }
}

customElements.define("cart-drawer", CartDrawer);

class CartDrawerItems extends CartItems {
  getSectionsToRender() {
    return [
      {
        id: "CartDrawer",
        section: "cart-drawer",
        selector: ".drawer__inner",
      },
      {
        id: "cart-icon-bubble",
        section: "cart-icon-bubble",
        selector: ".shopify-section",
      },
    ];
  }
}

customElements.define("cart-drawer-items", CartDrawerItems);
