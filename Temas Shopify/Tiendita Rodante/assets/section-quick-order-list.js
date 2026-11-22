class BulkAdd extends HTMLElement {
  static ASYNC_REQUEST_DELAY = 250;

  constructor() {
    super();
    // Pending { id, quantity } items waiting to be sent
    this.queue = [];
    this.setRequestStarted(false);
    this.ids = [];
  }

  // Adds an item to the queue and starts polling; fires sendRequest once no request is in flight
  startQueue(id, quantity) {
    this.queue.push({ id, quantity });

    const interval = setInterval(() => {
      if (this.queue.length > 0) {
        if (!this.requestStarted) {
          this.sendRequest(this.queue);
        }
      } else {
        clearInterval(interval);
      }
    }, BulkAdd.ASYNC_REQUEST_DELAY);
  }

  // Deduplicates the queue snapshot into { variantId: qty } and fires updateMultipleQty
  sendRequest(queue) {
    this.setRequestStarted(true);
    const items = {};

    queue.forEach((queueItem) => {
      items[parseInt(queueItem.id)] = queueItem.quantity;
    });
    this.queue = this.queue.filter(
      (queueElement) => !queue.includes(queueElement)
    );

    this.updateMultipleQty(items);
  }

  setRequestStarted(requestStarted) {
    this._requestStarted = requestStarted;
  }

  get requestStarted() {
    return this._requestStarted;
  }

  resetQuantityInput(id) {
    const input = this.querySelector(`#Quantity-${id}`);
    input.value = input.getAttribute("value");
  }

  setValidity(event, index, message) {
    event.target.setCustomValidity(message);
    event.target.reportValidity();
    this.resetQuantityInput(index);
    event.target.select();
  }

  // Validates min/max/step constraints — shows native browser error on failure, or enqueues the update
  validateQuantity(event) {
    const inputValue = parseInt(event.target.value);
    const index = event.target.dataset.index;

    if (inputValue < event.target.dataset.min) {
      this.setValidity(
        event,
        index,
        window.quickOrderListStrings.min_error.replace(
          "[min]",
          event.target.dataset.min
        )
      );
    } else if (inputValue > parseInt(event.target.max)) {
      this.setValidity(
        event,
        index,
        window.quickOrderListStrings.max_error.replace(
          "[max]",
          event.target.max
        )
      );
    } else if (inputValue % parseInt(event.target.step) != 0) {
      this.setValidity(
        event,
        index,
        window.quickOrderListStrings.step_error.replace(
          "[step]",
          event.target.step
        )
      );
    } else {
      event.target.setCustomValidity("");
      event.target.reportValidity();
      event.target.setAttribute("value", inputValue);
      this.toggleLoading(true, this, index);
      this.startQueue(index, inputValue);
    }
  }

  getSectionInnerHTML(html, selector) {
    return new DOMParser()
      .parseFromString(html, "text/html")
      .querySelector(selector).innerHTML;
  }
}

class QuickOrderList extends BulkAdd {
  cartUpdateUnsubscriber = undefined;
  hasPendingQuantityUpdate = false;

  constructor() {
    super();
    // Cache sticky header height and type for scroll offset compensation
    this.stickyHeaderElement = document.querySelector("sticky-header");
    if (this.stickyHeaderElement) {
      this.stickyHeader = {
        height: this.stickyHeaderElement.offsetHeight,
        type: `${this.stickyHeaderElement.getAttribute("data-sticky-type")}`,
      };
    }

    // Cache totals bar and its Y-position threshold for scroll collision detection
    this.totalBar = this.getTotalBar();
    if (this.totalBar) {
      this.totalBarPosition = window.innerHeight - this.totalBar.offsetHeight;

      this.handleResize = this.handleResize.bind(this);
      window.addEventListener("resize", this.handleResize);
    }

    this.querySelector("form").addEventListener("submit", (event) =>
      event.preventDefault()
    );
  }

  connectedCallback() {
    this.cartUpdateUnsubscriber = subscribe(
      PUB_SUB_EVENTS.cartUpdate,
      async (event) => {
        // skip if cart event was triggered by this section
        if (event.source === this.id) return;

        this.toggleTableLoading(true);
        await this.refresh();
        this.toggleTableLoading(false);
      }
    );

    this.initEventListeners();
  }

  disconnectedCallback() {
    this.cartUpdateUnsubscriber?.();
    window.removeEventListener("resize", this.handleResize);
  }

  // Recalculates scroll positions after window resize
  handleResize() {
    this.totalBarPosition = window.innerHeight - this.totalBar.offsetHeight;
    this.stickyHeader.height = this.stickyHeaderElement
      ? this.stickyHeaderElement.offsetHeight
      : 0;
  }

  initEventListeners() {
    this.querySelectorAll(".pagination__item").forEach((link) => {
      link.addEventListener("click", async (event) => {
        if (!event.currentTarget.href) return;

        event.preventDefault();
        event.stopPropagation();

        const url = new URL(event.currentTarget.href);

        this.toggleTableLoading(true);
        await this.refresh(url.searchParams.get("page") || "1");
        this.scrollTop();
        this.toggleTableLoading(false);
      });
    });

    this.querySelector(".quick-order-list__form").addEventListener(
      "keyup",
      this.handleScrollIntoView.bind(this)
    );

    this.quickOrderListTable.addEventListener(
      "keydown",
      this.handleSwitchVariantOnEnter.bind(this)
    );

    this.initVariantEventListeners();
  }

  // Re-attaches qty input and remove button listeners after DOM rebuild
  initVariantEventListeners() {
    this.allInputsArray = Array.from(
      this.querySelectorAll('input[type="number"]')
    );

    this.querySelectorAll("quantity-input").forEach((qty) => {
      const debouncedOnChange = debounce(
        this.onChange.bind(this),
        BulkAdd.ASYNC_REQUEST_DELAY
      );
      qty.addEventListener("change", (event) => {
        this.hasPendingQuantityUpdate = true;
        debouncedOnChange(event);
      });
    });

    this.querySelectorAll(".quick-variant-item__remove").forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        this.toggleLoading(true, this, button.dataset.index);
        this.startQueue(button.dataset.index, 0);
      });
    });
  }

  get currentPage() {
    return this.querySelector(".pagination-wrapper")?.dataset?.page ?? "1";
  }

  // Returns array of variant IDs currently in the cart for this product (used by "remove all")
  get cartVariantsForProduct() {
    return JSON.parse(
      this.querySelector("[data-cart-contents]")?.innerHTML || "[]"
    );
  }

  // qty=0 immediately removes the item; any other value goes through min/max/step validation
  onChange(event) {
    const inputValue = parseInt(event.target.value);
    this.cleanErrorMessageOnType(event);
    if (inputValue == 0) {
      event.target.setAttribute("value", inputValue);
      this.toggleLoading(true, this, event.target.dataset.index);
      this.startQueue(event.target.dataset.index, inputValue);
    } else {
      this.validateQuantity(event);
    }
  }

  // Clears the native validity tooltip on the next keydown so it doesn't linger while the user types
  cleanErrorMessageOnType(event) {
    const handleKeydown = () => {
      event.target.setCustomValidity(" ");
      event.target.reportValidity();
      event.target.removeEventListener("keydown", handleKeydown);
    };

    event.target.addEventListener("keydown", handleKeydown);
  }

  // Silent validity check (no UI) — used before allowing Enter key focus jump to the next input
  validateInput(target) {
    const targetValue = parseInt(target.value);
    const targetMin = parseInt(target.dataset.min);
    const targetStep = parseInt(target.step);

    if (target.max) {
      return (
        targetValue == 0 ||
        (targetValue >= targetMin &&
          targetValue <= parseInt(target.max) &&
          targetValue % targetStep == 0)
      );
    } else {
      return (
        targetValue == 0 ||
        (targetValue >= targetMin && targetValue % targetStep == 0)
      );
    }
  }

  get quickOrderListTable() {
    return this.querySelector(".quick-order-list__table");
  }

  getSectionsToRender() {
    return [
      {
        id: this.id,
        section: this.dataset.section,
        selector: `#${this.id}`,
      },
      {
        id: `quick-order-list-live-region-text-${this.dataset.productId}`,
        section: "cart-live-region-text",
        selector: ".shopify-section",
      },
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

  toggleTableLoading(enable) {
    this.quickOrderListTable.classList.toggle(
      "quick-order-list__container--disabled",
      enable
    );
    this.toggleLoading(enable);
  }

  // Full section re-fetch — called when an external cart update is detected
  async refresh(pageNumber = null) {
    const url = this.dataset.url || window.location.pathname;

    return fetch(
      `${url}?section_id=${this.dataset.section}&page=${pageNumber || this.currentPage}`
    )
      .then((response) => response.text())
      .then((responseText) => {
        const html = new DOMParser().parseFromString(responseText, "text/html");
        const responseQuickOrderList = html.querySelector(`#${this.id}`);

        if (!responseQuickOrderList) {
          return;
        }

        this.innerHTML = responseQuickOrderList.innerHTML;
        this.initEventListeners();
      })
      .catch((e) => {
        console.error(e);
      });
  }

  // Patches DOM from cart API response; skips re-render if queue is still active; preserves keyboard focus
  renderSections(parsedState) {
    const { items, sections } = parsedState;

    this.getSectionsToRender().forEach(({ id, selector, section }) => {
      const sectionElement = document.getElementById(id);
      if (!sectionElement) return;
      if (!sections[section]) return;

      const newSection = new DOMParser()
        .parseFromString(sections[section], "text/html")
        .querySelector(selector);
      if (!newSection) return;

      if (section === this.dataset.section) {
        if (this.queue.length > 0 || this.hasPendingQuantityUpdate) return;

        const focusTarget = document.activeElement?.dataset?.target;

        const total = this.getTotalBar();
        if (total) {
          total.innerHTML = newSection.querySelector(
            ".quick-order-list__total"
          ).innerHTML;
        }

        const table = this.quickOrderListTable;
        const newTable = newSection.querySelector(".quick-order-list__table");

        // only update variants if they are from the active page
        const shouldUpdateVariants =
          this.currentPage ===
          (newSection.querySelector(".pagination-wrapper")?.dataset.page ??
            "1");
        if (newTable && shouldUpdateVariants) {
          table.innerHTML = newTable.innerHTML;

          const newFocusTarget = this.querySelector(
            `[data-target='${focusTarget}']`
          );
          if (newFocusTarget) {
            newFocusTarget?.focus({ preventScroll: true });
          }

          this.initVariantEventListeners();
        }
      } else if (section === "cart-drawer") {
        sectionElement
          .closest("cart-drawer")
          ?.classList.toggle("is-empty", items.length === 0);
        sectionElement.querySelector(selector).innerHTML = newSection.innerHTML;
      } else {
        sectionElement.innerHTML = newSection.innerHTML;
      }
    });
  }

  getTotalBar() {
    return this.querySelector(".quick-order-list__total");
  }

  // Scrolls to the top of the section, offset by the sticky header height
  scrollTop() {
    const { top } = this.getBoundingClientRect();

    window.scrollTo({
      top: top + window.scrollY - (this.stickyHeader?.height || 0),
      behavior: "instant",
    });
  }

  // Scrolls the focused input into view if it's obscured by the sticky header or the totals bar
  scrollQuickOrderListTable(target) {
    const inputTopBorder = target.getBoundingClientRect().top;
    const inputBottomBorder = target.getBoundingClientRect().bottom;

    const stickyHeaderBottomBorder =
      this.stickyHeaderElement?.getBoundingClientRect().bottom;
    const totalBarCrossesInput = inputBottomBorder > this.totalBarPosition;
    const inputOutsideOfViewPort =
      inputBottomBorder <
      this.querySelector(".quick-variant-item__quantity-wrapper").offsetHeight;
    const stickyHeaderCrossesInput =
      this.stickyHeaderElement &&
      this.stickyHeader.type !== "on-scroll-up" &&
      this.stickyHeader.height > inputTopBorder;
    const stickyHeaderScrollupCrossesInput =
      this.stickyHeaderElement &&
      this.stickyHeader.type === "on-scroll-up" &&
      this.stickyHeader.height > inputTopBorder &&
      stickyHeaderBottomBorder > 0;

    if (
      totalBarCrossesInput ||
      inputOutsideOfViewPort ||
      stickyHeaderCrossesInput ||
      stickyHeaderScrollupCrossesInput
    ) {
      this.scrollToCenter(target);
    }
  }

  scrollToCenter(target) {
    target.scrollIntoView({
      block: "center",
      behavior: "smooth",
    });
  }

  handleScrollIntoView(event) {
    if (
      (event.key === "Tab" || event.key === "Enter") &&
      this.allInputsArray.length !== 1
    ) {
      this.scrollQuickOrderListTable(event.target);
    }
  }

  // Moves focus to the next (Shift: previous) input on Enter; validates the current input before jumping
  handleSwitchVariantOnEnter(event) {
    if (event.key !== "Enter" || event.target.tagName !== "INPUT") return;

    event.preventDefault();
    event.target.blur();

    if (!this.validateInput(event.target) || this.allInputsArray.length <= 1)
      return;

    const currentIndex = this.allInputsArray.indexOf(event.target);
    const offset = event.shiftKey ? -1 : 1;
    const nextIndex =
      (currentIndex + offset + this.allInputsArray.length) %
      this.allInputsArray.length;

    this.allInputsArray[nextIndex]?.select();
  }

  // POSTs cart/update.js with all queued items, re-renders affected sections, publishes cartUpdate event
  updateMultipleQty(items) {
    if (this.queue.length == 0) this.hasPendingQuantityUpdate = false;

    // Track IDs being updated so their per-row spinners can be hidden when the request completes
    this.pendingItemIds = Object.keys(items).map(String);
    this.toggleLoading(true);
    const url = this.dataset.url || window.location.pathname;

    const body = JSON.stringify({
      updates: items,
      sections: this.getSectionsToRender().map(({ section }) => section),
      sections_url: `${url}?page=${this.currentPage}`,
    });

    this.setErrorMessage();

    fetch(`${window.routes.cart_update_url}`, { ...fetchConfig(), ...{ body } })
      .then((response) => response.text())
      .then(async (state) => {
        const parsedState = JSON.parse(state);
        this.renderSections(parsedState);
        publish(PUB_SUB_EVENTS.cartUpdate, {
          source: this.id,
          cartData: parsedState,
        });
      })
      .catch((e) => {
        console.error(e);
        this.setErrorMessage(window.cartStrings.error);
      })
      .finally(() => {
        if (this.queue.length === 0) {
          this.toggleLoading(false);
          this.pendingItemIds?.forEach((id) =>
            this.toggleLoading(false, this, id)
          );
          this.pendingItemIds = [];
        }
        this.setRequestStarted(false);
      });
  }

  // Renders a cart error into all .quick-order-list__error slots; pass null to clear
  setErrorMessage(message = null) {
    this.errorMessageTemplate =
      this.errorMessageTemplate ??
      document
        .getElementById(`QuickOrderListErrorTemplate-${this.dataset.productId}`)
        .cloneNode(true);
    const errorElements = document.querySelectorAll(".quick-order-list__error");

    errorElements.forEach((errorElement) => {
      errorElement.innerHTML = "";
      if (!message) return;
      const updatedMessageElement = this.errorMessageTemplate.cloneNode(true);
      updatedMessageElement.content.querySelector(
        ".quick-order-list__error-message"
      ).innerText = message;
      errorElement.appendChild(updatedMessageElement.content);
    });
  }

  toggleLoading(loading, target = this, item = null) {
    target
      .querySelector("#shopping-cart-variant-item-status")
      .toggleAttribute("aria-hidden", !loading);
    target
      .querySelectorAll(
        ".quick-order-list__remove-all .loading-overlay__spinner"
      )
      ?.forEach((spinner) => {
        spinner.classList.toggle("hidden", !loading);
      });
    if (item) {
      const row = target.querySelector(`[data-variant-id="${item}"]`);
      const itemSpinner = row?.querySelector(
        `.quick-variant-item__totals-wrapper .loading-overlay__spinner`
      );
      if (itemSpinner) {
        itemSpinner.classList.toggle("hidden", !loading);
      }
    }
  }
}

if (!customElements.get("quick-order-list")) {
  customElements.define("quick-order-list", QuickOrderList);
}

class QuickOrderListRemoveAllButton extends HTMLElement {
  constructor() {
    super();
    this.quickOrderList = this.closest("quick-order-list");

    this.actions = {
      confirm: "confirm",
      remove: "remove",
      cancel: "cancel",
    };

    this.addEventListener("click", (event) => {
      event.preventDefault();
      if (this.dataset.action === this.actions.confirm) {
        this.toggleConfirmation(false, true);
      } else if (this.dataset.action === this.actions.remove) {
        const items = this.quickOrderList.cartVariantsForProduct.reduce(
          (acc, variantId) => ({ ...acc, [variantId]: 0 }),
          {}
        );

        this.quickOrderList.updateMultipleQty(items);
        this.toggleConfirmation(true, false);
      } else if (this.dataset.action === this.actions.cancel) {
        this.toggleConfirmation(true, false);
      }
    });
  }

  toggleConfirmation(showConfirmation, showInfo) {
    this.quickOrderList
      .querySelector(".quick-order-list__total-confirmation")
      .classList.toggle("hidden", showConfirmation);
    this.quickOrderList
      .querySelector(".quick-order-list__total-info")
      .classList.toggle("hidden", showInfo);
  }
}

if (!customElements.get("quick-order-list-remove-all-button")) {
  customElements.define(
    "quick-order-list-remove-all-button",
    QuickOrderListRemoveAllButton
  );
}
