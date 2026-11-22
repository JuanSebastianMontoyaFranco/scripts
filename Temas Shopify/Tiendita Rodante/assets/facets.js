class FacetFiltersForm extends HTMLElement {
  constructor() {
    super();
    this.onActiveFilterClick = this.onActiveFilterClick.bind(this);

    this.debouncedOnSubmit = debounce((event) => {
      this.onSubmitHandler(event);
    }, 500);

    this.facetForm = this.querySelector("form");
    this.facetForm?.addEventListener(
      "input",
      this.debouncedOnSubmit.bind(this)
    );

    const facetWrapper = this.querySelector("#FacetsWrapperDesktop");
    if (facetWrapper) facetWrapper.addEventListener("keyup", onKeyUpEscape);

    if (this.querySelector(".facet-sort-select")) {
      this.initCustomSortSelect();
    }
  }

  static setListeners() {
    const onHistoryChange = (event) => {
      const searchParams = event.state
        ? event.state.searchParams
        : FacetFiltersForm.searchParamsInitial;
      if (searchParams === FacetFiltersForm.searchParamsPrev) return;
      FacetFiltersForm.renderPage(searchParams, null, false);
    };
    window.addEventListener("popstate", onHistoryChange);
  }

  static toggleActiveFacets(disable = true) {
    document.querySelectorAll(".js-facet-remove").forEach((element) => {
      element.classList.toggle("disabled", disable);
    });
  }

  static renderPage(searchParams, event, updateURLHash = true) {
    FacetFiltersForm.searchParamsPrev = searchParams;
    const sections = FacetFiltersForm.getSections();
    const countContainer = document.getElementById("ProductCount");
    const countContainerDesktop = document.getElementById(
      "ProductCountDesktop"
    );
    document
      .getElementById("ProductGridContainer")
      .querySelector(".collection")
      .classList.add("loading");
    if (countContainer) {
      countContainer.classList.add("loading");
    }
    if (countContainerDesktop) {
      countContainerDesktop.classList.add("loading");
    }

    sections.forEach((section) => {
      const url = `${window.location.pathname}?section_id=${section.section}&${searchParams}`;
      const filterDataUrl = (element) => element.url === url;

      FacetFiltersForm.filterData.some(filterDataUrl)
        ? FacetFiltersForm.renderSectionFromCache(filterDataUrl, event)
        : FacetFiltersForm.renderSectionFromFetch(url, event);
    });

    if (updateURLHash) FacetFiltersForm.updateURLHash(searchParams);
  }

  static renderSectionFromFetch(url, event) {
    fetch(url)
      .then((response) => response.text())
      .then((responseText) => {
        const html = responseText;
        FacetFiltersForm.filterData = [
          ...FacetFiltersForm.filterData,
          { html, url },
        ];
        FacetFiltersForm.renderFilters(html, event);
        FacetFiltersForm.renderProductGridContainer(html);
        FacetFiltersForm.renderProductCount(html);
      });
  }

  static renderSectionFromCache(filterDataUrl, event) {
    const html = FacetFiltersForm.filterData.find(filterDataUrl).html;
    FacetFiltersForm.renderFilters(html, event);
    FacetFiltersForm.renderProductGridContainer(html);
    FacetFiltersForm.renderProductCount(html);
  }

  static renderProductGridContainer(html) {
    document.getElementById("ProductGridContainer").innerHTML = new DOMParser()
      .parseFromString(html, "text/html")
      .getElementById("ProductGridContainer").innerHTML;

    if (
      document.querySelector(".js-load-more") ||
      document.querySelector(".js-infinite-scroll")
    ) {
      loadMore();
    }

    colorSwatches();
    applyViewCard();
    renderProductCardSpecs();
    initButtonsAnimation();
    document.querySelectorAll("facet-filters-form").forEach((facet) => {
      if (facet && typeof facet.initCustomSortSelect === "function") {
        facet.initCustomSortSelect();
      }
    });
    initButtonsAnimation(document.getElementById("ProductGridContainer"));
  }

  static renderProductCount(html) {
    const parsedHTML = new DOMParser().parseFromString(html, "text/html");
    const count = parsedHTML.getElementById("ProductCount")?.innerHTML;
    const countDesktop = parsedHTML.getElementById(
      "ProductCountDesktop"
    )?.innerHTML;
    const container = document.getElementById("ProductCount");
    const containerDesktop = document.getElementById("ProductCountDesktop");
    const containerMobile = document.querySelector(".mobile-facets__count");
    container.innerHTML = count;
    container.classList.remove("loading");
    if (containerDesktop) {
      containerDesktop.innerHTML = countDesktop;
      containerDesktop.classList.remove("loading");
    }
    if (containerMobile) {
      containerMobile.innerHTML = count;
    }
  }

  static renderFilters(html, event) {
    const parsedHTML = new DOMParser().parseFromString(html, "text/html");

    const facetDetailsElements = parsedHTML.querySelectorAll(
      "#FacetFiltersForm .js-filter, #FacetFiltersFormMobile .js-filter, #FacetFiltersPillsForm .js-filter"
    );
    const matchesIndex = (element) => {
      const jsFilter = event ? event.target.closest(".js-filter") : undefined;
      return jsFilter
        ? element.dataset.index === jsFilter.dataset.index
        : false;
    };
    const facetsToRender = Array.from(facetDetailsElements).filter(
      (element) => !matchesIndex(element)
    );
    const countsToRender = Array.from(facetDetailsElements).find(matchesIndex);

    facetsToRender.forEach((element) => {
      document.querySelector(
        `.js-filter[data-index="${element.dataset.index}"]`
      ).innerHTML = element.innerHTML;
    });

    FacetFiltersForm.renderActiveFacets(parsedHTML);
    FacetFiltersForm.renderAdditionalElements(parsedHTML);

    if (countsToRender)
      FacetFiltersForm.renderCounts(
        countsToRender,
        event.target.closest(".js-filter")
      );
  }

  static renderActiveFacets(html) {
    const activeFacetElementSelectors = [
      ".active-facets-mobile",
      ".active-facets-desktop",
    ];

    activeFacetElementSelectors.forEach((selector) => {
      const activeFacetsElement = html.querySelector(selector);
      if (!activeFacetsElement) return;
      document.querySelector(selector).innerHTML =
        activeFacetsElement.innerHTML;
    });

    FacetFiltersForm.toggleActiveFacets(false);
  }

  static renderAdditionalElements(html) {
    const mobileElementSelectors = [
      ".mobile-facets__open",
      ".mobile-facets__count",
      ".sorting",
    ];

    mobileElementSelectors.forEach((selector) => {
      if (!html.querySelector(selector)) return;
      document.querySelector(selector).innerHTML =
        html.querySelector(selector).innerHTML;
    });

    document
      .getElementById("FacetFiltersFormMobile")
      .closest("menu-drawer")
      .bindEvents();
  }

  static renderCounts(source, target) {
    const targetElement = target.querySelector(".facets__selected");
    const sourceElement = source.querySelector(".facets__selected");

    const targetElementAccessibility = target.querySelector(".facets__summary");
    const sourceElementAccessibility = source.querySelector(".facets__summary");

    if (sourceElement && targetElement) {
      target.querySelector(".facets__selected").outerHTML =
        source.querySelector(".facets__selected").outerHTML;
    }

    if (targetElementAccessibility && sourceElementAccessibility) {
      target.querySelector(".facets__summary").outerHTML =
        source.querySelector(".facets__summary").outerHTML;
    }
  }

  static updateURLHash(searchParams) {
    history.pushState(
      { searchParams },
      "",
      `${window.location.pathname}${searchParams && "?".concat(searchParams)}`
    );
  }

  static getSections() {
    return [
      {
        section: document.getElementById("product-grid").dataset.id,
      },
    ];
  }

  createSearchParams(form) {
    const formData = new FormData(form);
    return new URLSearchParams(formData).toString();
  }

  onSubmitForm(searchParams, event) {
    FacetFiltersForm.renderPage(searchParams, event);
  }

  onSubmitHandler(event) {
    event.preventDefault();
    const sortFilterForms = document.querySelectorAll(
      "facet-filters-form form"
    );
    if (event.srcElement.className == "mobile-facets__checkbox") {
      const searchParams = this.createSearchParams(
        event.target.closest("form")
      );
      this.onSubmitForm(searchParams, event);
    } else {
      const forms = [];
      const isMobile =
        event.target.closest("form").id === "FacetFiltersFormMobile";

      sortFilterForms.forEach((form) => {
        if (!isMobile) {
          if (
            form.id === "FacetSortForm" ||
            form.id === "FacetFiltersForm" ||
            form.id === "FacetSortDrawerForm"
          ) {
            const noJsElements = document.querySelectorAll(".no-js-list");
            noJsElements.forEach((el) => el.remove());
            forms.push(this.createSearchParams(form));
          }
        } else if (form.id === "FacetFiltersFormMobile") {
          forms.push(this.createSearchParams(form));
        }
      });
      this.onSubmitForm(forms.join("&"), event);
    }
  }

  onActiveFilterClick(event) {
    event.preventDefault();
    FacetFiltersForm.toggleActiveFacets();
    const url =
      event.currentTarget.href.indexOf("?") == -1
        ? ""
        : event.currentTarget.href.slice(
            event.currentTarget.href.indexOf("?") + 1
          );
    FacetFiltersForm.renderPage(url);
  }

  initCustomSortSelect() {
    const selectWrapper = this.querySelector(".facet-sort-select");
    if (!selectWrapper) return;
    // prevent duplicate initialization when called multiple times
    if (selectWrapper.dataset.sortInitialized === "true") return;
    const customSelect = this.querySelector(".facet-sort-select__select");
    const currentValue = this.querySelector(".facet-sort-select__current");
    const customOptions = this.querySelectorAll(".facet-sort-select__option");
    const hiddenInput = this.querySelector(".facet-sort-select__input");

    customSelect.addEventListener("click", () => {
      selectWrapper.classList.toggle("open");
    });

    customSelect.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        selectWrapper.classList.toggle("open");
      }
    });

    const onSelectOption = (event) => {
      event.preventDefault();

      if (event.target.matches("[data-value]")) {
        const option = event.target;
        const value = event.target.dataset.value;
        const text = event.target.textContent;

        hiddenInput.value = value;
        this.facetForm?.dispatchEvent(new Event("input", { bubbles: true }));

        currentValue.textContent = text;
        customOptions.forEach((option) => {
          option.classList.remove("facet-sort-select__option--selected");
        });
        option.classList.add("facet-sort-select__option--selected");
        selectWrapper.classList.remove("open");
      }
    };

    customOptions.forEach((option) => {
      option.addEventListener("click", (event) => onSelectOption(event));

      option.addEventListener("keydown", (event) => {
        if (event.key === "Enter") {
          onSelectOption(event);
        }
      });
    });

    document.addEventListener("click", (event) => {
      if (!selectWrapper.contains(event.target)) {
        selectWrapper.classList.remove("open");
      }
    });

    selectWrapper.dataset.sortInitialized = "true";
  }
}

FacetFiltersForm.filterData = [];
FacetFiltersForm.searchParamsInitial = window.location.search.slice(1);
FacetFiltersForm.searchParamsPrev = window.location.search.slice(1);
customElements.define("facet-filters-form", FacetFiltersForm);
FacetFiltersForm.setListeners();

class PriceRange extends HTMLElement {
  constructor() {
    super();
    this.querySelectorAll("input").forEach((element) =>
      element.addEventListener("change", this.onRangeChange.bind(this))
    );
    this.setMinAndMaxValues();
    this.controlSlider();
    this.controlInput();
  }

  onRangeChange(event) {
    this.adjustToValidValues(event.currentTarget);
    this.setMinAndMaxValues();
  }

  setMinAndMaxValues() {
    const inputs = this.querySelectorAll("input");
    const minInput = inputs[0];
    const maxInput = inputs[1];
    if (maxInput.value) minInput.setAttribute("max", maxInput.value);
    if (minInput.value) maxInput.setAttribute("min", minInput.value);
    if (minInput.value === "") maxInput.setAttribute("min", 0);
    if (maxInput.value === "")
      minInput.setAttribute("max", maxInput.getAttribute("max"));
  }

  adjustToValidValues(input) {
    const value = Number(input.value);
    const min = Number(input.getAttribute("min"));
    const max = Number(input.getAttribute("max"));

    if (value < min) input.value = min;
    if (value > max) input.value = max;
  }

  fillSlider() {
    const inputRangeWrappers = document.querySelectorAll(
      ".facets__price .facets__range"
    );
    inputRangeWrappers.forEach((inputWrapper) => {
      const inputsRange = inputWrapper.querySelectorAll(".field__range");
      const inputRangeFrom = inputsRange[0];
      const inputRangeTo = inputsRange[1];

      const range = inputRangeTo.max - inputRangeTo.min;
      const fromCurrent = inputRangeFrom.value - inputRangeTo.min;
      const toCurrent = inputRangeTo.value - inputRangeTo.min;
      const minRange = (fromCurrent / range) * 100;
      const maxRange = (toCurrent / range) * 100;

      inputWrapper.setAttribute(
        "style",
        `--range-min: ${minRange}%; --range-max: ${maxRange}%`
      );
    });
  }

  controlSlider() {
    const inputRangeWrappers = document.querySelectorAll(
      ".facets__price .facets__range"
    );
    const inputNumberWrappers = document.querySelectorAll(
      ".facets__price .facets__price-wrapper"
    );

    inputRangeWrappers.forEach((inputWrapper, index) => {
      const inputsRange = inputWrapper.querySelectorAll(".field__range");
      const inputRangeFrom = inputsRange[0];
      const inputRangeTo = inputsRange[1];
      const inputNumberFrom =
        inputNumberWrappers[index].querySelectorAll(".field__input")[0];
      const inputNumberTo =
        inputNumberWrappers[index].querySelectorAll(".field__input")[1];

      inputRangeFrom.oninput = () => {
        const from = Number(inputRangeFrom.value);
        const to = Number(inputRangeTo.value);
        if (from > to) {
          inputRangeFrom.value = to;
          inputNumberFrom.value = to;
        } else {
          inputNumberFrom.value = from;
        }

        this.fillSlider();
      };

      if (Number(inputRangeTo.value) <= 0) {
        inputRangeTo.style.zIndex = 2;
      } else {
        inputRangeTo.style.zIndex = 0;
      }

      inputRangeTo.oninput = () => {
        const from = Number(inputRangeFrom.value);
        const to = Number(inputRangeTo.value);
        if (from <= to) {
          inputRangeTo.value = to;
          inputNumberTo.value = to;
        } else {
          inputNumberTo.value = from;
          inputRangeTo.value = from;
        }

        if (Number(inputRangeTo.value) <= 0) {
          inputRangeTo.style.zIndex = 2;
        } else {
          inputRangeTo.style.zIndex = 0;
        }

        this.fillSlider();
      };
    });
  }

  controlInput() {
    const inputRangeWrappers = document.querySelectorAll(
      ".facets__price .facets__range"
    );
    const inputNumberWrappers = document.querySelectorAll(
      ".facets__price .facets__price-wrapper"
    );

    inputNumberWrappers.forEach((inputWrapper, index) => {
      const inputsNumber = inputWrapper.querySelectorAll(".field__input");
      const inputNumberFrom = inputsNumber[0];
      const inputNumberTo = inputsNumber[1];
      const inputRangeFrom =
        inputRangeWrappers[index].querySelectorAll(".field__range")[0];
      const inputRangeTo =
        inputRangeWrappers[index].querySelectorAll(".field__range")[1];

      inputNumberFrom.oninput = () => {
        const from = Number(inputNumberFrom.value);
        const to = Number(inputNumberTo.value);
        if (to && from > to) {
          inputRangeFrom.value = to;
        } else {
          inputRangeFrom.value = from;
        }

        this.fillSlider();
      };

      inputNumberTo.oninput = () => {
        const from = Number(inputNumberFrom.value);
        const to = Number(inputNumberTo.value);
        if (from <= to) {
          inputRangeTo.value = to;
        } else {
          inputRangeTo.value = from;
        }

        this.fillSlider();
      };

      inputNumberFrom.onblur = () => {
        const from = Number(inputNumberFrom.value);
        const min = Number(inputNumberFrom.getAttribute("min"));
        if (from < min) {
          inputRangeFrom.value = min;
          inputNumberFrom.value = min;
        } else {
          inputRangeFrom.value = from;
        }

        this.fillSlider();
      };

      inputNumberTo.onblur = () => {
        const to = Number(inputNumberTo.value);
        const max = Number(inputNumberTo.getAttribute("max"));
        if (to > max) {
          inputRangeTo.value = max;
          inputNumberTo.value = max;
        } else {
          inputRangeTo.value = to;
        }

        this.fillSlider();
      };
    });
  }
}

customElements.define("price-range", PriceRange);

class FacetRemove extends HTMLElement {
  constructor() {
    super();
    const facetLink = this.querySelector("a");
    facetLink.setAttribute("role", "button");
    facetLink.addEventListener("click", this.closeFilter.bind(this));
    facetLink.addEventListener("keyup", (event) => {
      event.preventDefault();
      if (event.code.toUpperCase() === "SPACE") this.closeFilter(event);
    });
  }

  closeFilter(event) {
    event.preventDefault();
    const form =
      this.closest("facet-filters-form") ||
      document.querySelector("facet-filters-form");
    form.onActiveFilterClick(event);
  }
}

customElements.define("facet-remove", FacetRemove);
(function () {
  const initSection = () => {
    const toggle = document.getElementById("facetToggle");
    const facetsWrapper = document.getElementById("FacetsWrapperDesktop");
    const activeFacetsDesktop = document.querySelector(
      ".active-facets-desktop__wrapper"
    );
    const productCountHeader = document.querySelector(
      ".facets__product-count--header"
    );

    if (toggle && facetsWrapper) {
      if (toggle.checked) {
        productCountHeader.style.setProperty(
          "color",
          "rgb(var(--color-accent))",
          "important"
        );
      }
      toggle.addEventListener("change", () => {
        facetsWrapper.classList.toggle("is-hidden", !toggle.checked);

        if (!toggle.checked) {
          activeFacetsDesktop.style.display = "none";

          productCountHeader.style.setProperty(
            "color",
            "rgb(var(--color-foreground-secondary))",
            "important"
          );
        } else {
          activeFacetsDesktop.style.display = "flex";

          productCountHeader.style.setProperty(
            "color",
            "rgb(var(--color-accent))",
            "important"
          );
        }
      });

      const toggleWrapper = document.querySelector(".toggle-switch");

      toggleWrapper.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();

          toggle.checked = !toggle.checked;
          facetsWrapper.classList.toggle("is-hidden", !toggle.checked);

          if (!toggle.checked) {
            activeFacetsDesktop.style.display = "none";
            productCountHeader.style.setProperty(
              "color",
              "rgb(var(--color-foreground-secondary))",
              "important"
            );
          } else {
            activeFacetsDesktop.style.display = "flex";
            productCountHeader.style.setProperty(
              "color",
              "rgb(var(--color-accent))",
              "important"
            );
          }
        }
      });
    }
  };

  document.addEventListener("DOMContentLoaded", function () {
    initSection();
    document.addEventListener("shopify:section:load", function () {
      initSection();
    });
  });
})();

function applyViewCard() {
  const switcher = document.querySelector(".view-switcher");
  const grid = document.querySelector(".product-grid-grid");
  const list = document.querySelector(".product-grid-list");
  const cardsTypeWrapper = document.querySelector(
    ".product-cards-type-wrapper"
  );

  const savedView = localStorage.getItem("productViewType") || "grid";
  applyView(savedView);

  switcher.addEventListener("click", (event) => {
    const button = event.target.closest(".switch-button");
    if (!button) return;

    switcher.querySelectorAll(".switch-button").forEach((btn) => {
      btn.classList.remove("active");
    });

    button.classList.add("active");

    const viewType = button.getAttribute("data-view-switcher-type");
    localStorage.setItem("productViewType", viewType);
    applyView(viewType);
  });

  function applyView(viewType) {
    switcher.querySelectorAll(".switch-button").forEach((btn) => {
      btn.classList.remove("active");
    });

    const activeBtn = switcher.querySelector(
      `[data-view-switcher-type="${viewType}"]`
    );
    activeBtn?.classList.add("active");

    if (viewType === "grid") {
      if (list) {
        list.style.display = "none";
        list.remove();
      }
      if (grid) {
        grid.style.display = "";
        cardsTypeWrapper.appendChild(grid);
      }
    } else {
      if (grid) {
        grid.style.display = "none";
        grid.remove();
      }
      if (list) {
        list.style.display = "";
        cardsTypeWrapper.appendChild(list);
      }
    }
  }
}

applyViewCard();
