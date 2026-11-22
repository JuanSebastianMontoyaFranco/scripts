(() => {
  let isScrolling = false;
  let scrollTimeout = null;

  const getScrollOffset = () => {
    const root = document.documentElement;
    const headerHeight =
      Number.parseInt(
        getComputedStyle(root).getPropertyValue("--header-height")
      ) || 0;
    const breadcrumbsHeight =
      Number.parseInt(
        getComputedStyle(root).getPropertyValue("--breadcrumbs-height")
      ) || 0;
    return headerHeight + breadcrumbsHeight;
  };

  const scrollToElement = (element) => {
    if (!element) return;
    const offset = getScrollOffset();
    const elementPosition =
      element.getBoundingClientRect().top + window.pageYOffset;
    const offsetPosition = elementPosition - offset;

    isScrolling = true;
    if (scrollTimeout) clearTimeout(scrollTimeout);
    scrollTimeout = setTimeout(() => {
      isScrolling = false;
    }, 1000);

    window.scrollTo({
      top: offsetPosition,
      behavior: "smooth",
    });
  };

  const scrollIntoView = (element) => {
    const container =
      element.closest(".product__info-navigations--container") ||
      element.closest(".popular-products__collection-tabs-list--wrapper");
    if (!container || !element) return;

    const containerRect = container.getBoundingClientRect();
    const elementRect = element.getBoundingClientRect();

    if (elementRect.left < containerRect.left) {
      const scrollAmount = elementRect.left - containerRect.left;
      const newScrollLeft = container.scrollLeft + scrollAmount;

      if (newScrollLeft < 50) {
        container.scrollTo({
          left: 0,
          behavior: "smooth",
        });
      } else {
        container.scrollBy({
          left: scrollAmount,
          behavior: "smooth",
        });
      }
    } else if (elementRect.right > containerRect.right) {
      const scrollAmount = elementRect.right - containerRect.right;
      const newScrollLeft = container.scrollLeft + scrollAmount;
      const maxScrollLeft = container.scrollWidth - container.clientWidth;

      if (maxScrollLeft - newScrollLeft < 50) {
        container.scrollTo({
          left: maxScrollLeft,
          behavior: "smooth",
        });
      } else {
        container.scrollBy({
          left: scrollAmount,
          behavior: "smooth",
        });
      }
    }
  };

  const observeHorizontalScroll = (el) => {
    if (!el) return;

    const checkScrollState = () => {
      const scrollLeft = el.scrollLeft;
      const maxScrollLeft = el.scrollWidth - el.clientWidth;

      const hasScroll = el.scrollWidth > el.clientWidth;
      const atStart = scrollLeft <= 1;
      const atEnd = scrollLeft >= maxScrollLeft - 2;

      if (hasScroll && !atStart) {
        el.classList.add("left-shadow");
      } else {
        el.classList.remove("left-shadow");
      }

      if (hasScroll && !atEnd) {
        el.classList.add("right-shadow");
      } else {
        el.classList.remove("right-shadow");
      }
    };

    el.addEventListener("scroll", checkScrollState);

    const ro = new ResizeObserver(() => {
      checkScrollState();
    });

    ro.observe(el);

    checkScrollState();
  };

  const sectionToTabMap = {
    description: "description", // About product
    specs_long: "specs_long", // Specs
    instructions: "instructions",
    "media-section": "all-features",
    "description-grid-section": "all-features",
    "image-with-text-section": "all-features",
    "testimonials-quote": "testimonials-quote", // Reviews
  };

  // Each .tabs-lava-lamp track (product page tabs, every popular-products
  // section) gets its own blob and listeners.
  const initTrack = (track) => {
    if (track.dataset.lavaLampInitialized) return;

    const blob = track.querySelector("#tabs-lava-lamp-blob");
    const navItems = Array.from(track.querySelectorAll("li")).filter(
      (item) => item !== blob
    );

    if (!blob || navItems.length === 0) return;
    track.dataset.lavaLampInitialized = "true";

    // With no active tab (e.g. a single collection, so no "All" tab) the blob
    // rests hidden at the first tab and slides in on hover.
    let activeItem = track.querySelector(".active");

    // ====== Blob ======
    const makeBlob = () => {
      const item = activeItem || navItems[0];
      blob.style.width = activeItem ? `${item.clientWidth}px` : "0px";
      blob.style.height = `${item.clientHeight}px`;
      blob.style.left = `${item.offsetLeft}px`;
      blob.style.top = `${item.offsetTop}px`;
      blob.style.opacity = activeItem ? "1" : "0";
      blob.classList.add("blob-initialized");
    };

    const animate = (to) => {
      navItems.forEach((item) => item.classList.remove("active-hover"));
      const target = to
        ? { left: to.offsetLeft, width: to.clientWidth, opacity: 1 }
        : { left: navItems[0].offsetLeft, width: 0, opacity: 0 };
      blob.animate(
        [
          {
            left: `${blob.offsetLeft}px`,
            width: `${blob.clientWidth}px`,
            opacity: getComputedStyle(blob).opacity,
          },
          {
            left: `${target.left}px`,
            width: `${target.width}px`,
            opacity: target.opacity,
          },
        ],
        {
          duration: 300,
          easing: "cubic-bezier(0.87,0,0.13,1)",
          fill: "forwards",
        }
      );
      to?.classList.add("active-hover");
    };

    const setActive = (li) => {
      if (!li || li === blob) return;
      navItems.forEach((item) => item.classList.remove("active"));
      li.classList.add("active");
      activeItem = li;
      animate(li);
      scrollIntoView(li);
    };

    observeHorizontalScroll(
      track.closest(".product__info-navigations--container")
    );
    observeHorizontalScroll(
      track.closest(".popular-products__collection-tabs-list--wrapper")
    );

    makeBlob();

    track.addEventListener("click", (e) => {
      const li = e.target.closest("li");
      if (!li || li === blob) return;

      const link = li.querySelector("a");
      const targetId = link?.getAttribute("href")?.replace("#", "");

      if (targetId) {
        e.preventDefault();

        if (targetId === "all-features") {
          const firstAllFeaturesSectionId = Object.keys(sectionToTabMap).find(
            (key) => sectionToTabMap[key] === "all-features"
          );

          const targetSection = document.getElementById(
            firstAllFeaturesSectionId
          );
          scrollToElement(targetSection);
        } else {
          const targetSection = document.getElementById(targetId);
          scrollToElement(targetSection);
        }
      }

      setActive(li);
    });

    if (window.innerWidth > 767) {
      track.addEventListener("mouseover", (e) => {
        const li = e.target.closest("li");
        if (!li || li === blob) return;
        animate(li);
      });
      track.addEventListener("mouseout", () => {
        navItems.forEach((item) => item.classList.remove("active-hover"));
        animate(activeItem);
      });
    }

    const hasSectionLinks = Object.keys(sectionToTabMap).some((id) =>
      track.querySelector(`a[href="#${sectionToTabMap[id]}"]`)
    );
    if (!hasSectionLinks) return;

    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting && !isScrolling) {
            const tabId = sectionToTabMap[entry.target.id];
            if (!tabId) return;

            const li = track
              .querySelector(`a[href="#${tabId}"]`)
              ?.closest("li");

            if (li) setActive(li);
          }
        });
      },
      { rootMargin: "0px 0px 0px 0px", threshold: 0.8 }
    );

    Object.keys(sectionToTabMap).forEach((id) => {
      const section = document.getElementById(id);
      if (section) observer.observe(section);
    });
  };

  const init = (root = document) => {
    root.querySelectorAll(".tabs-lava-lamp").forEach(initTrack);
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", () => init());
  } else {
    init();
  }

  document.addEventListener("shopify:section:load", (event) =>
    init(event.target)
  );
})();
