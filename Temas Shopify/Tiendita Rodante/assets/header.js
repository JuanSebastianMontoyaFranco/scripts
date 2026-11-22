(function () {
  const header = () => {
    const header = document.querySelector(".shopify-section-header");
    const headerWrapper = document.querySelector(".header-wrapper");
    const menu = document.querySelector(
      ".header__second-level .list-menu--inline"
    );
    const menuLinks = document.querySelectorAll(
      ".header__second-level .list-menu-item"
    );
    const menuLinksWithDropdown = document.querySelectorAll(
      ".header__second-level .list-menu--submenu, .header__second-level .list-menu--megamenu"
    );
    const megaMenuLink = document.querySelector(".list-menu--megamenu");
    const search = document.querySelector("details-modal.header__search");
    const searchModal = document.querySelector(
      "details-modal.header__search > details"
    );
    const allSubmenu = document.querySelectorAll(".header__submenu");
    const searchOpenBtn = document.querySelectorAll(".search-open-btn");
    const searchOpenBtnFooter = document.querySelector(
      ".search-open-btn--footer"
    );
    const searchInner = document.querySelectorAll(".header__search-inner");
    const searchPromoBg = document.querySelector(".search__modal");
    const searchPredictiveBg = document.querySelectorAll(".search-modal__form");
    const headerSimpleMenuLinks = document.querySelectorAll(
      ".simple_mega_menu .header__submenu-li"
    );

    const blob = document.getElementById("blob");
    const track = document.querySelector(
      ".header__second-level .list-menu--inline"
    );
    const navItems = document.querySelectorAll(
      ".header__second-level .list-menu--inline li:not(#blob)"
    );
    const activeItem = document
      .querySelector(
        ".header__second-level .list-menu--inline .header__active-menu-item"
      )
      ?.closest("li");

    let hasBlobPosition = false;

    if (blob) {
      blob.style.opacity = "0";
      blob.style.left = "0px";
      blob.style.top = "0px";
    }

    headerSimpleMenuLinks.forEach((link) => {
      link.addEventListener("mouseenter", (e) => {
        headerSimpleMenuLinks.forEach((el) => {
          if (el === e.target) {
            el.querySelector(".header__submenu")?.classList.add("show-submenu");
            el.querySelector(".header__submenu")?.classList.remove(
              "remove-submenu"
            );
          } else {
            el.querySelector(".header__submenu")?.classList.remove(
              "show-submenu"
            );
            el.querySelector(".header__submenu")?.classList.add(
              "remove-submenu"
            );
          }
        });
      });
    });

    const addHeaderOverlay = (trigger) => {
      trigger?.addEventListener("mouseover", () => {
        document.querySelector(".header__overlay").classList.add("visible");
      });
      trigger?.addEventListener("mouseout", () => {
        document.querySelector(".header__overlay").classList.remove("visible");
      });
    };

    addHeaderOverlay(megaMenuLink);

    const makeBlob = (target = activeItem) => {
      if (!target || !blob) return;

      blob.style.width = `${target.clientWidth}px`;
      blob.style.height = `${target.clientHeight}px`;
      blob.style.left = `${target.offsetLeft}px`;
      blob.style.top = `${target.offsetTop}px`;
      blob.style.opacity = "1";
      blob.style.backgroundColor = "rgb(var(--color-accent))";
    };

    const hideBlob = () => {
      if (!blob) return;
      blob.style.backgroundColor = "transparent";
    };

    const animate = (from, to) => {
      if (!from || !to) return;

      const keyframes = [
        {
          left: `${from.offsetLeft}px`,
          width: `${from.clientWidth}px`,
          top: `${from.offsetTop}px`,
          opacity: "1",
        },
        {
          left: `${to.offsetLeft}px`,
          width: `${to.clientWidth}px`,
          top: `${to.offsetTop}px`,
          opacity: "1",
        },
      ];

      const timing = {
        duration: 300,
        easing: "cubic-bezier(0.87, 0, 0.13, 1)",
        fill: "forwards",
      };

      blob.animate(keyframes, timing);
    };

    if (activeItem) {
      makeBlob(activeItem);
      blob.style.opacity = "1";
      hasBlobPosition = true;
    } else {
      hideBlob();
    }

    //track?.addEventListener("mouseover", (e) => {
    //	const targetItem = e.target.matches(".list-menu-item")
    //		? e.target
    //		: e.target.closest(".list-menu-item");

    //	if (targetItem && targetItem.id !== "blob") {
    //		blob.style.backgroundColor = "rgb(var(--color-accent))";
    //		animate(blob, targetItem);
    //	}
    //});

    track?.addEventListener("mouseover", (e) => {
      const targetItem = e.target.matches(".list-menu-item")
        ? e.target
        : e.target.closest(".list-menu-item");

      if (!targetItem || targetItem.id === "blob") return;

      blob.style.backgroundColor = "rgb(var(--color-accent))";

      if (!hasBlobPosition) {
        makeBlob(targetItem);

        blob.animate([{ opacity: 0 }, { opacity: 1 }], {
          duration: 300,
          easing: "ease-out",
          fill: "forwards",
        });

        hasBlobPosition = true;
        return;
      }

      animate(blob, targetItem);
    });

    track?.addEventListener("mouseleave", () => {
      if (activeItem) {
        animate(blob, activeItem);
      } else {
        hideBlob();
      }
    });

    window.addEventListener("resize", () => {
      if (activeItem) {
        makeBlob(activeItem);
      }
    });

    header.addEventListener("keydown", (e) => {
      if (e.code === "Escape") {
        searchInner.forEach((item) => {
          item.classList.remove("header__search-inner--open");
        });
        document.body.classList.remove(`overflow-hidden-search`);
      }
    });

    const annBar = document.querySelector(".section-announcement");
    let annBarObserver;
    const createAnnBarObserver = () => {
      if (annBarObserver) annBarObserver.disconnect();
      annBarObserver = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.isIntersecting) {
              const annBarHeight = annBar?.getBoundingClientRect().height || 0;
              document.documentElement.style.setProperty(
                "--ann-height",
                `${(annBarHeight * entry.intersectionRatio).toFixed(2)}px`
              );
            } else {
              document.documentElement.style.setProperty("--ann-height", "0px");
            }
          });
        },
        {
          threshold: Array.from({ length: 1000 }, (_, i) => i / 1000),
        }
      );

      if (annBar) annBarObserver.observe(annBar);
    };
    createAnnBarObserver();

    menuLinks.forEach((link) => {
      link.addEventListener("mouseenter", (e) => {
        if (link.classList.contains("list-menu--megamenu")) {
          menuLinks.forEach((el) => {
            el.classList.add("list-menu-item--inactive");
            el.classList.remove("list-menu-item--active");
          });
          link.classList.remove("list-menu-item--inactive");
          link.classList.add("list-menu-item--active");
        } else {
          menuLinks.forEach((el) => {
            el.classList.add("list-menu-item--inactive");
            el.classList.remove("list-menu-item--active");
          });
          link.classList.remove("list-menu-item--inactive");
          link.classList.add("list-menu-item--active");
        }
      });
    });

    menuLinksWithDropdown.forEach((item) => {
      item.addEventListener("mouseenter", () => {
        if (headerWrapper) headerWrapper.preventHide = true;
      });

      item.addEventListener("mouseleave", () => {
        if (headerWrapper) headerWrapper.preventHide = false;
      });
    });

    menu?.addEventListener("mouseleave", () => {
      // Find the actual active page item (the one with header__active-menu-item class)
      const activePageItem = Array.from(menuLinks).find((link) =>
        link.querySelector(".header__active-menu-item")
      );

      // Remove all active/inactive classes
      menuLinks.forEach((el) => {
        el.classList.remove("list-menu-item--active");
        el.classList.remove("list-menu-item--inactive");
      });

      if (activePageItem) {
        activePageItem.classList.add("list-menu-item--active");
      }
    });

    const submenuOffset = 48;
    const flyoutCloseDelay = 400;

    const setSubmenuMaxHeight = (submenu) => {
      const available =
        window.innerHeight -
        submenu.getBoundingClientRect().top -
        submenuOffset;
      submenu.style.setProperty(
        "--submenu-max-height",
        `${Math.max(available, 160)}px`
      );
    };

    const positionFlyout = (flyout) => {
      const parentList = flyout.parentElement.closest(".header__submenu");

      flyout.classList.remove("header__submenu--flyout");
      flyout.style.removeProperty("top");
      flyout.style.removeProperty("left");
      flyout.style.removeProperty("--submenu-max-height");

      const rect = flyout.getBoundingClientRect();
      const minTop = parentList.getBoundingClientRect().top;
      const maxTop = window.innerHeight - submenuOffset - flyout.scrollHeight;
      const top = Math.max(Math.min(rect.top, maxTop), minTop);

      flyout.classList.add("header__submenu--flyout");
      flyout.style.setProperty("top", `${top}px`, "important");
      flyout.style.setProperty("left", `${rect.left}px`, "important");

      const placed = flyout.getBoundingClientRect();
      flyout.style.setProperty("top", `${2 * top - placed.top}px`, "important");
      flyout.style.setProperty(
        "left",
        `${2 * rect.left - placed.left}px`,
        "important"
      );

      setSubmenuMaxHeight(flyout);
    };

    allSubmenu.forEach((submenu) => {
      if (submenu.closest(".simple_mega_menu")) return;

      const trigger = submenu.parentElement;
      const isFlyout = trigger.closest(".header__submenu");
      const update = () =>
        isFlyout ? positionFlyout(submenu) : setSubmenuMaxHeight(submenu);

      trigger.addEventListener("mouseenter", update);
      trigger.addEventListener("focusin", (e) => {
        if (!submenu.contains(e.target)) update();
      });

      if (!isFlyout) return;

      isFlyout.addEventListener("scroll", () => {
        if (
          trigger.matches(":hover, :focus-within") ||
          submenu.classList.contains("header__submenu--open")
        ) {
          positionFlyout(submenu);
        }
      });

      let closeTimer;
      const close = () => {
        clearTimeout(closeTimer);
        submenu.classList.remove("header__submenu--open");
      };

      trigger.addEventListener("mouseenter", () => {
        clearTimeout(closeTimer);
        isFlyout
          .querySelectorAll(":scope > li > .header__submenu--open")
          .forEach(
            (el) =>
              el !== submenu && el.dispatchEvent(new Event("flyout:close"))
          );
      });
      trigger.addEventListener("mouseleave", () => {
        submenu.classList.add("header__submenu--open");
        clearTimeout(closeTimer);
        closeTimer = setTimeout(close, flyoutCloseDelay);
      });
      submenu.addEventListener("flyout:close", close);
    });

    allSubmenu.forEach((submenu) => {
      const links = submenu.querySelectorAll(
        ".header__submenu-item:not(.header__submenu-item--grandchild)"
      );

      links.forEach((link) => {
        const childLinks = link.parentElement.querySelectorAll(
          ".header__submenu-item--grandchild"
        );

        link.addEventListener("mouseenter", (e) => {
          e.stopPropagation();

          links.forEach((el) => {
            el.classList.add("header__submenu-item--inactive");
          });
          link.classList.remove("header__submenu-item--inactive");
        });
      });

      submenu.addEventListener("mouseleave", (e) => {
        e.stopPropagation();

        links.forEach((el) => {
          el.classList.remove("header__submenu-item--inactive");
        });
      });
    });

    //if (header && header.classList.contains("color-background-overlay")) {
    //	header.addEventListener("mouseenter", () => {
    //		header.classList.remove("color-background-overlay");
    //		header.classList.add("color-background-overlay-hidden");
    //	});

    //	header.addEventListener("mouseleave", () => {
    //		if (!searchModal?.hasAttribute("open")) {
    //			header.classList.add("color-background-overlay");
    //			header.classList.remove("color-background-overlay-hidden");
    //		}
    //	});
    //}
    searchOpenBtn.forEach((btn) => {
      const openSearch = () => {
        searchInner.forEach((item) => {
          item.classList.add("header__search-inner--open");
        });
        document.body.classList.add(`overflow-hidden-search`);
        trapFocus(document.querySelector(".search__content"));
      };

      btn?.addEventListener("click", openSearch);

      btn?.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          openSearch();
        }
      });
    });

    searchOpenBtnFooter.addEventListener("click", () => {
      window.scrollBy({
        top: -1,
        behavior: "smooth",
      });
    });

    searchInner.forEach((item) => {
      item.addEventListener("click", (e) => {
        if (e.target === item) {
          item.classList.remove("header__search-inner--open");
          document.body.classList.remove(`overflow-hidden-search`);
        }
      });
    });

    searchPromoBg?.addEventListener("click", (e) => {
      if (e.target.classList.contains("search__modal")) {
        searchInner.forEach((item) => {
          item.classList.remove("header__search-inner--open");
          document.body.classList.remove(`overflow-hidden-search`);
        });
      }
    });
    searchPredictiveBg.forEach((item) => {
      item.addEventListener("click", (e) => {
        if (e.target.classList.contains("search-modal__form")) {
          searchInner.forEach((item) => {
            item.classList.remove("header__search-inner--open");
            document.body.classList.remove(`overflow-hidden-search`);
          });
        }
      });
    });

    if (searchModal?.hasAttribute("open")) {
      header.classList.add("color-background-overlay-hidden");
    } else {
      header.classList.remove("color-background-overlay-hidden");
    }

    searchModal?.addEventListener("toggle", () => {
      if (searchModal.hasAttribute("open")) {
        header.classList.add("color-background-overlay-hidden");
      } else {
        header.classList.remove("color-background-overlay-hidden");
      }
    });
  };

  //function initAsideMenuToggle() {
  //	var asideToggles = document.querySelectorAll('.js-aside-menu-toggle')
  //	var asideMenu = document.querySelector('.aside-menu')

  //	if (asideToggles.length && asideMenu) {
  //		asideToggles.forEach(function (asideToggle) {
  //			if (asideToggle.dataset.asideMenuBound) return
  //			asideToggle.addEventListener('click', asideMenuClickHandler)
  //			asideToggle.dataset.asideMenuBound = 'true'
  //		})

  //		function asideMenuClickHandler(e) {
  //			e.stopPropagation()
  //			var asideToggle = e.currentTarget
  //			var isOpen = asideToggle.classList.toggle('menu-opening')
  //			asideMenu.classList.toggle('open', isOpen)

  //			asideToggles.forEach((btn) => {
  //				if (btn !== asideToggle) btn.classList.remove('menu-opening')
  //			})

  //			if (isOpen) {
  //				document.body.addEventListener('click', closeAsideMenu, {
  //					once: true,
  //				})
  //			}
  //		}

  //		function closeAsideMenu(e) {
  //			asideToggles.forEach((btn) => {
  //				btn.classList.remove('menu-opening')
  //			})
  //			asideMenu.classList.remove('open')
  //		}
  //	}
  //}

  document.addEventListener("shopify:section:load", function () {
    header();
    //initAsideMenuToggle()
  });
  document.addEventListener("shopify:section:unload", header);
  document.addEventListener("shopify:section:reorder", header);

  header();
  //initAsideMenuToggle()
})();
