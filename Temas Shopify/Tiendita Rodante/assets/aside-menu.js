//(function () {
//	const initAsideMenu = () => {
//		const header = document.querySelector(".header-wrapper");
//		const asideMenu = document.querySelector(".aside-menu");
//		const openAsideMenuBtn = document.querySelector(
//			".aside-menu__toggle--open-btn"
//		);
//		const asideMenuOverlay = document.querySelector(".aside-menu__overlay");
//		if (header) header.preventHide = false;

//		const openAsideMenu = (e) => {
//			e.preventDefault();
//			if (header) header.preventHide = true;
//			asideMenu.classList.add("aside-menu--open");
//			asideMenuOverlay.classList.add("aside-menu__overlay--active");
//			document.body.classList.add("overflow-hidden-drawer");

//			trapFocus(document.querySelector(".aside-menu"));
//		};

//		const closeAsideMenu = (e) => {
//			e.preventDefault();
//			asideMenu.classList.remove("aside-menu--open");
//			asideMenuOverlay.classList.remove("aside-menu__overlay--active");
//			document.body.classList.remove("overflow-hidden-drawer");
//			if (header) header.preventHide = false;
//		};

//		openAsideMenuBtn.addEventListener("click", openAsideMenu);
//		asideMenuOverlay.addEventListener("click", closeAsideMenu);

//		asideMenu.addEventListener("keydown", (e) => {
//			if (e.code === "Escape") {
//				closeAsideMenu(e);

//				openAsideMenuBtn.classList.toggle("menu-opening");
//			}
//		});
//	};

//	document.addEventListener("shopify:section:load", initAsideMenu);
//	document.addEventListener("shopify:section:unload", initAsideMenu);
//	document.addEventListener("shopify:section:reorder", initAsideMenu);

//	initAsideMenu();
//})();
