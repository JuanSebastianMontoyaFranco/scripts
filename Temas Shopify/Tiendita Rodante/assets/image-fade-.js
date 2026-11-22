(function () {
	const initScaleAnimation = (section) => {
		const mediaEl = section.querySelector(".animate-on-scroll");

		if (!mediaEl) return;

		const observerOptions = {
			threshold: 0.1,
			rootMargin: "0px 0px -10% 0px",
		};

		const observer = new IntersectionObserver((entries, obs) => {
			entries.forEach((entry) => {
				if (entry.isIntersecting) {
					mediaEl.classList.add("animated");
					obs.unobserve(entry.target);
				}
			});
		}, observerOptions);

		observer.observe(mediaEl);
	};

	initScaleAnimation(document.currentScript.parentElement);

	document.addEventListener("shopify:section:load", function (section) {
		initScaleAnimation(section.target);
	});
})();
