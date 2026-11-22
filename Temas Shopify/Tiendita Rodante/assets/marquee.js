(function () {
	const marqueeContent = document.getElementById("marquee-content");

	for (let i = 0; i <= 4; i++) {
		marqueeContent.innerHTML += marqueeContent.innerHTML;
	}

	window.addEventListener("load", () => {
		const contentWidth = marqueeContent.offsetWidth / 2;
		document.documentElement.style.setProperty(
			"--marquee-distance",
			contentWidth + "px"
		);
	});

	const marqueeContainer = document.getElementById("marquee-content");
	let lastScrollTop = window.pageYOffset || document.documentElement.scrollTop;

	window.addEventListener("scroll", () => {
		const scrollTop = window.pageYOffset || document.documentElement.scrollTop;
		if (scrollTop > lastScrollTop) {
			marqueeContainer.style.animationDirection = "reverse";
		} else {
			marqueeContainer.style.animationDirection = "normal";
		}
		lastScrollTop = scrollTop <= 0 ? 0 : scrollTop;
	});
})();
