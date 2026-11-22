(function () {
	const marquee = document.querySelector(".marquee__inner");
	let position = window.innerWidth;

	function animateMarquee() {
		position -= 2;
		if (position < -marquee.clientWidth) {
			position = window.innerWidth;
		}
		marquee.style.transform = `translateX(${position}px)`;
		requestAnimationFrame(animateMarquee);
	}

	animateMarquee();
})();
