document.addEventListener("DOMContentLoaded", () => {
	const cancelIcons = document.querySelectorAll(".icon-cancel");

	cancelIcons.forEach((icon) => {
		icon.addEventListener("click", function () {

			const searchForm = this.closest("form");
			if (searchForm) {
				const searchInput = searchForm.querySelector(".search__input");
				if (searchInput) {
					searchInput.value = "";
					searchInput.focus();

					searchInput.dispatchEvent(new Event("input", { bubbles: true }));
				}
			}
		});
	});
});
