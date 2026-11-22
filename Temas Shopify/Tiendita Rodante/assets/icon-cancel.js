(function () {
  const initCancelButtons = () => {
    const cancelIcons = document.querySelectorAll(".reset__button");

    cancelIcons.forEach((cancelIcon) => {
      const searchForm = cancelIcon.closest(".search");
      if (!searchForm) return;

      const searchInput = searchForm.querySelector(".search__input");
      if (!searchInput) return;

      cancelIcon.style.display = searchInput.value ? "flex" : "none";

      cancelIcon.onclick = function () {
        searchInput.value = "";
        searchInput.focus();
        this.style.display = "none";
      };

      searchInput.oninput = function () {
        cancelIcon.style.display = this.value ? "flex" : "none";
      };
    });
  };

  document.addEventListener("DOMContentLoaded", initCancelButtons);

  document.addEventListener("shopify:section:load", function (event) {
    initCancelButtons();
  });
})();
