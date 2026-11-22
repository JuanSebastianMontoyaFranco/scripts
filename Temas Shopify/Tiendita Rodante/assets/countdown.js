if (!customElements.get("countdown-timer")) {
  class Countdown extends HTMLElement {
    constructor() {
      super();

      this.enableMockTimer =
        this.getAttribute("data-enable-mock-timer") === "true";

      if (!this.enableMockTimer) {
        this.userDate = this.getAttribute("data-date");
        this.userTime = this.getAttribute("data-time");
      } else {
        // get data for infinite mock timer
        const INFINITE_DAYS = this.getAttribute("data-enable-mock-picker-day")
          ? Number(this.getAttribute("data-enable-mock-picker-day"))
          : 2;
        const DIFF_HOURS = this.getAttribute("data-enable-mock-picker-hours")
          ? Number(this.getAttribute("data-enable-mock-picker-hours"))
          : 3;
        const now = new Date();
        const futureDate = new Date(
          now.getTime() + (INFINITE_DAYS * 24 + DIFF_HOURS) * 60 * 60 * 1000
        );
        const year = futureDate.getFullYear();
        const month = String(futureDate.getMonth() + 1).padStart(2, "0");
        const day = String(futureDate.getDate()).padStart(2, "0");
        const hours = String(futureDate.getHours()).padStart(2, "0");
        const minutes = String(futureDate.getMinutes()).padStart(2, "0");
        this.userDate = `${year}-${month}-${day}`;
        this.userTime = `${hours}:${minutes}`;
      }

      this.completedAction = this.getAttribute("data-completed");
      this.completedText = this.getAttribute("data-completed-text");
      this.interval = null;
      this.initCountdown();
    }

    initCountdown() {
      this.countdownBlocks = this.querySelector(".countdown__main");
      this.daysEl = this.querySelector(".countdown_block_days");
      this.hoursEl = this.querySelector(".countdown_block_hours");
      this.minutesEl = this.querySelector(".countdown_block_minutes");
      this.secondsEl = this.querySelector(".countdown_block_seconds");
      this.countDownMain = this.querySelector(".countdown__main");
      this.section = this.closest(".countdown-section");

      this.completedTextEl = document.createElement("div");
      this.completedTextEl.className = "countdown-completed-text";
      this.completedTextEl.style.display = "none";
      this.countdownBlocks?.parentNode.insertBefore(
        this.completedTextEl,
        this.countdownBlocks.nextSibling
      );

      this.updateCountdown();
      this.interval = setInterval(this.updateCountdown.bind(this), 1000);
    }

    updateCountdown() {
      const countdownDate = new Date(`${this.userDate}T${this.userTime}`);
      const now = new Date();
      const distance = countdownDate.getTime() - now.getTime();

      if (distance <= 0) {
        clearInterval(this.interval);

        if (this.completedAction === "hide_section") {
          this.section?.style.setProperty("display", "none", "important");
          this.countDownMain?.style.setProperty("display", "none", "important");

          const notificationCountdown = document.querySelector(
            ".notification-banner__countdown"
          );

          if (notificationCountdown) {
            const countdownHeadingNotification =
              notificationCountdown.dataset.heading;
            if (!countdownHeadingNotification.length) {
              notificationCountdown.style.setProperty(
                "display",
                "none",
                "important"
              );
            }
          }
        } else if (this.completedAction === "show_text") {
          this.countdownBlocks.style.display = "none";
          this.completedTextEl.style.display = "block";
          this.completedTextEl.textContent =
            this.completedText || "Timer ends!";
        } else {
          this.daysEl.textContent = "00";
          this.hoursEl.textContent = "00";
          this.minutesEl.textContent = "00";
          this.secondsEl.textContent = "00";
        }
        return;
      }

      const days = Math.floor(distance / (1000 * 60 * 60 * 24));
      const hours = Math.floor(
        (distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60)
      );
      const minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
      const seconds = Math.floor((distance % (1000 * 60)) / 1000);

      this.daysEl.textContent = String(days).padStart(2, "0");
      this.hoursEl.textContent = String(hours).padStart(2, "0");
      this.minutesEl.textContent = String(minutes).padStart(2, "0");
      this.secondsEl.textContent = String(seconds).padStart(2, "0");
    }

    disconnectedCallback() {
      clearInterval(this.interval);
    }
  }

  customElements.define("countdown-timer", Countdown);
}
