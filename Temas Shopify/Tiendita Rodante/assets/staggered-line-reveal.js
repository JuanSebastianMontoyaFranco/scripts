(function () {
  function initStaggeredLineReveal(element) {
    const words = element.innerText.trim().split(" ");
    element.innerHTML = words
      .map((word) => `<span class="staggered-line-reveal__word">${word}</span>`)
      .join(" ");

    const spans = Array.from(
      element.querySelectorAll(".staggered-line-reveal__word")
    );

    const lines = [];
    let currentLineTop = null;
    let currentLine = [];

    spans.forEach((span) => {
      const top = span.offsetTop;
      if (currentLineTop === null) currentLineTop = top;

      if (top === currentLineTop) {
        currentLine.push(span);
      } else {
        lines.push(currentLine);
        currentLine = [span];
        currentLineTop = top;
      }
    });
    if (currentLine.length) lines.push(currentLine);
    lines.forEach((line, i) => {
      setTimeout(() => {
        line.forEach((span) => span.classList.add("show"));
      }, i * 200);
    });
  }

  const observer = new IntersectionObserver(
    (entries, obs) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          initStaggeredLineReveal(entry.target);
          obs.unobserve(entry.target);
        }
      });
    },
    { threshold: 0.6 }
  );

  document.querySelectorAll(".staggered-line-reveal").forEach((el) => {
    observer.observe(el);
  });
})();
