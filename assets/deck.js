/* AI Agents: deck engine
   Navigation, per-slide entrance choreography, drawn wires, lightbox,
   speaker notes and a synced presenter window. No dependencies. */
(() => {
  "use strict";

  const W = 1920, H = 1080, OUT_MS = 420;
  const SVG_NS = "http://www.w3.org/2000/svg";

  const stage = document.getElementById("stage");
  const slides = [...stage.querySelectorAll(".slide")];
  const chrome = stage.querySelector(".chrome");
  const curEl = chrome.querySelector(".cur");
  const sectionEl = chrome.querySelector(".section-name");
  const progress = stage.querySelector(".progress");
  const notes = document.getElementById("notes");
  const help = document.getElementById("help");
  const lightbox = document.getElementById("lightbox");
  const pad = (n) => String(n).padStart(2, "0");

  const params = new URLSearchParams(location.search);
  const isPresenter = params.has("presenter");
  const channel = "BroadcastChannel" in window ? new BroadcastChannel("ai-agents-deck") : null;

  let index = -1;
  let leaveTimer = null;

  /* ---------------------------------------------------------------- icons */
  document.querySelectorAll("i[data-icon]").forEach((el) => {
    const svg = document.createElementNS(SVG_NS, "svg");
    svg.setAttribute("viewBox", "0 0 24 24");
    svg.setAttribute("aria-hidden", "true");
    svg.setAttribute("class", ["icon", el.className].filter(Boolean).join(" "));
    svg.innerHTML = (window.ICONS || {})[el.dataset.icon] || "";
    svg.querySelectorAll("path, circle, rect, line, polyline, polygon, ellipse")
      .forEach((p) => p.setAttribute("pathLength", "1"));
    el.replaceWith(svg);
  });

  /* ------------------------------------------- entrance order and delays */
  slides.forEach((slide) => {
    let i = 0;
    slide.querySelectorAll("[data-a]").forEach((el) => {
      if (el.dataset.d != null) el.style.setProperty("--d", el.dataset.d + "ms");
      else el.style.setProperty("--i", i++);
    });
    slide.querySelector(".speaker-notes")?.setAttribute("aria-hidden", "true");
  });
  document.getElementById("stage").querySelector(".tot").textContent = pad(slides.length);

  /* ----------------------------------------------------------- scaling */
  function fit() {
    const presenterW = isPresenter ? Math.min(window.innerWidth * 0.62, window.innerWidth - 380) : window.innerWidth;
    const s = Math.min(presenterW / W, window.innerHeight / H);
    stage.style.transform = `scale(${s})`;
    stage.dataset.scale = s;
  }
  window.addEventListener("resize", () => { fit(); drawWires(); });

  /* ------------------------------------------------------------- wires */
  // Elbow connectors drawn between elements; they trace themselves in on entrance.
  function rel(el, box) {
    const s = parseFloat(stage.dataset.scale) || 1;
    const r = el.getBoundingClientRect();
    return {
      l: (r.left - box.left) / s, t: (r.top - box.top) / s,
      r: (r.right - box.left) / s, b: (r.bottom - box.top) / s,
      cx: (r.left + r.width / 2 - box.left) / s, cy: (r.top + r.height / 2 - box.top) / s,
    };
  }
  function addPath(svg, d, delay) {
    const p = document.createElementNS(SVG_NS, "path");
    p.setAttribute("d", d);
    p.setAttribute("pathLength", "1");
    p.setAttribute("class", "draw-path");
    p.style.setProperty("--delay", delay + "ms");
    svg.appendChild(p);
  }
  function drawWires() {
    const org = document.getElementById("org");
    if (org) {
      const svg = org.querySelector("svg.wires");
      svg.innerHTML = "";
      const box = org.getBoundingClientRect();
      const me = rel(org.querySelector('[data-wire="me"]'), box);
      const ceo = rel(org.querySelector('[data-wire="ceo"]'), box);
      addPath(svg, `M ${me.cx} ${me.b} V ${ceo.t}`, 250);
      const icons = [...org.querySelectorAll('[data-wire="crew"] .icon-wrap')].map((c) => rel(c, box));
      const busY2 = (ceo.b + icons[0].t) / 2;
      icons.forEach((c, k) => {
        const r = 18, dir = Math.sign(c.cx - ceo.cx);
        const d = dir === 0
          ? `M ${ceo.cx} ${ceo.b} V ${c.t}`
          : `M ${ceo.cx} ${ceo.b} V ${busY2 - r} Q ${ceo.cx} ${busY2} ${ceo.cx + dir * r} ${busY2} H ${c.cx - dir * r} Q ${c.cx} ${busY2} ${c.cx} ${busY2 + r} V ${c.t}`;
        addPath(svg, d, 650 + Math.abs(k - 3) * 90);
      });
    }
    const hr = document.getElementById("hrmap");
    if (hr) {
      const svg = hr.querySelector("svg.wires");
      svg.innerHTML = "";
      const box = hr.getBoundingClientRect();
      const team = rel(hr.querySelector('[data-wire="hrteam"]'), box);
      const head = rel(hr.querySelector('[data-wire="hrhead"]'), box);
      addPath(svg, `M ${team.l + 40} ${team.b} V ${head.t}`, 350);
      const icons = [...hr.querySelectorAll('[data-wire="spec"] .icon-wrap')].map((c) => rel(c, box));
      const busX = head.r + 60, r = 16;
      icons.forEach((c, k) => {
        const dy = c.cy - head.cy, dir = Math.sign(dy);
        const d = Math.abs(dy) < r * 2
          ? `M ${head.r} ${head.cy} H ${c.l}`
          : `M ${head.r} ${head.cy} H ${busX - r} Q ${busX} ${head.cy} ${busX} ${head.cy + dir * r} V ${c.cy - dir * r} Q ${busX} ${c.cy} ${busX + r} ${c.cy} H ${c.l}`;
        addPath(svg, d, 700 + k * 70);
      });
    }
  }

  /* -------------------------------------------------------- navigation */
  function go(n, opts = {}) {
    n = Math.max(0, Math.min(slides.length - 1, n));
    if (n === index && !opts.force) return;
    const back = n < index;
    const prev = slides[index];
    const next = slides[n];

    clearTimeout(leaveTimer);
    slides.forEach((s) => s.classList.remove("is-leaving", "back"));
    if (prev && prev !== next) {
      prev.classList.remove("is-active", "is-in");
      prev.classList.add("is-leaving");
      prev.classList.toggle("back", back);
      leaveTimer = setTimeout(() => prev.classList.remove("is-leaving", "back"), OUT_MS);
    }

    index = n;
    next.classList.remove("is-in");
    next.classList.add("is-active");
    stage.classList.toggle("is-dark", next.classList.contains("dark"));
    stage.classList.toggle("no-chrome", n === 0 || n === slides.length - 1);

    // Let the outgoing slide clear before the incoming one starts its entrance.
    const startIn = () => { void next.offsetWidth; next.classList.add("is-in"); };
    if (prev && prev !== next && !opts.instant) setTimeout(startIn, back ? 120 : 220);
    else startIn();

    curEl.textContent = pad(n + 1);
    sectionEl.textContent = next.dataset.section || "";
    progress.style.transform = `scaleX(${(n + 1) / slides.length})`;
    document.title = `${pad(n + 1)} · ${next.getAttribute("aria-label")}: AI Agents`;
    if (location.hash !== "#" + (n + 1)) history.replaceState(null, "", "#" + (n + 1));
    renderNotes();
    if (!opts.remote && channel) channel.postMessage({ type: "go", index: n });
  }
  const nextSlide = () => go(index + 1);
  const prevSlide = () => go(index - 1);

  /* ----------------------------------------------------------- notes */
  function noteOf(i) { return slides[i]?.querySelector(".speaker-notes")?.textContent.trim() || ""; }
  function renderNotes() {
    notes.querySelector(".meta").textContent =
      `${pad(index + 1)} / ${pad(slides.length)} · ${slides[index].getAttribute("aria-label")}` +
      (slides[index + 1] ? `   →   Next: ${slides[index + 1].getAttribute("aria-label")}` : "");
    notes.querySelector(".text").textContent = noteOf(index);
  }

  /* ------------------------------------------------------- presenter */
  if (isPresenter) {
    document.body.classList.add("presenter");
    notes.classList.add("open");
    const timer = document.createElement("div");
    timer.className = "timer";
    notes.prepend(timer);
    const t0 = Date.now();
    setInterval(() => {
      const s = Math.floor((Date.now() - t0) / 1000);
      timer.textContent = `${pad(Math.floor(s / 60))}:${pad(s % 60)}`;
    }, 500);
  }
  channel?.addEventListener("message", (e) => {
    if (e.data?.type === "go") go(e.data.index, { remote: true });
  });

  /* -------------------------------------------------------- lightbox */
  let lbSource = null;
  function openLightbox(img) {
    lbSource = img;
    const big = new Image();
    big.src = img.currentSrc || img.src;
    big.alt = img.alt;
    lightbox.innerHTML = "";
    lightbox.appendChild(big);
    lightbox.setAttribute("aria-hidden", "false");
    const go = () => {
      const from = img.getBoundingClientRect();
      const to = big.getBoundingClientRect();
      const sx = from.width / to.width, sy = from.height / to.height;
      big.animate(
        [
          { transform: `translate(${from.left - to.left}px, ${from.top - to.top}px) scale(${sx}, ${sy})` },
          { transform: "none" },
        ],
        { duration: 560, easing: "cubic-bezier(0.16, 1, 0.3, 1)" }
      );
      img.style.visibility = "hidden";
      lightbox.classList.add("open");
    };
    big.decode ? big.decode().then(go, go) : go();
  }
  function closeLightbox() {
    if (!lightbox.classList.contains("open")) return;
    const big = lightbox.querySelector("img");
    const from = big.getBoundingClientRect();
    const to = lbSource.getBoundingClientRect();
    const anim = big.animate(
      [
        { transform: "none" },
        { transform: `translate(${to.left - from.left}px, ${to.top - from.top}px) scale(${to.width / from.width}, ${to.height / from.height})` },
      ],
      { duration: 420, easing: "cubic-bezier(0.16, 1, 0.3, 1)", fill: "forwards" }
    );
    lightbox.classList.remove("open");
    anim.onfinish = () => {
      lbSource.style.visibility = "";
      lightbox.innerHTML = "";
      lightbox.setAttribute("aria-hidden", "true");
    };
  }
  lightbox.addEventListener("click", closeLightbox);

  /* ---------------------------------------------------------- input */
  function toggleFullscreen() {
    if (!document.fullscreenElement) document.documentElement.requestFullscreen?.();
    else document.exitFullscreen?.();
  }

  document.addEventListener("keydown", (e) => {
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    const k = e.key;
    if (lightbox.classList.contains("open")) {
      if (k === "Escape" || k === " " || k === "Enter") { e.preventDefault(); closeLightbox(); }
      return;
    }
    if (["ArrowRight", "ArrowDown", "PageDown", " ", "Enter"].includes(k)) { e.preventDefault(); nextSlide(); }
    else if (["ArrowLeft", "ArrowUp", "PageUp", "Backspace"].includes(k)) { e.preventDefault(); prevSlide(); }
    else if (k === "Home") go(0);
    else if (k === "End") go(slides.length - 1);
    else if (k === "f" || k === "F") toggleFullscreen();
    else if (k === "n" || k === "N") notes.classList.toggle("open");
    else if (k === "p" || k === "P") window.open(location.pathname + "?presenter#" + (index + 1), "ai-agents-presenter", "width=1280,height=760");
    else if (k === "?") help.classList.toggle("open");
    else if (k === "Escape") { help.classList.remove("open"); if (!isPresenter) notes.classList.remove("open"); }
    else if (/^[1-9]$/.test(k)) go(parseInt(k, 10) - 1);
  });

  stage.addEventListener("click", (e) => {
    const shot = e.target.closest(".shot");
    if (shot) { openLightbox(shot); return; }
    if (window.getSelection()?.toString()) return;
    const s = parseFloat(stage.dataset.scale) || 1;
    const x = (e.clientX - stage.getBoundingClientRect().left) / s;
    x < W * 0.25 ? prevSlide() : nextSlide();
  });

  let touch = null;
  stage.addEventListener("touchstart", (e) => { touch = e.touches[0]; }, { passive: true });
  stage.addEventListener("touchend", (e) => {
    if (!touch) return;
    const t = e.changedTouches[0];
    const dx = t.clientX - touch.clientX, dy = t.clientY - touch.clientY;
    if (Math.abs(dx) > 50 && Math.abs(dx) > Math.abs(dy)) { e.preventDefault(); dx < 0 ? nextSlide() : prevSlide(); }
    touch = null;
  });

  window.addEventListener("hashchange", () => {
    const n = parseInt(location.hash.slice(1), 10) - 1;
    if (!Number.isNaN(n) && n !== index) go(n);
  });

  /* ----------------------------------------------------------- boot */
  fit();
  const start = Math.max(0, (parseInt(location.hash.slice(1), 10) || 1) - 1);
  const boot = () => { drawWires(); go(start, { instant: true, force: true, remote: true }); };
  (document.fonts?.ready || Promise.resolve()).then(boot);
})();
