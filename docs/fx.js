// Qent site effects: network background, Lenis smooth scroll, GSAP / ScrollTrigger reveals, diagram draw-in.
// Everything is additive: with JS off, or reduced motion, the page shows all content immediately.
(() => {
const $ = (s, r = document) => r.querySelector(s), all = (s, r = document) => [...r.querySelectorAll(s)];
const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches, G = window.gsap, ST = window.ScrollTrigger;
document.documentElement.classList.add("fx");
if (G && ST) G.registerPlugin(ST);

// network background
(() => {
  const c = $("#net"); if (!c) return;
  const x = c.getContext("2d"); let W = 0, H = 0, dpr = 1, pts = [], mx = -1e4, my = -1e4, tick = 0;
  const size = () => {
    dpr = Math.min(devicePixelRatio || 1, 2); W = c.width = innerWidth * dpr; H = c.height = innerHeight * dpr;
    const n = Math.round(Math.min(72, innerWidth * innerHeight / 24000));
    while (pts.length < n) pts.push({x: Math.random() * W, y: Math.random() * H, vx: (Math.random() - .5) * .3 * dpr, vy: (Math.random() - .5) * .3 * dpr, r: (1 + Math.random() * 1.5) * dpr, hot: Math.random() < .07, ph: Math.random() * 6.3});
    pts.length = n;
  };
  const draw = () => {
    x.clearRect(0, 0, W, H); const D = 165 * dpr, M = 190 * dpr; tick += .02;
    if (!reduce) for (const p of pts) { p.x += p.vx; p.y += p.vy; if (p.x < 0 || p.x > W) p.vx *= -1; if (p.y < 0 || p.y > H) p.vy *= -1; }
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i];
      for (let j = i + 1; j < pts.length; j++) {
        const b = pts[j], d = Math.hypot(a.x - b.x, a.y - b.y);
        if (d < D) { x.strokeStyle = `rgba(62,230,168,${(1 - d / D) * .3})`; x.lineWidth = dpr; x.beginPath(); x.moveTo(a.x, a.y); x.lineTo(b.x, b.y); x.stroke(); }
      }
      const dm = Math.hypot(a.x - mx, a.y - my);
      if (dm < M) { x.strokeStyle = `rgba(90,169,255,${(1 - dm / M) * .5})`; x.lineWidth = dpr; x.beginPath(); x.moveTo(a.x, a.y); x.lineTo(mx, my); x.stroke(); }
    }
    for (const p of pts) {
      x.fillStyle = p.hot ? "rgba(255,107,120,.9)" : "rgba(62,230,168,.8)"; x.beginPath(); x.arc(p.x, p.y, p.r, 0, 6.3); x.fill();
      if (p.hot) { const k = (Math.sin(tick + p.ph) + 1) / 2; x.strokeStyle = `rgba(255,107,120,${.35 * (1 - k)})`; x.beginPath(); x.arc(p.x, p.y, p.r + k * 16 * dpr, 0, 6.3); x.stroke(); }
    }
  };
  const loop = () => { if (!document.hidden) draw(); requestAnimationFrame(loop); };
  size(); addEventListener("resize", () => { size(); if (reduce) draw(); });
  addEventListener("pointermove", e => { mx = e.clientX * dpr; my = e.clientY * dpr; }, {passive: true});
  reduce ? draw() : loop();
})();

// Lenis smooth scroll on GSAP's ticker; anchor links glide
let lenis = null;
all(".scroll").forEach(e => e.setAttribute("data-lenis-prevent", ""));
if (!reduce && window.Lenis) {
  lenis = new Lenis({lerp: .09, smoothWheel: true});
  if (G && ST) { lenis.on("scroll", ST.update); G.ticker.add(t => lenis.raf(t * 1000)); G.ticker.lagSmoothing(0); }
  else { const raf = t => { lenis.raf(t); requestAnimationFrame(raf); }; requestAnimationFrame(raf); }
  all('a[href^="#"]').forEach(a => a.addEventListener("click", e => { const t = $(a.getAttribute("href")); if (t) { e.preventDefault(); lenis.scrollTo(t, {offset: -70}); } }));
}

const prog = $("#prog");
if (prog && ST) ST.create({start: 0, end: "max", onUpdate: s => prog.style.transform = `scaleX(${s.progress})`});

function split(el) {
  const walk = n => [...n.childNodes].forEach(c => {
    if (c.nodeType === 3) {
      const f = document.createDocumentFragment();
      c.textContent.split(/(\s+)/).forEach(s => { if (!s.trim()) return f.append(s); const w = document.createElement("span"), i = document.createElement("span"); w.className = "w"; i.className = "wi"; i.textContent = s; w.append(i); f.append(w); });
      c.replaceWith(f);
    } else if (c.nodeType === 1) walk(c);
  });
  walk(el); return all(".wi", el);
}
function countUp(el, d = 1.3) {
  const m = el.textContent.match(/\d[\d,]*/); if (!m || !G) return;
  const to = +m[0].replace(/,/g, ""), comma = m[0].includes(","), pre = el.textContent.slice(0, m.index), post = el.textContent.slice(m.index + m[0].length), o = {v: 0};
  if (!to) return;
  el.textContent = pre + "0" + post;
  G.to(o, {v: to, duration: d, ease: "power2.out", onUpdate: () => { const v = Math.round(o.v); el.textContent = pre + (comma ? v.toLocaleString("en") : v) + post; }});
}

if (G && !reduce) {
  // hero entrance
  const h = $("h1");
  if (h) {
    const words = split(h);
    const tl = G.timeline({defaults: {ease: "power3.out"}}), has = s => $(s) ? s : null;
    if (has(".hero .tag")) tl.from(".hero .tag", {y: 12, opacity: 0, duration: .5});
    tl.from(words, {yPercent: 115, duration: .85, stagger: .05}, "-=.2");
    if (has(".hero .lead")) tl.from(".hero .lead", {y: 16, opacity: 0, duration: .6, clearProps: "transform,opacity"}, "-=.5");
    if (has(".hero .btn")) tl.from(".hero .btn", {y: 16, opacity: 0, duration: .5, stagger: .1, clearProps: "transform,opacity"}, "-=.4");
  }
  // scroll reveals
  if (ST) {
    all("section .wrap > *, header .wrap > p, header .wrap > nav").forEach(el => {
      if (el.getBoundingClientRect().top < innerHeight * .9) return;
      G.set(el, {y: 34, opacity: 0});
      ST.create({trigger: el, start: "top 90%", once: true, onEnter: () => G.to(el, {y: 0, opacity: 1, duration: .8, ease: "power3.out", clearProps: "transform,opacity"})});
    });
    all(".grid").forEach(g => {
      const cards = all(".card", g); if (!cards.length) return;
      G.set(cards, {y: 28, opacity: 0});
      ST.create({trigger: g, start: "top 88%", once: true, onEnter: () => { G.to(cards, {y: 0, opacity: 1, duration: .7, stagger: .1, ease: "power3.out", clearProps: "transform,opacity"}); all(".num", g).forEach(n => countUp(n)); }});
    });
    // nav: highlight the section in view
    all("nav a[href^='#']").forEach(a => { const s = $(a.getAttribute("href")); if (s) ST.create({trigger: s, start: "top 55%", end: "bottom 55%", toggleClass: {targets: a, className: "on"}}); });
    // diagrams: zones fade, arrows draw, boxes and text pop in sequence
    all(".scroll svg").forEach(svg => {
      const paths = all("path.ar", svg), boxes = all("rect.box", svg), text = all("text", svg), labels = all("rect.mask", svg).filter(r => !r.matches(".bg")), zones = all("rect.zone", svg);
      paths.forEach(p => { if (!p.classList.contains("opt")) { const L = p.getTotalLength(); p.style.strokeDasharray = L; p.style.strokeDashoffset = L; } });
      G.set([...paths, ...boxes, ...text, ...labels, ...zones], {opacity: 0});
      const tl = G.timeline({paused: true, defaults: {ease: "power2.out"}})
        .to(zones, {opacity: 1, duration: .5, stagger: .1})
        .to(boxes, {opacity: 1, duration: .45, stagger: .04}, "-=.2")
        .to(text, {opacity: 1, duration: .4, stagger: .01}, "<")
        .to(paths, {opacity: 1, strokeDashoffset: 0, duration: .7, stagger: .06}, "-=.2")
        .to(labels, {opacity: 1, duration: .4, stagger: .02}, "-=.4")
        .add(() => paths.forEach(p => { p.style.strokeDasharray = ""; p.style.strokeDashoffset = ""; }));
      ST.create({trigger: svg, start: "top 80%", once: true, onEnter: () => tl.play()});
    });
  }
  // spotlight
  all(".card").forEach(c => c.classList.add("spot"));
  document.addEventListener("pointermove", e => {
    const el = e.target.closest && e.target.closest(".spot"); if (!el) return;
    const r = el.getBoundingClientRect(); el.style.setProperty("--mx", e.clientX - r.left + "px"); el.style.setProperty("--my", e.clientY - r.top + "px");
  }, {passive: true});
  // magnetic buttons
  all(".hero .btn.p").forEach(b => {
    const qx = G.quickTo(b, "x", {duration: .4, ease: "power3"}), qy = G.quickTo(b, "y", {duration: .4, ease: "power3"});
    b.addEventListener("pointermove", e => { const r = b.getBoundingClientRect(); qx((e.clientX - r.left - r.width / 2) * .25); qy((e.clientY - r.top - r.height / 2) * .35); });
    b.addEventListener("pointerleave", () => { qx(0); qy(0); });
  });
}
})();
