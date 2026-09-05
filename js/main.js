const nav = document.querySelector(".nav-links");
const toggle = document.querySelector(".menu-toggle");
const links = [...document.querySelectorAll(".nav-links a")];
const sections = [...document.querySelectorAll("main section[id]")];

toggle.addEventListener("click", () => {
  const open = nav.classList.toggle("is-open");
  toggle.setAttribute("aria-expanded", String(open));
});

links.forEach((link) => {
  link.addEventListener("click", () => {
    nav.classList.remove("is-open");
    toggle.setAttribute("aria-expanded", "false");
  });
});

document.addEventListener("click", (event) => {
  if (!nav.contains(event.target) && !toggle.contains(event.target)) {
    nav.classList.remove("is-open");
    toggle.setAttribute("aria-expanded", "false");
  }
});

const syncActive = () => {
  const y = window.scrollY + 100;
  let current = sections[0]?.id;
  for (const section of sections) {
    if (section.offsetTop <= y) current = section.id;
  }
  links.forEach((link) => {
    link.classList.toggle("is-active", link.getAttribute("href") === `#${current}`);
  });
};

window.addEventListener("scroll", syncActive, { passive: true });
syncActive();

const parseRegions = (value) =>
  value
    .split(";")
    .map((part) => part.split(",").map(Number))
    .filter((box) => box.length === 4 && box.every((n) => Number.isFinite(n)));

const boundsOf = (regions) => {
  const left = Math.min(...regions.map(([x]) => x));
  const top = Math.min(...regions.map(([, y]) => y));
  const right = Math.max(...regions.map(([x, , w]) => x + w));
  const bottom = Math.max(...regions.map(([, y, , h]) => y + h));
  return { left, top, width: right - left, height: bottom - top };
};

const resetShot = (shot) => {
  shot.classList.remove("is-zoomed");
  shot.style.removeProperty("--ox");
  shot.style.removeProperty("--oy");
  shot.style.removeProperty("--zoom");
  shot.querySelectorAll(".shot-ring").forEach((ring) => ring.remove());
  const reset = shot.querySelector(".shot-reset");
  if (reset) reset.hidden = true;
};

const zoomShot = (shot, regions) => {
  const stage = shot.querySelector(".shot-stage");
  const box = boundsOf(regions);
  const zoom = Math.min(3.4, Math.max(1.35, Math.min(88 / box.width, 78 / box.height)));

  shot.querySelectorAll(".shot-ring").forEach((ring) => ring.remove());
  regions.forEach(([x, y, w, h]) => {
    const ring = document.createElement("span");
    ring.className = "shot-ring";
    ring.style.setProperty("--rx", `${x}%`);
    ring.style.setProperty("--ry", `${y}%`);
    ring.style.setProperty("--rw", `${w}%`);
    ring.style.setProperty("--rh", `${h}%`);
    stage.appendChild(ring);
  });

  shot.style.setProperty("--ox", `${box.left + box.width / 2}%`);
  shot.style.setProperty("--oy", `${box.top + box.height / 2}%`);
  shot.style.setProperty("--zoom", String(zoom));
  shot.classList.add("is-zoomed");
  const reset = shot.querySelector(".shot-reset");
  if (reset) reset.hidden = false;
};

const cards = [...document.querySelectorAll(".card[data-zoom]")];

const clearZooms = () => {
  document.querySelectorAll(".shot.is-zoomed").forEach(resetShot);
  cards.forEach((card) => {
    card.classList.remove("is-active");
    const btn = card.querySelector(".zoom-btn");
    if (btn) btn.textContent = "نمایش در عکس";
  });
};

cards.forEach((card) => {
  card.addEventListener("click", () => {
    const shot = card.closest(".section")?.querySelector(".shot");
    const regions = parseRegions(card.dataset.zoom || "");
    if (!shot || !regions.length) return;

    if (card.classList.contains("is-active")) {
      clearZooms();
      return;
    }

    clearZooms();
    card.classList.add("is-active");
    const btn = card.querySelector(".zoom-btn");
    if (btn) btn.textContent = "بستن زوم";
    shot.scrollIntoView({ behavior: "smooth", block: "center" });
    zoomShot(shot, regions);
  });
});

document.querySelectorAll(".shot-reset").forEach((button) => {
  button.addEventListener("click", (event) => {
    event.stopPropagation();
    clearZooms();
  });
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") clearZooms();
});
