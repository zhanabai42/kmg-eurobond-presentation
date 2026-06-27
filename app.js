/* KMG Investor Presentation — app.js */
(function () {
  'use strict';

  let slides = [];
  let currentSlide = 0;

  /* ── Fetch data and boot ──────────────────────────────────── */
  fetch('data.json')
    .then(r => r.json())
    .then(data => {
      slides = data.sections.flatMap(s => s.slides);
      renderMetrics(data.metrics);
      renderTOC(data.sections);
      renderGallery(slides);
      updateFooterDate(data.lastUpdated);
    })
    .catch(() => console.warn('data.json not found; using static HTML'));

  /* ── Metrics ─────────────────────────────────────────────── */
  function renderMetrics(metrics) {
    const grid = document.getElementById('metrics-grid');
    if (!grid) return;
    grid.innerHTML = metrics.map(m => `
      <div class="metric-card">
        <div class="metric-card__label">${m.label}</div>
        <div class="metric-card__value">${m.value}</div>
        <div class="metric-card__sub">${m.sub}</div>
        ${m.change ? `<div class="metric-card__change">${m.change}</div>` : ''}
      </div>
    `).join('');
  }

  /* ── Table of Contents ───────────────────────────────────── */
  function renderTOC(sections) {
    const grid = document.getElementById('toc-grid');
    if (!grid) return;
    grid.innerHTML = sections.map(sec => `
      <div class="toc-group">
        <div class="toc-group__title">${sec.title}</div>
        ${sec.slides.map(s => `
          <div class="toc-slide">
            <span class="toc-slide__num">${String(s.num).padStart(2,'0')}</span>
            <span class="toc-slide__title">${s.title}</span>
          </div>
        `).join('')}
      </div>
    `).join('');
  }

  /* ── Gallery ─────────────────────────────────────────────── */
  function renderGallery(slideList) {
    const grid = document.getElementById('gallery-grid');
    if (!grid) return;
    grid.innerHTML = slideList.map((s, i) => {
      const imgPath = `assets/slides/slide-${String(s.num).padStart(2,'0')}.jpg`;
      return `
        <div class="slide-thumb" data-index="${i}" tabindex="0" role="button"
             aria-label="Slide ${s.num}: ${s.title}">
          <div class="slide-thumb__placeholder" id="ph-${i}">
            <div class="slide-thumb__placeholder-num">${String(s.num).padStart(2,'0')}</div>
            <div class="slide-thumb__placeholder-title">${s.title}</div>
          </div>
          <div class="slide-thumb__caption">
            <span class="slide-thumb__num-badge">${String(s.num).padStart(2,'0')}</span>
            <span class="slide-thumb__title">${s.title}</span>
          </div>
        </div>
      `;
    }).join('');

    /* lazy-load real images if they exist */
    slideList.forEach((s, i) => {
      const img = new Image();
      const imgPath = `assets/slides/slide-${String(s.num).padStart(2,'0')}.jpg`;
      img.onload = () => {
        const ph = document.getElementById(`ph-${i}`);
        if (ph) {
          const imgEl = document.createElement('img');
          imgEl.src = imgPath;
          imgEl.alt = `Slide ${s.num}: ${s.title}`;
          imgEl.style.width = '100%';
          imgEl.style.aspectRatio = '16/9';
          imgEl.style.objectFit = 'cover';
          ph.replaceWith(imgEl);
        }
      };
      img.src = imgPath;
    });

    /* click to open lightbox */
    grid.addEventListener('click', e => {
      const thumb = e.target.closest('.slide-thumb');
      if (thumb) openLightbox(parseInt(thumb.dataset.index, 10));
    });
    grid.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        const thumb = e.target.closest('.slide-thumb');
        if (thumb) { e.preventDefault(); openLightbox(parseInt(thumb.dataset.index, 10)); }
      }
    });
  }

  /* ── Lightbox ────────────────────────────────────────────── */
  const lightbox = document.getElementById('lightbox');

  function openLightbox(idx) {
    currentSlide = idx;
    renderLightboxSlide();
    lightbox.classList.add('is-open');
    document.body.style.overflow = 'hidden';
    lightbox.focus();
  }

  function closeLightbox() {
    lightbox.classList.remove('is-open');
    document.body.style.overflow = '';
  }

  function renderLightboxSlide() {
    const s = slides[currentSlide];
    if (!s) return;
    const imgPath = `assets/slides/slide-${String(s.num).padStart(2,'0')}.jpg`;
    const imgWrap = document.getElementById('lightbox-img');
    const caption = document.getElementById('lightbox-caption');

    /* try real image, fall back to placeholder */
    const testImg = new Image();
    testImg.onload = () => {
      imgWrap.innerHTML = `<img src="${imgPath}" alt="Slide ${s.num}: ${s.title}">`;
    };
    testImg.onerror = () => {
      imgWrap.innerHTML = `
        <div class="lightbox__placeholder">
          <div class="lightbox__placeholder-num">${String(s.num).padStart(2,'0')}</div>
          <div class="lightbox__placeholder-title">${s.title}</div>
        </div>`;
    };
    testImg.src = imgPath;

    caption.innerHTML = `<strong>Slide ${s.num} / ${slides.length}</strong> — ${s.title}`;
  }

  function navigate(dir) {
    currentSlide = (currentSlide + dir + slides.length) % slides.length;
    renderLightboxSlide();
  }

  if (lightbox) {
    document.getElementById('lb-close').addEventListener('click', closeLightbox);
    document.getElementById('lb-prev').addEventListener('click', () => navigate(-1));
    document.getElementById('lb-next').addEventListener('click', () => navigate(1));
    lightbox.addEventListener('click', e => { if (e.target === lightbox) closeLightbox(); });
    document.addEventListener('keydown', e => {
      if (!lightbox.classList.contains('is-open')) return;
      if (e.key === 'ArrowLeft')  navigate(-1);
      if (e.key === 'ArrowRight') navigate(1);
      if (e.key === 'Escape')     closeLightbox();
    });
  }

  /* ── Mobile nav ──────────────────────────────────────────── */
  const burger = document.getElementById('nav-burger');
  const navLinks = document.getElementById('nav-links');
  if (burger && navLinks) {
    burger.addEventListener('click', () => {
      navLinks.classList.toggle('is-open');
      burger.setAttribute('aria-expanded', navLinks.classList.contains('is-open'));
    });
    navLinks.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', () => navLinks.classList.remove('is-open'));
    });
  }

  /* ── Scroll-spy active nav ───────────────────────────────── */
  const sections = document.querySelectorAll('section[id]');
  const navAnchors = document.querySelectorAll('.nav__links a[href^="#"]');
  const observer = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        navAnchors.forEach(a => {
          a.style.color = a.getAttribute('href') === '#' + entry.target.id
            ? 'var(--gold-light)' : '';
        });
      }
    });
  }, { rootMargin: '-40% 0px -50% 0px' });
  sections.forEach(s => observer.observe(s));

  /* ── Footer date ─────────────────────────────────────────── */
  function updateFooterDate(date) {
    const el = document.getElementById('footer-date');
    if (el && date) el.textContent = date;
  }
})();
