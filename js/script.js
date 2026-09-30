// =========================================================
// SHEFALI DIXIT PORTFOLIO — minimal JS
// Handles: mobile nav toggle, footer year, contact form mailto fallback
// =========================================================

document.addEventListener('DOMContentLoaded', () => {

  // --- Render project cards from js/projects-data.js ---
  // To add/remove/update a project, edit projects-data.js only.
  const escapeHtml = (str) => String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  const projectsGrid = document.getElementById('projectsGrid');
  const projects = Array.isArray(window.PORTFOLIO_PROJECTS) ? window.PORTFOLIO_PROJECTS : [];

  if (projectsGrid) {
    if (projects.length === 0) {
      projectsGrid.innerHTML = '<p class="section-intro">No projects added yet — add one in js/projects-data.js.</p>';
    } else {
      projectsGrid.innerHTML = projects.map((p, i) => {
        const tags = Array.isArray(p.tags) ? p.tags : [];
        const images = Array.isArray(p.images) ? p.images.filter(Boolean) : [];
        const hasVideo = !!p.video;
        const hasCaseStudy = !!p.caseStudy;
        const hasGallery = images.length > 1;
        const opensPopup = hasCaseStudy || hasGallery || hasVideo;
        const hasPreviewLink = !!p.previewLink && p.previewLink !== '#';
        const hasSheetLink = !!p.sheetLink && p.sheetLink !== '#';
        const thumb = p.image || images[0] || '';

        // A project with a `video`, a `caseStudy`, and/or 2+ `images` opens
        // the in-page popup (data-project-index below) — full writeup if
        // caseStudy is set, plus the video/screenshot strip above it when
        // present. `previewLink`/`sheetLink` each get their own button that
        // opens straight in a new tab.
        let linksHtml = '';
        if (opensPopup) {
          const label = hasCaseStudy ? 'View case study →' : (hasVideo ? 'Watch demo →' : 'View screenshots →');
          linksHtml += `<a href="#" class="project-link" data-project-index="${i}">${label}</a>`;
        }
        if (hasPreviewLink) {
          linksHtml += `<a href="${escapeHtml(p.previewLink)}" class="project-link project-link-preview" target="_blank" rel="noopener">Live preview ↗</a>`;
        }
        if (hasSheetLink) {
          linksHtml += `<a href="${escapeHtml(p.sheetLink)}" class="project-link project-link-sheet" target="_blank" rel="noopener">View Sheet ↗</a>`;
        }
        if (!opensPopup && !hasPreviewLink && !hasSheetLink) {
          linksHtml = `<a href="#" class="project-link">View case study →</a>`;
        }

        // Thumbnail: a short muted looping preview clip if `video` is set,
        // otherwise the `image`/first `images` screenshot.
        const thumbHtml = hasVideo
          ? `<video class="project-thumb-video" src="${escapeHtml(p.video)}" ${thumb ? `poster="${escapeHtml(thumb)}"` : ''} autoplay muted loop playsinline></video>`
          : `<img src="${escapeHtml(thumb)}" alt="${escapeHtml(p.imageAlt || '')}">`;

        return `
        <article class="project-card">
          <div class="project-thumb">${thumbHtml}</div>
          <div class="project-body">
            <h3>${escapeHtml(p.title || 'Untitled project')}</h3>
            <p>${escapeHtml(p.description || '')}</p>
            <div class="project-tags">
              ${tags.map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join('')}
            </div>
            <div class="project-links">${linksHtml}</div>
          </div>
        </article>`;
      }).join('');
    }
  }

  // --- Case study popup (only for projects with a `caseStudy` field) ---
  const caseModal = document.getElementById('caseModal');
  const caseModalBody = document.getElementById('caseModalBody');
  // Tracks the element to return focus to when the modal closes.
  let lastFocusedBeforeModal = null;

  function openCaseStudy(project) {
    const images = Array.isArray(project && project.images) ? project.images.filter(Boolean) : [];
    const hasGallery = images.length > 1;
    const hasVideo = !!(project && project.video);
    if (!caseModal || !caseModalBody || !project || !(project.caseStudy || hasGallery || hasVideo)) return;

    const cs = project.caseStudy;
    const tags = Array.isArray(project.tags) ? project.tags : [];
    const videoHtml = hasVideo
      ? `<video class="case-modal-video" src="${escapeHtml(project.video)}" ${project.image ? `poster="${escapeHtml(project.image)}"` : ''} controls playsinline preload="metadata"></video>`
      : '';
    const galleryHtml = hasGallery
      ? `<div class="case-modal-gallery">${images.map((src) => `<img src="${escapeHtml(src)}" alt="${escapeHtml(project.title || '')} screenshot">`).join('')}</div>`
      : '';
    const stats = Array.isArray(project.stats) ? project.stats.filter((s) => s && s.value) : [];
    const statsHtml = stats.length
      ? `<div class="case-modal-stats">${stats.map((s) => `<div class="case-modal-stat"><strong>${escapeHtml(String(s.value))}</strong><span>${escapeHtml(s.label || '')}</span></div>`).join('')}</div>`
      : '';
    const bodyHtml = cs
      ? `
        <div class="case-modal-section"><h4>Challenge</h4><p>${escapeHtml(cs.challenge || '')}</p></div>
        <div class="case-modal-section"><h4>Approach</h4><p>${escapeHtml(cs.approach || '')}</p></div>
        <div class="case-modal-section"><h4>Result</h4><p>${escapeHtml(cs.result || '')}</p></div>
        ${cs.learnings ? `<div class="case-modal-section"><h4>Learnings</h4><p>${escapeHtml(cs.learnings)}</p></div>` : ''}
      `
      : (project.description ? `<div class="case-modal-section"><p>${escapeHtml(project.description)}</p></div>` : '');

    caseModalBody.innerHTML = `
      <h3 id="caseModalTitle">${escapeHtml(project.title || '')}</h3>
      <div class="case-modal-tags">${tags.map((t) => `<span class="tag">${escapeHtml(t)}</span>`).join('')}</div>
      ${statsHtml}
      ${videoHtml}
      ${galleryHtml}
      ${bodyHtml}
    `;
    caseModal.classList.add('open');
    caseModal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('case-modal-lock');
    lastFocusedBeforeModal = document.activeElement;
    const closeBtn = caseModal.querySelector('.case-modal-close');
    if (closeBtn) closeBtn.focus();
  }

  function closeCaseStudy() {
    if (!caseModal) return;
    caseModal.classList.remove('open');
    caseModal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('case-modal-lock');
    // Return focus to whatever opened the modal, so keyboard users don't
    // lose their place on the page.
    if (lastFocusedBeforeModal && typeof lastFocusedBeforeModal.focus === 'function') {
      lastFocusedBeforeModal.focus();
    }
    lastFocusedBeforeModal = null;
  }

  if (caseModal) {
    caseModal.querySelectorAll('[data-case-close]').forEach((el) => {
      el.addEventListener('click', closeCaseStudy);
    });
    document.addEventListener('keydown', (e) => {
      if (!caseModal.classList.contains('open')) return;
      if (e.key === 'Escape') {
        closeCaseStudy();
        return;
      }
      // Keep keyboard focus trapped inside the open modal (Tab / Shift+Tab
      // wrap around instead of escaping into the page behind it).
      if (e.key === 'Tab') {
        const focusable = caseModal.querySelectorAll(
          'a[href], button:not([disabled]), textarea, input, select, [tabindex]:not([tabindex="-1"])'
        );
        if (!focusable.length) return;
        const first = focusable[0];
        const last = focusable[focusable.length - 1];
        if (e.shiftKey && document.activeElement === first) {
          e.preventDefault();
          last.focus();
        } else if (!e.shiftKey && document.activeElement === last) {
          e.preventDefault();
          first.focus();
        }
      }
    });
  }

  document.querySelectorAll('.project-link[data-project-index]').forEach((link) => {
    const idx = Number(link.getAttribute('data-project-index'));
    const project = projects[idx];
    const images = Array.isArray(project && project.images) ? project.images.filter(Boolean) : [];
    if (project && (project.caseStudy || images.length > 1 || project.video)) {
      link.addEventListener('click', (e) => {
        e.preventDefault();
        openCaseStudy(project);
      });
    }
  });

  // --- Mobile nav toggle ---
  const navToggle = document.getElementById('navToggle');
  const navLinks = document.getElementById('navLinks');

  if (navToggle && navLinks) {
    navToggle.addEventListener('click', () => {
      const isOpen = navLinks.classList.toggle('open');
      navToggle.setAttribute('aria-expanded', isOpen);
    });

    // Close menu after tapping a link (mobile)
    navLinks.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navLinks.classList.remove('open');
        navToggle.setAttribute('aria-expanded', 'false');
      });
    });
  }

  // --- Footer year ---
  const yearEl = document.getElementById('year');
  if (yearEl) yearEl.textContent = new Date().getFullYear();

  // --- Contact form: submits to a Google Sheet via Apps Script web app ---
  // Every submission lands as a row (Timestamp | Name | Email | Message) in
  // the Sheet behind CONTACT_FORM_ENDPOINT — see scripts/contact-form-backend/Code.gs
  // for the backend code and setup steps. Uses mode:'no-cors' because Apps
  // Script web apps don't return CORS headers a browser can read; we can't
  // inspect the response, so we treat "the request didn't throw" as success
  // and fall back to opening the visitor's own email client only if the
  // network request itself fails (e.g. offline).
  const CONTACT_FORM_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxczSxFd0djoGf6Od5FA_EJ1aW-VNUvbu_4AgyR5f7BAw4HiPmSdZ2JVywXcuhBIkwPpA/exec';

  const contactForm = document.getElementById('contactForm');
  const formNote = contactForm ? contactForm.querySelector('.form-note') : null;
  const formNoteDefaultText = formNote ? formNote.textContent : '';

  function setFormNote(text, isError) {
    if (!formNote) return;
    formNote.textContent = text;
    formNote.style.color = isError ? '#B3261E' : '';
  }

  function mailtoFallback(name, email, message) {
    const subject = encodeURIComponent(`Portfolio inquiry from ${name}`);
    const body = encodeURIComponent(`${message}\n\n— ${name} (${email})`);
    window.location.href = `mailto:dixit.shefali123@gmail.com?subject=${subject}&body=${body}`;
  }

  if (contactForm) {
    contactForm.addEventListener('submit', (e) => {
      e.preventDefault();
      const submitBtn = contactForm.querySelector('button[type="submit"]');
      const name = document.getElementById('name').value.trim();
      const email = document.getElementById('email').value.trim();
      const message = document.getElementById('message').value.trim();

      if (submitBtn) submitBtn.disabled = true;
      setFormNote('Sending…', false);

      fetch(CONTACT_FORM_ENDPOINT, {
        method: 'POST',
        mode: 'no-cors',
        body: new URLSearchParams({ name, email, message }),
      })
        .then(() => {
          setFormNote("Thanks — your message is in! I'll get back to you soon.", false);
          contactForm.reset();
        })
        .catch(() => {
          setFormNote("Couldn't reach the server — opening your email app instead…", true);
          mailtoFallback(name, email, message);
        })
        .finally(() => {
          if (submitBtn) submitBtn.disabled = false;
          setTimeout(() => setFormNote(formNoteDefaultText, false), 6000);
        });
    });
  }

  // --- 3D tilt effect (hero portrait + project cards) ---
  // Skipped for touch devices and anyone who prefers reduced motion.
  const supportsHoverTilt = window.matchMedia('(hover: hover) and (pointer: fine)').matches;
  const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  if (supportsHoverTilt && !prefersReducedMotion) {
    const tiltEls = document.querySelectorAll('.hero-photo-frame, .project-card');

    tiltEls.forEach((el) => {
      const maxTilt = el.classList.contains('hero-photo-frame') ? 10 : 6;

      el.addEventListener('mousemove', (e) => {
        const rect = el.getBoundingClientRect();
        const px = (e.clientX - rect.left) / rect.width;  // 0 → 1
        const py = (e.clientY - rect.top) / rect.height;  // 0 → 1
        const rotateY = (px - 0.5) * (maxTilt * 2);
        const rotateX = (0.5 - py) * (maxTilt * 2);
        el.style.transform = `perspective(800px) rotateX(${rotateX}deg) rotateY(${rotateY}deg)`;
      });

      el.addEventListener('mouseleave', () => {
        el.style.transform = 'perspective(800px) rotateX(0deg) rotateY(0deg)';
      });
    });
  }

});