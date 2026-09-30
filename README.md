# Shefali Dixit — Portfolio Site

A static HTML/CSS/JS portfolio. No build step, no framework — open `index.html` directly or deploy as-is.

## Adding, removing, or updating a project

Edit **`js/projects-data.js`** only — it has full instructions in its own comments. Drop a screenshot in `assets/`, copy/edit one `{ ... }` block in that file, save, reload. No HTML or CSS editing required, and it works the same whether you're previewing locally, on this Cowork link, or once it's deployed to GitHub Pages/Netlify.

## Before you publish

1. **Résumé PDF**: export your résumé as a PDF and save it to `assets/Shefali_Dixit_Resume.pdf` (the download buttons already point to this path) — already done if you got this file from Claude.
2. **Photo**: replace `assets/avatar-placeholder.svg` with a real headshot. Easiest: add `assets/photo.jpg` and swap the `<img>` reference in `index.html`'s hero section (`hero-photo-frame`) if you want a real photo instead of the placeholder graphic.
3. **Project screenshots & links**: see "Adding, removing, or updating a project" above — everything project-related lives in `js/projects-data.js` now.
4. **Contact form**: it currently opens the visitor's email client via `mailto:`. For real form submissions, connect it to a service like Formspree or Netlify Forms (see comment in `js/script.js`).

## Deployment

Any static host works — no server required:
- **GitHub Pages**: push this folder to a repo, enable Pages on the `main` branch.
- **Netlify / Vercel**: drag-and-drop the folder in their dashboard, or connect the repo.
- **Custom domain**: point your domain's DNS to whichever host you choose, then add it in that host's settings.

## Structure

```
index.html              — page structure/content (sections: hero, about, skills, experience, résumé, projects, contact)
css/style.css           — all styling ("Chartwell" design: navy + gold, corporate/professional, 3D icon tiles, layered background)
js/script.js            — nav toggle, footer year, contact form, project-card rendering, 3D tilt effect
js/projects-data.js     — ⭐ EDIT THIS to add/remove/update projects — see its own comments
assets/                 — images, icons, and the résumé PDF
```
