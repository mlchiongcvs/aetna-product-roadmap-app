# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

An interactive strategic roadmap application for Aetna Medicaid Technology. Combines a React SPA shell with embedded standalone HTML pages to present the 3-year technology strategy, Gantt-chart roadmap (69 initiatives across 12 product lines), storyboard narrative, and operational dashboards.

**Live URL:** https://mlchiongcvs.github.io/aetna-product-roadmap-app/
**Repo:** https://github.com/mlchiongcvs/aetna-product-roadmap-app

---

## Commands

```bash
npm run dev       # Vite dev server with HMR
npm run build     # Production build → dist/
npm run preview   # Preview production build locally
npm start         # Express server (serves dist/ + Rally API proxy) on port 8080
```

No test suite exists. No linter configured.

---

## Architecture

### React Shell + Embedded HTML Pages

The app has two layers:

1. **React SPA** (`src/AetnaMedicaidRoadmap.jsx`) — Provides the collapsible left sidebar nav, routing between views, and renders server-dependent features (Dashboard, Ingest, Rally).

2. **Standalone HTML pages** (`public/strategy-overview.html`, `public/storyboard.html`) — Complex self-contained apps (2000+ lines each) with their own CSS, JS, SVGs, and interactivity. Embedded via iframes in the React shell.

The iframe `src` attributes MUST use `import.meta.env.BASE_URL` (not absolute `/` paths) to work on GitHub Pages:
```jsx
<iframe src={`${import.meta.env.BASE_URL || '/'}strategy-overview.html?page=roadmap`} />
```

### URL-Based Page Routing in strategy-overview.html

`strategy-overview.html?page=overview|story|roadmap|assumptions` — the `?page=` param selects which internal page to show and hides the built-in nav when embedded in an iframe.

### Server

`server/index.js` — Express on port 8080. Serves `dist/` static files, proxies `/api/rally/*` to Rally, exposes `/health` for container health checks.

---

## Roadmap Gantt Chart Data Model (strategy-overview.html)

```javascript
const PLs = [
  { name, icon, color, ns, pl: "pl0"-"pl11",
    products: [
      { name, initiatives: [
        { id, name, funded, src, desc, outcomes, users,
          capex26, capex27, ebit, nonEbit, roi, benefits, measures,
          priority: ["ef-cost", ...],   // 1-3 priority codes
          category: "cat-ai",           // single category code
          now, next, future }           // quarter ranges for Gantt bar positioning
      ]}
    ]
  }
]
```

### 12 Product Lines (pl0–pl11)

pl0 Member Access, pl1 Enrollment, pl2 Claims, pl3 Provider & Network, pl4 Clinical & Care Mgmt, pl5 Finance & Billing, pl6 Data & Analytics, pl7 Contract & Compliance, pl8 Security, pl9 Engineering Excellence, pl10 Clinical, pl11 Medicaid Quality

### Filter System

Filters use prefix-based keys dispatched in `renderRM(filter)`:
- `all` — show everything
- `pl0`–`pl11` — by product line
- `cat-ai`, `cat-platform`, `cat-cost`, `cat-data`, `cat-exp`, `cat-security`, `cat-stability` — by category
- `pr-ef-cost`, `pr-bd-member`, etc. — by priority (prefixed `pr-`)
- `u-member`, `u-provider`, etc. — by user type

### Gantt Bar Positioning

`colsFor(init)` maps `now/next/future` strings to column indices:
- now: Q1-Q4 2026 → cols 0-3
- next: Q1-Q4 2027 → cols 4-7
- future: H1/H2 2028 → cols 8-9, 2029+ → col 10

### Key Functions
- `renderRM(filter)` — renders full Gantt grid; handles all filter types
- `openPanel(initId)` — opens flyout with initiative detail + category/priority badges
- `toggleRmFilters()` — opens/closes filter panel
- `showPage(name)` — switches internal pages

---

## Deployment

### GitHub Pages (for sharing)
```bash
npm run build
git checkout gh-pages
rm -rf assets index.html storyboard.html strategy-overview.html aetna-logo.svg
cp dist/index.html . && cp -r dist/assets . && cp dist/storyboard.html . && cp dist/strategy-overview.html . && cp dist/aetna-logo.svg .
git add -A && git commit -m "Deploy: <description>"
git -c http.sslVerify=false push origin gh-pages
git checkout main
```

Always push both branches after changes. Use `git -c http.sslVerify=false push` (CVS proxy SSL cert issue).

### CAP (container platform) — IN PROGRESS
- Docker image pushed to JFrog: `cvsh.jfrog.io/cvsdigital-docker-local/aetna-medicaid-roadmap`
- CI: `.github/workflows/build-and-push.yml` (triggered on push to main)
- CAP app: `aetnamedicaidroadmap`
- Status: JFrog credentials need regeneration

---

## Environment Variables (server-side only)

| Variable | Purpose |
|----------|---------|
| ANTHROPIC_API_KEY | CVS Anthropic gateway auth (Hyperion) |
| ANTHROPIC_BASE_URL | `https://hyperion-lms-api.prod.cvshealth.com` |
| PORT | Express port (default 8080) |

These are only needed for the Dashboard/Ingest AI features. The strategy pages, roadmap, and storyboard work entirely client-side.

---

## Common Tasks

**Add a new initiative:** Edit `public/strategy-overview.html` → find `const PLs` array → add to the appropriate product line's `initiatives` array. Must include `priority`, `category`, `now/next/future` fields.

**Add a new filter dimension:** Add the constant, extend `renderRM()` with a new prefix check for `plSet` and `filteredInits`, add buttons to the filter panel HTML.

**Modify React nav/views:** Edit `src/AetnaMedicaidRoadmap.jsx` → `NAV` array and corresponding render blocks.

---

## Key Design Decisions

- Standalone HTML pages embedded via iframe (not React) — preserves 2000+ lines of complex interactive JS/CSS per page
- `vite.config.js` uses `base: './'` (relative paths) — required for GitHub Pages subdirectory hosting
- Collapsible sidebar with `useState` toggle — shows icon-only nav when collapsed
- Gantt bars show EBIT and ROI inline for quick scanning
- Flyout panels (not tooltips) for initiative detail and filters
- Summary cards use colored top accent bars to differentiate metrics visually
