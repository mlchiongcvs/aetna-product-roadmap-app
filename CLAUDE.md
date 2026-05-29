# CLAUDE.md — Aetna Medicaid Product Roadmap App

## What This Is

An interactive strategic roadmap application for Aetna Medicaid Technology. It combines a React SPA with embedded standalone HTML pages to present the 3-year technology strategy, Gantt-chart roadmap with 49 funded initiatives, storyboard narrative, and operational dashboards.

**Live URL:** https://mlchiongcvs.github.io/aetna-product-roadmap-app/
**Repo:** https://github.com/mlchiongcvs/aetna-product-roadmap-app

---

## Tech Stack

- **Frontend:** React 18 + Vite (SPA)
- **Server:** Express (Node.js) on port 8080
- **Deployment:** GitHub Pages (gh-pages branch) for sharing; CAP (CVS container platform) for internal hosting
- **Build:** `npm run build` → outputs to `dist/`
- **Dev:** `npm run dev` (Vite HMR)
- **Prod:** `npm start` (Express serves `dist/` + proxies Rally API)

---

## Project Structure

```
src/
  AetnaMedicaidRoadmap.jsx   — Main React app (nav, dashboard, ingest, Rally, timeline, investments, team views)
  main.jsx                   — React entry point

public/
  strategy-overview.html     — Standalone strategy site (4 internal pages: overview, story, roadmap, assumptions)
  storyboard.html            — Horizontal scrolling narrative for leader onboarding

server/
  index.js                   — Express server (static files + Rally API proxy + /health endpoint)
  features.json              — Feature flags

.github/workflows/
  build-and-push.yml         — CI: build Docker image → push to JFrog
```

---

## Architecture: Embedded HTML Pages

The strategy pages (`strategy-overview.html`, `storyboard.html`) are complex standalone apps (2000+ lines each) with their own CSS, JS, SVGs, and interactive features. They are NOT React components — they're embedded via iframes in the React shell:

- `strategy-overview.html?page=overview|story|roadmap|assumptions` — URL param selects which internal page to show and hides the internal nav
- `storyboard.html` — hides internal nav when `window.self !== window.top` (iframe detection)

The React app provides the left nav and routes each tab to the correct iframe URL.

---

## Navigation Structure

| Nav Tab | View ID | What It Shows |
|---------|---------|---------------|
| 3-Year Overview | strat-overview | Strategy hero + 5 dials + investment breakdown |
| Story Arc | strat-story | 5-year evolution narrative |
| Roadmap | strat-roadmap | Interactive Gantt chart (49 initiatives, filterable) |
| Context & Assumptions | strat-context | Planning assumptions + risk register |
| Storyboard | storyboard | Horizontal scrolling onboarding narrative |
| Dashboard | dashboard | AI-powered roadmap management |
| Ingest | ingest | Document upload/analysis |
| Rally | rally | Rally API integration |
| Timeline | timeline | Visual timeline |
| Investments | investments | Capital portfolio view |
| Team | team | Team structure |

Default landing: **3-Year Overview** (`strat-overview`)

---

## Roadmap Gantt Chart (strategy-overview.html)

### Data Model

```javascript
const PLs = [
  { name, icon, color, ns (north star), pl (filter key "pl0"-"pl7"),
    products: [
      { name, initiatives: [
        { id, name, funded, src, desc, outcomes, users,
          capex26, capex27, ebit, nonEbit, roi, benefits, measures,
          priority: ["ef-cost", "bd-member", ...],  // 1-3 priority codes
          now, next, future }  // quarter ranges for Gantt positioning
      ]}
    ]
  }
]
```

### 10 Priorities (from Medicaid Priorities document)

**Excel at the Fundamentals:**
- `ef-footprint` — Expand Strategic Footprint
- `ef-cost` — Best-in-Class Cost Structure
- `ef-tcoc` — Total Cost of Care Solutions
- `ef-community` — Community-Integrated Support

**Be Truly Distinctive:**
- `bd-provider` — Improve Provider Experience
- `bd-member` — Enhance Member Experience
- `bd-state` — Advance State Partnerships

**Build a Winning Culture:**
- `wc-performance` — High-Performance Culture
- `wc-talent` — Modernize Talent Capabilities
- `wc-colleague` — Best-in-Class Colleague Experience

### Filter System

Filters use prefix-based keys:
- `all` — show everything
- `pl0`–`pl7` — by product line
- `pr-ef-cost`, `pr-bd-member`, etc. — by priority
- `u-member`, `u-provider`, etc. — by user type

### Key Functions
- `renderRM(filter)` — renders the full Gantt grid based on active filter
- `openPanel(initId)` — opens right-side flyout with initiative detail
- `toggleRmFilters()` — opens/closes filter panel
- `showPage(name)` — switches internal pages (overview/story/roadmap/assumptions)

---

## Deployment Workflows

### GitHub Pages (for colleague sharing)
```bash
npm run build
git checkout gh-pages
rm -rf assets index.html storyboard.html strategy-overview.html
cp dist/index.html . && cp -r dist/assets . && cp dist/storyboard.html . && cp dist/strategy-overview.html .
git add -A && git commit -m "Deploy: <description>"
git -c http.sslVerify=false push origin gh-pages
git checkout main
```

### CAP (internal hosting) — IN PROGRESS
- Docker image → JFrog (`cvsh.jfrog.io/cvsdigital-docker-local/`)
- CI workflow: `.github/workflows/build-and-push.yml`
- App name: `aiplatformclaudecode`
- Env vars: `ANTHROPIC_API_KEY`, `ANTHROPIC_BASE_URL` (hyperion gateway)
- Status: JFrog auth unresolved; service not yet created

---

## Environment Variables

| Variable | Value | Purpose |
|----------|-------|---------|
| ANTHROPIC_API_KEY | sk-aip-... | CVS Anthropic gateway auth |
| ANTHROPIC_BASE_URL | https://hyperion-lms-api.prod.cvshealth.com | CVS gateway proxy |
| PORT | 8080 | Express server port |

---

## Our Strategy (North Star)

> Improve population health and program sustainability by advancing total cost of care, strengthening state and provider partnerships, and delivering simpler, more personalized member experiences.

This statement appears on: Overview hero, Roadmap hero banner, Context/Assumptions page, Storyboard opener, and Investments view.

---

## Common Tasks

**Add a new initiative to the Gantt chart:**
Edit `public/strategy-overview.html` → find the `const PLs` array → add to the appropriate product line's `initiatives` array. Include `priority:["code"]` field.

**Change filter behavior:**
Edit `renderRM()` function in `strategy-overview.html` script block.

**Modify React nav/views:**
Edit `src/AetnaMedicaidRoadmap.jsx` → `NAV` array and corresponding render blocks.

**SSL push issues:**
Use `git -c http.sslVerify=false push origin <branch>` (CVS proxy cert issue).

---

## Key Decisions Made

- Standalone HTML pages embedded via iframe (not converted to React) — preserves 2000+ lines of complex interactive JS/CSS
- Priority-based color coding replaced original "strategic alignment" categories
- Gantt bars show EBIT and ROI inline for quick scanning
- Flyout panels (not tooltips) for both initiative detail and filters
- `?page=` URL param routing for strategy pages when embedded
