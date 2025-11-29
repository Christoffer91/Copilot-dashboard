# Copilot Impact Dashboard

Author: Christoffer Besler Hansen  
License: AGPL-3.0

A client-only dashboard for exploring Viva Insights Copilot CSV exports: filter by org/region, compare metrics, export charts, and ship password-protected snapshots or a SharePoint-ready bundle.

## Features
- Upload Viva Insights Copilot CSVs (sample included at `samples/copilot-sample.csv`).
- Rich filters, multi-metric charts, agent hub tables, theme toggles, and per-capability insights.
- Exports: PNG/PDF, Excel summaries, encrypted snapshots, and SharePoint bundle generator.
- Privacy-first: processing happens in the browser; local caching is opt-in each session.

## Quick start
1) Clone the repo.  
2) Open `copilot-dashboard.html` in a modern Chromium-based browser (Chrome/Edge) for full feature support.  
3) Use "Select CSV" or "Load sample dataset" to explore.

Optional: serve locally for stricter CSP behavior, e.g. `python3 -m http.server 8000` then visit `http://localhost:8000/copilot-dashboard.html`.

## Data expectations
Minimum headers to unlock the main views (see `samples/copilot-sample.csv` for shape):
- Identity & grouping: `PersonId`, `Organization`, `CountryOrRegion`, `Domain`, `MetricDate` (ISO date)
- Activity metrics: `Total Active Copilot Actions Taken`, `Copilot assisted hours`
- App-specific metrics (used when present): meeting recaps/summaries, email/chat/powerpoint/word/excel actions, Copilot enabled flags, etc.

How to export the CSV from Viva Insights:
- You need admin access to Viva Insights.
- Go to https://analysis.insights.cloud.microsoft/ and run the **Microsoft 365 Copilot impact** report.
- Before running, add all available metrics from **Microsoft 365 Copilot** to the report to maximize coverage.
- Ensure your Organizational data mapping in Viva Insights aligns with the dashboard defaults: `Organization` for offices and `CountryOrRegion` for country. If your tenant uses different fields, map them accordingly before exporting.

## Privacy & storage
- All parsing/rendering runs locally; no analytics or remote posts.
- Dataset persistence uses `localStorage` but is **off by default** and requires explicit opt-in each session via the “Save dataset on this device” checkbox.
- Encrypted snapshots use Web Crypto; keep your password safe—there is no recovery.

## Dependencies (runtime)
- CDN: `pako` (gzip), `chart.js`, `papaparse`, `gif.js`, `xlsx`, `html2canvas`, `jspdf`, Inter font.
- Local: `assets/vendor/fflate.min.js` (compression fallback), `assets/sharepoint-static-assets.js` (SharePoint bundle styling), `assets/copilot-dashboard.js/css`.

## Development
- Entry point: `copilot-dashboard.html`
- Scripts/styles: `assets/copilot-dashboard.js`, `assets/copilot-dashboard.css`
- Sample data: `samples/copilot-sample.csv`

Contributions: open PRs/issues. Please keep CSP-tight changes and avoid introducing analytics without a clear opt-in.

## License
Copyright (c) 2025 Christoffer Besler Hansen.  
Licensed under the GNU Affero General Public License v3.0 (see `LICENSE`).
