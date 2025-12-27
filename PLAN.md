# Copilot Dashboard v2 (Viva Copilot Dashboard export) — Plan

> Scope: This plan covers **design + implementation approach** for a new static dashboard that analyzes **Viva Copilot Dashboard CSV exports** (can be 100k+ rows) entirely **client-side**.
>
> Constraints:
> - Do **not** modify `copilot-dashboard/` (original). All work happens in `copilot-dashboard-v2/`.
> - Static HTML/JS/CSS (no build pipeline assumed).
> - Privacy-first: no uploads, no analytics by default.

---

## Workflow Orchestrator Report
- Scope: New dashboard in `copilot-dashboard-v2/` for Viva Copilot Dashboard CSV exports (large files, many columns).
- Non-trivial: Yes (new UI + data pipeline + performance/security considerations).
- Reviews run: UX / Architecture / Red / Blue / Risk (+ Security-first).
- Risk decision (proposed): **APPROVED WITH CONDITIONS** (see Risk Assessment); manual approval required before implementation.
- Manual approval required: **Yes** (per requested gate before any functional changes).
- Next action: You review/approve this plan + conditions; then we implement in `copilot-dashboard-v2/`.

---

## Security-first Check
- Change implied: **Yes** (new dashboard capabilities + parsing/export surface).
- Security touchpoints: **input validation** (CSV parsing), **secrets** (none), **logging** (avoid sensitive), **CSP/SRI** (supply chain), **export safety** (CSV/Excel injection).
- Risk assessment: **Required** (included below).
- Notes:
  - Treat CSV as untrusted input (even if exported by Microsoft) because it can be modified before upload.
  - Avoid rendering CSV-derived strings via `innerHTML`. Prefer `textContent`.
  - Add export sanitization to prevent formula injection when generating CSV/XLSX.
  - Add file-size/row-count guardrails + worker-based parsing to mitigate client-side DoS.

---

## Architecture Review

### Scope
- Components:
  - `copilot-dashboard-v2/copilot-dashboard.html` (entry)
  - `copilot-dashboard-v2/assets/*` (JS/CSS + optional worker)
  - Optional: sample CSVs in `copilot-dashboard-v2/samples/`
- Constraints:
  - Static site only; must handle 100k+ rows without freezing UI.
  - Must support the Viva Copilot Dashboard export schema (wide CSV with many metrics columns).
  - Must remain privacy-first (local processing).

### Findings
1. **Data volume risk**: Storing per-row objects for 100k+ rows can exceed memory and cause GC stalls.
2. **Parsing/UI coupling risk**: Existing v1 code is monolithic; adding a new schema risks regression and complexity.
3. **Schema drift risk**: Column names may vary slightly (casing, spacing, parentheses); robust header normalization is required.
4. **Export safety risk**: Any “download to CSV/Excel” feature must guard against formula injection.
5. **Operational impact**: Static-only means no server-side validation/telemetry; resilience must be implemented in-browser (progress, cancel, guardrails).

### Recommended Approach
- Keep v2 as a **fork of v1’s UI patterns** (theme, cards, filters, export UX) but implement a **new ingestion + aggregation pipeline** optimized for wide/large CSVs:
  - Parse CSV in a **Web Worker** (or PapaParse worker mode) and aggregate incrementally.
  - Store **aggregates**, not full raw rows (optional “keep raw rows” toggle for smaller files only).
  - Use a **schema layer** that:
    - normalizes headers,
    - maps known Viva Copilot Dashboard columns to internal metric IDs,
    - exposes metric metadata (label, group/workload, units, formatting).
- UI renders from aggregates and supports:
  - timeframe + date range filtering,
  - org + FunctionType filtering,
  - metric selection (with sensible defaults),
  - drilldowns (top orgs, top functions, time series).

### Alternatives Considered
- **A) Extend v1 in-place**: lowest effort but high coupling/complexity and increases regression risk; also v1 keeps all rows in memory.
- **B) Rebuild minimal dashboard from scratch**: clean design, but loses tested UX features (exports, theming, snapshot, accessibility patterns).
- **C) Hybrid (recommended)**: reuse v1 UI components/styling patterns, but replace ingestion + data model with v2 aggregates.

### Open Questions
- Should v2 support **multi-file import** (e.g., weekly exports) and de-duplicate by `(PersonId, MetricDate)`?
- Are `Organization` and `FunctionType` always present? What is the desired behavior if one is blank?
- Are you primarily interested in:
  - adoption (active users, enabled days),
  - volume (actions/prompts),
  - workload breakdown (Word/Excel/Teams/etc),
  - or “Copilot Chat without license” tracking?
- Do we need SharePoint bundle generation in v2, or is “open locally” sufficient?

---

## UX Review Report (pre-implementation)

### Issue 1: Large-file import must feel safe
- Category: Interaction
- Severity: High
- Location: Upload/import flow
- Recommendation: Show deterministic progress, ETA-ish, and a **Cancel** button; keep UI responsive via worker parsing.

### Issue 2: Metric discoverability in a very wide schema
- Category: Content/Interaction
- Severity: High
- Location: Metric selector
- Recommendation: Group metrics by workload (Chat/Word/Excel/Teams/Outlook/PowerPoint/Other) + provide search.

### Issue 3: Accessibility for dense dashboards
- Category: Accessibility
- Severity: Medium
- Location: Filter bar + charts + tables
- WCAG Reference: 1.3.1, 2.1.1, 1.4.3
- Recommendation: Keyboard-first filter controls, descriptive chart summaries, proper labels, focus management for dialogs.

### Action Checklist
- [ ] Build upload panel with progress + cancel + clear error states
- [ ] Metric selector supports grouping + search + favorites/defaults
- [ ] All interactive elements have labels; dialogs trap focus; skip-link works

---

## Red Team Assessment (planned surface)

### Finding 1: Client-side DoS via oversized CSV
- Severity: High
- Attack Vector: Upload extremely large CSV / malformed CSV with pathological quoting to spike CPU/memory
- Prerequisites: Access to the page
- Impact: Browser freeze/crash; poor UX; possible data loss (unsaved filters)
- Safe Reproduction: Use a generated CSV with 1M+ rows; observe UI responsiveness
- Detection Signals: N/A (static); user-visible “page unresponsive”
- Mitigations: size/row guardrails, worker parsing, cancel, progressive aggregation, memory caps

### Finding 2: DOM XSS via CSV-derived labels
- Severity: Medium
- Attack Vector: Put HTML/JS payload in `Organization` or other label-like fields; rely on unsafe rendering
- Prerequisites: Any place the app uses `innerHTML` with untrusted data
- Impact: Script execution in origin context (if hosted)
- Safe Reproduction: Set Organization to `<img src=x onerror=alert(1)>` and verify it never executes
- Detection Signals: CSP violation reports (only if configured server-side)
- Mitigations: `textContent` everywhere, strict CSP, optional DOMPurify for any rich text (prefer none)

### Finding 3: CSV/Excel formula injection in exports
- Severity: Medium
- Attack Vector: Values beginning with `=`, `+`, `-`, `@` in exported rows
- Prerequisites: User exports and opens in Excel
- Impact: Potential data exfiltration prompts / malicious links when opened
- Safe Reproduction: Organization = `=HYPERLINK("http://example","click")` and export; open in Excel
- Mitigations: prefix `'` for formula-like cells on export; document behavior

---

## Blue Team Assessment (static app)

### Findings
1. **Supply chain**: CDN-loaded libraries are acceptable if pinned + SRI; prefer vendoring for offline/SharePoint bundles.
2. **Telemetry**: No analytics by default (privacy). If diagnostics needed later, add opt-in and avoid sending dataset content.
3. **Safe defaults**: Disable persistence by default; sanitize exports; guard against large files.

### Defensive Gaps
- No server-side logging/alerting (by design). User-facing diagnostics must be clear and non-sensitive.

### Recommended Mitigations
- Keep CSP tight; keep SRI; avoid `unsafe-inline`.
- Provide a “Diagnostics” panel that exposes only aggregate counters and parsing errors (no raw row dumps unless user explicitly copies).
- Add robust error boundaries: invalid headers, missing required columns, date parsing failures.

### Detection Signals
- CSP violation reports are only available if the app is hosted with `report-uri`/`report-to` configured (optional future enhancement).

---

## Risk Assessment Report

### Change Summary
New static dashboard (`copilot-dashboard-v2/`) that parses Viva Copilot Dashboard CSV exports locally and provides interactive filtering, charts, and safe exports.

### Assessment Scores
| Dimension | Score | Rationale |
|---|---:|---|
| UX Impact | 3/5 | New UI/workflow; needs strong accessibility + large-file UX |
| Maintenance Burden | 3/5 | Wide schema + perf considerations; needs good structure/tests |
| Breaking Changes | 1/5 | New folder; does not alter existing dashboard |
| Reversibility | 1/5 | Easy rollback by removing v2 folder |
| Dependency Impact | 2/5 | Reuses existing libs; may add one small helper (optional) |

### Final Risk Score: 2.4/5 — APPROVED WITH CONDITIONS

### Identified Risks
1. Performance regression (freezing on big files)
2. XSS / unsafe rendering via CSV labels
3. Formula injection via exported CSV/XLSX
4. Schema drift causing silent wrong numbers

### Mitigations Required (conditions)
- [ ] Worker-based parsing + incremental aggregation; do not keep full raw rows by default
- [ ] Header normalization + explicit “required fields” gating with clear error UI
- [ ] Export sanitization for CSV/XLSX
- [ ] Performance budgets + manual test checklist for 100k/250k rows

### Decision
- [x] APPROVED WITH CONDITIONS (requires your explicit approval before implementation begins)

---

## Proposed Dashboard (what it should look like)

### Pages/sections (single-page, scroll)
1. **Upload & Validation**
   - File picker + drag/drop
   - Schema summary: rows, date range, distinct orgs/functions, detected metric groups
   - Warnings: missing columns, blank org/function rates, parsing errors
2. **Executive Summary**
   - KPI cards (selectable timeframe):
     - Total Copilot actions taken
     - Active users (unique PersonId)
     - Total Copilot active days
     - Total Copilot enabled days
     - Copilot Chat usage (work/web/unlicensed) (if present)
3. **Trends**
   - Time series (weekly/monthly toggle):
     - Actions
     - Active users
     - Enabled days
4. **Workload Breakdown**
   - Stacked/Grouped bars by workload (Word/Excel/Teams/Outlook/PowerPoint/Other)
   - Metric-type toggles: Actions / Prompts submitted / Active days / Enabled days
5. **Org & FunctionType Drilldowns**
   - Top orgs/functions by selected metric
   - Small multiples or table + sparkline
6. **Exports**
   - “Download summary (CSV)”
   - “Download charts (PNG/PDF)”
   - Optional: SharePoint-ready offline bundle (later)

### Default filters
- Timeframe: last 28 days (if data supports), else “All time”
- Aggregation: weekly
- Dimension: Organization
- Metric: Total Copilot actions taken

---

## Implementation Plan (high-level)

### Milestone 0 — Fork and stabilize v2 baseline
- Copy v1 into `copilot-dashboard-v2/` (done)
- Rename visible branding to “Copilot Dashboard v2” and keep v1 untouched

### Milestone 1 — Schema + aggregation engine (worker-first)
- Implement header normalization and a “Viva Copilot Dashboard schema” map
- Parse + aggregate in worker, emitting:
  - global totals
  - per-period totals (week/month)
  - per-dimension totals (Org/FunctionType)
  - unique active users (exact, without per-period Set explosion)

### Milestone 2 — Core UI + filters
- Build filter state (timeframe/date range/org/function/metric/aggregation)
- Render KPI cards + trend chart + breakdown chart + top tables
- Provide strong empty/invalid states

### Milestone 3 — Exports + safety
- CSV export of aggregates (with formula injection protection)
- Chart exports (PNG/PDF) reusing v1 patterns
- Optional: snapshot/persistence (opt-in only)

### Milestone 4 — Hardening + validation
- Performance test checklist (100k/250k rows)
- Accessibility pass (keyboard, labels, contrast)
- Security spot-check (XSS, export injection, CSP)

