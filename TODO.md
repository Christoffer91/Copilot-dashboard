# Copilot Dashboard v2 — TODO (after plan approval)

## 0) Clarifications to confirm (blockers)
- [ ] Confirm whether v2 must support **multi-file import** and de-duplication by `(PersonId, MetricDate)`
- [ ] Confirm which dimensions must be filterable: `Organization`, `FunctionType` (any more?)
- [ ] Confirm default reporting window: last 28 days vs “All time”
- [ ] Confirm whether we need **SharePoint/offline bundle** support in v2 (yes/no)
- [ ] Provide 1–2 real CSV exports (sanitized) or confirm the sample header shape is stable

---

## 1) Project setup (v2 isolation)
- [ ] Keep `copilot-dashboard/` unchanged; all edits in `copilot-dashboard-v2/`
- [ ] Update `copilot-dashboard-v2/README.md` to describe v2 purpose + supported export type
- [ ] Rename page title/branding in `copilot-dashboard-v2/copilot-dashboard.html` to “Copilot Dashboard v2”
- [ ] Decide whether to keep v2 filenames as-is or rename to `copilot-dashboard-v2.html` (avoid confusion)

---

## 2) Data contract (Viva Copilot Dashboard export)
- [ ] Define **required columns**: `PersonId`, `MetricDate`, `Organization`, `FunctionType`
- [ ] Define **optional columns** and groups (examples):
  - [ ] Totals: `Total Copilot actions taken`, `Total Copilot active days`, `Total Copilot enabled days`
  - [ ] App actions: `Copilot actions taken in Word/Excel/Teams/Outlook/Powerpoint`
  - [ ] Chat prompts: `Copilot chat (work) prompts submitted`, `Copilot Chat (web) prompts submitted`, “without license …”
  - [ ] Active/enabled days per workload: `Days of active Copilot usage in *`, `Copilot enabled days for *`
  - [ ] Meetings: recapped/summarized counts and hours
- [ ] Implement **header normalization** rules (case-insensitive, collapse whitespace, normalize parentheses)
- [ ] Add a schema validation report:
  - [ ] Missing required columns
  - [ ] Unknown columns count
  - [ ] Detected metric groups
  - [ ] Row count, date range, blank/missing Organization/FunctionType rates

---

## 3) Performance-first ingestion & aggregation (worker)
- [ ] Decide approach:
  - [ ] A) Web Worker does parsing + aggregation (recommended)
  - [ ] B) PapaParse `worker:true` but aggregate on main thread (fallback)
- [ ] Implement cancellation (AbortController-like messaging)
- [ ] Implement progress:
  - [ ] bytes read (if available) and/or rows processed
  - [ ] “Parsing headers… / Aggregating… / Finalizing…”
- [ ] Implement aggregation outputs (minimal memory):
  - [ ] Global totals per metric
  - [ ] Time series per period (week/month) per metric
  - [ ] Dimension totals (Organization / FunctionType) per metric
  - [ ] Optional: dimension time series (top N dims only)
  - [ ] Exact unique users (no `Set` per period):
    - [ ] Active users per period (global)
    - [ ] Active users per period per dimension (optional, top N only)
- [ ] Guardrails:
  - [ ] Max file size warning + “continue anyway”
  - [ ] Max row count warning + “continue anyway”
  - [ ] Detect pathological CSV (excessive parse errors) and stop

---

## 4) Metric catalog & formatting
- [ ] Create a metric registry:
  - [ ] `id`, `label`, `group` (Chat/Word/Excel/Teams/Outlook/PowerPoint/Meetings/Other)
  - [ ] `unit` (count/days/hours), `format` (int/1-decimal)
  - [ ] `sourceColumns` (normalized header aliases)
- [ ] Add derived metrics:
  - [ ] Actions per active user
  - [ ] Enabled-day utilization = active days / enabled days (when both exist)
  - [ ] Chat mix: work vs web vs unlicensed (when present)
- [ ] Decide how to treat blanks:
  - [ ] Blank Organization -> “(blank)” bucket
  - [ ] Blank FunctionType -> “(blank)” bucket

---

## 5) UI layout (single page)
- [ ] Upload panel:
  - [ ] Drag/drop, file picker, progress, cancel, clear
  - [ ] Schema summary + warnings
- [ ] Filters:
  - [ ] Timeframe quick presets (7/28/90/all/custom)
  - [ ] Date range (from/to)
  - [ ] Dimension selector (Organization / FunctionType)
  - [ ] Metric selector (grouped + searchable)
  - [ ] Aggregation selector (daily/weekly/monthly; default weekly)
- [ ] Summary cards:
  - [ ] Total actions, active users, active days, enabled days
  - [ ] Optional: Copilot Chat key KPIs if present
- [ ] Charts:
  - [ ] Trend line (selected metric)
  - [ ] Breakdown bar (top N dims)
  - [ ] Workload stacked bar/pie (optional)
- [ ] Drilldown tables:
  - [ ] Top orgs/functions with totals + per-user normalization
  - [ ] Search + pagination or virtualized rendering (avoid DOM bloat)

---

## 6) Accessibility (WCAG 2.1 AA)
- [ ] Ensure all inputs have `<label>` and programmatic names
- [ ] Keyboard navigation:
  - [ ] Tab order sane
  - [ ] Skip link works
  - [ ] Dialogs trap focus + return focus
- [ ] Color/contrast check for default themes + chart palettes
- [ ] Provide text summaries for charts (screen-reader friendly)

---

## 7) Security hardening (front-end)
- [ ] Keep CSP tight (no `unsafe-inline`; restrict sources; worker-src)
- [ ] Avoid CSV-derived `innerHTML`; use `textContent`
- [ ] Export safety (formula injection):
  - [ ] Prefix `'` for cells starting with `=`, `+`, `-`, `@`
  - [ ] Apply to CSV and XLSX generation paths
- [ ] Avoid persistent storage by default; if added, make it opt-in per session

---

## 8) Exports
- [ ] “Download summary (CSV)” of aggregates (safe escaping + formula protection)
- [ ] “Download charts (PNG)”
- [ ] Optional: PDF export (selected sections)
- [ ] Optional: Excel export (aggregates only; avoid raw rows by default)

---

## 9) Validation checklist (manual)
- [ ] Small file (<= 1k rows): parses instantly; numbers match manual pivot in Excel
- [ ] Medium file (~100k rows): stays responsive; completes within acceptable time
- [ ] Large file (>= 250k rows): shows warning; cancel works; no browser crash
- [ ] Missing required columns: clear error message and next-step guidance
- [ ] Weird values:
  - [ ] `MetricDate` invalid -> row skipped + reported
  - [ ] numeric fields blank/`NaN` -> treated as 0
  - [ ] Organization contains commas/quotes/newlines -> renders safely and exports correctly
- [ ] Export injection test:
  - [ ] Org = `=HYPERLINK(...)` does not execute on opening exports

---

## 10) Definition of done (v1 for v2)
- [ ] Imports Viva Copilot Dashboard CSV export reliably
- [ ] Handles 100k+ rows without freezing UI
- [ ] Provides KPI summary + trends + top org/function drilldowns
- [ ] Exports aggregates safely (no formula injection)
- [ ] Accessibility baseline met (keyboard + labels + focus management)
- [ ] `copilot-dashboard/` remains unchanged

