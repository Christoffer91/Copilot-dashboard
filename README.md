# Copilot Dashboard v2

A client-only dashboard for analyzing the **Viva Copilot Dashboard** CSV export locally in the browser (supports 100k+ rows via worker-based aggregation).

## Quick start
1) Serve the folder locally (recommended): `python -m http.server 8000`
2) Open `http://localhost:8000/copilot-dashboard-v2/copilot-dashboard.html`
3) Upload your Viva Copilot Dashboard CSV export (or click “Load sample”)

Note: Dimension and date filters are applied by re-processing the file in a Web Worker (keeps memory usage low for large exports).
Exports: summary CSV + per-chart PNG (via buttons or the Export menu).

## Data expectations
Required columns:
- `PersonId`
- `MetricDate` (ISO date)
- `Organization`
- `FunctionType`

All other columns are treated as numeric metrics when possible (e.g. “Total Copilot actions taken”, “Copilot actions taken in Teams”, chat prompts submitted, active/enabled days, meeting recap/summarize metrics).

Sample: `copilot-dashboard-v2/samples/viva-copilot-dashboard-sample.csv`

## Privacy & security
- Processing happens locally; no automatic uploads/analytics.
- Treat CSV as untrusted input: UI uses text rendering (no HTML injection) and keeps CSP tight.

## Planning docs
- `copilot-dashboard-v2/PLAN.md`
- `copilot-dashboard-v2/TODO.md`
