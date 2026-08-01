# AI Schedule Evaluation Dashboard — SIRMOUR

A Streamlit dashboard that evaluates an AI-generated power schedule against
actual meter data and automatically produces a DSM (Deviation Settlement
Mechanism) penalty report — no manual calculation required.

Upload an **AI Schedule file** (predicted generation) and a **Meter Data
file** (actual generation) for one day of 15-minute blocks, and the app will:

- Match both files block-wise on `Date` + `Block`
- Calculate deviation (MW and % of installed capacity)
- Calculate DSM penalty using **piecewise slab logic**
- Generate interactive Plotly charts
- Show KPI summary cards
- Give you a searchable / sortable / paginated block-wise table with
  highest-penalty and highest-deviation highlighting
- Let you download the merged dataset and a fully formatted, **formula-driven
  Excel DSM penalty report** that matches the plant's standard
  "Schedule vs Meter + Penalty" template (see below)

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the local URL Streamlit prints (usually http://localhost:8501),
upload the two files in the sidebar, and click **Process & Evaluate**.

## Input files

The app is intentionally flexible about column names and will auto-detect
common variants. At minimum each file needs:

1. **AI Schedule file** — a Date/Block (or Time/TimeStamp) column, plus a
   predicted-generation value column, e.g.:
   `Date, Block, Predicted_MW` or `Time, Predicted (MW)` or
   `Date, Block, Predicted (kW)` (kW is auto-converted to MW).
2. **Meter Data file** — the same, but for actual generation, e.g.:
   `Date, Block, Actual_MW` or `TimeStamp, Active Power (kW)`.

If a file has a `Time`/`TimeStamp` column instead of an explicit block
number, the app derives the 15-minute block from the time of day. Because
different SCADA/forecast systems label a block by its **start** or **end**
time, there's a toggle for each file under *Advanced: timestamp alignment*
in the sidebar (default: start-of-block, the most common convention).

CSV files with leading `#` comment/metadata rows (like typical backtest
exports) are handled automatically — those rows are skipped.

## Plant configuration

| Setting | Value |
|---|---|
| Plant Name | SIRMOUR |
| Installed Capacity | 5.1 MW |
| Scheduling Block | 15 minutes (96 blocks/day) |

## DSM penalty rules

```
Deviation MW  = Actual MW - Scheduled MW
Deviation %   = (Deviation MW / Installed Capacity) x 100
Block Energy  = Deviation MW x 0.25 x 1000   (kWh)
```

| Band | Rate |
|---|---|
| 0% – 10% | ₹0.00 / kWh |
| 10% – 15% | ₹0.50 / kWh |
| 15% – 20% | ₹0.75 / kWh |
| Above 20% | ₹1.00 / kWh |

**Penalty is charged with piecewise slab logic** — only the portion of the
deviation inside each slab is billed at that slab's rate (never the full
deviation at one flat rate). See the in-app **DSM Formula Explainer** tab
for a full worked example, or `calculator.py` for the implementation.

## Excel report format

The downloadable Excel report has two sheets:

1. **"Schedule vs Meter + Penalty"** — matches the plant's standard template
   exactly, including two charts placed directly on the sheet (to the right
   of the data table): AI Schedule vs Actual power (line) and DSM Penalty
   per block (bar). These are embedded as **static images** by default, so
   they display correctly in every viewer — Excel, LibreOffice, Google
   Sheets, and lightweight/quick file previews that don't run a full chart
   engine (native Excel chart objects often don't render in those quick
   previews, which is why images are the default). If chart figures aren't
   supplied, the report falls back to native, live Excel chart objects
   linked to the sheet's own cells instead.
2. **"Graphs"** — all 6 dashboard charts (deviation, cumulative penalty,
   over/under pie, histogram, etc.) embedded as static images, for a fuller
   visual summary alongside the data.

Sheet 1 layout:

- Title + note row, then a block-wise table (`Block, Time, AI Schedule (MW),
  Actual (MW), Deviation (MW), Deviation % (Capacity), Penalty (Rs),
  Scheduled at`).
- The **Deviation, Deviation % and Penalty columns are live Excel formulas**
  (not pasted numbers) — the Penalty formula references a **DSM Slab
  Parameters** table at the bottom of the sheet (installed capacity, block
  energy factor, and each slab's band/rate/upper-edge-in-MW), so the whole
  report is auditable and re-computes automatically if you edit a slab rate.
  Each formula cell also stores its already-computed value, so the numbers
  display correctly even in viewers that don't run a formula engine (e.g.
  pandas or a quick preview pane) — opening it in Excel/LibreOffice still
  shows (and can recalculate) the live formula.
- A **Day Summary** block below the data (blocks scored, total scheduled/
  actual energy, mean & max absolute deviation, blocks over/under 0.5 MW
  deviation, blocks penalised, worst single-block penalty, and the
  **TOTAL DSM PENALTY FOR THE DAY**), also formula-driven.
- Conditional formatting: deviation >0.5 MW in red, ≤0.5 MW in blue, and any
  block with a non-zero penalty in red.
- `Scheduled at` is populated automatically if the AI Schedule file has a
  "Generated At Interval" (or similarly named) column; otherwise it's left
  blank, matching the template.

## Project structure

```
app.py            Streamlit dashboard (UI, filters, tabs)
calculator.py      DSM deviation & piecewise penalty engine
utils.py            File loading, column auto-detection, validation, merge
graphs.py            Plotly chart builders (used for the in-app Charts tab)
excel_export.py     Formula-driven "Schedule vs Meter + Penalty" Excel report generator
config.py            Plant configuration & DSM slab registry
requirements.txt
README.md
sample_data/          Example input files for testing
```

## Extending to other plants

Add a new entry to `PLANT_CONFIGS` in `config.py`:

```python
PLANT_CONFIGS["MY_PLANT"] = PlantConfig(
    name="MY_PLANT",
    installed_capacity_mw=10.0,
    block_minutes=15,
    dsm_slabs=[
        DSMSlab(0, 12, 0.00, "0% - 12%"),
        DSMSlab(12, 20, 0.60, "12% - 20%"),
        DSMSlab(20, None, 1.20, "Above 20%"),
    ],
)
```

It will immediately appear in the plant selector in the sidebar — no other
code changes are needed. The installed capacity can also be overridden
directly in the sidebar for quick what-if analysis.

## Notes

- All processing happens in memory; no data is persisted outside your
  session.
- Duplicate blocks are de-duplicated (first occurrence kept); blocks present
  in only one of the two files are excluded from the DSM calculation and
  reported as a validation warning.
- The Excel report includes all 6 dashboard charts (Actual vs Predicted,
  deviation, penalty, cumulative penalty, over/under pie, deviation
  histogram) as **native, interactive Excel chart objects** — not static
  images — on both the main sheet and a separate, larger "Graphs" sheet.
  Because they're real Excel charts (built from worksheet formulas via
  XlsxWriter), you get hover tooltips, zoom, and editing, and the charts
  update automatically if you edit the underlying data. There is no
  kaleido / headless-browser dependency for this at all — chart export
  always works the same way on every machine.
