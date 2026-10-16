# Costos · a spreadsheet analytics demo

A public portfolio demonstration of the workflow from my private Costos app: organize spreadsheet records with Python, compare costs in a dashboard, and export the selected data to Excel.

This is a focused demonstration, not a copy of the production application. Every client, batch, material price, quantity, and pattern shown here is synthetic. It contains no company records, customer details, commercial rates, or production database.

I'm [Eduardo](https://github.com/js0596a), a Statistics student at American University. During my Curfimex internship, I worked with Excel records in Python and built Dash and Plotly dashboards with Codex assistance. This demo gives you a way to explore that kind of workflow without accessing company data.

## Take a look

The overview brings together the selected months, clients, and cost summaries. All amounts below are fictional.

![Costos overview with filters and summaries using synthetic data](assets/screenshots/overview.png)

Monthly comparisons and the Pareto chart show how costs are distributed in the generated sample.

![Monthly costs and Pareto chart using synthetic data](assets/screenshots/pareto.png)

## What you can explore

- Filter fictional records by client and month.
- Compare monthly costs and cost per kilogram.
- Use a Pareto chart to see which spending items account for most of the total.
- Review the batch-level records behind the charts.
- Download an Excel workbook containing the current selection and its summaries.

The demo starts with an actual Excel file generated from a fixed random seed: 72 fictional batches across May–August 2026. Python reads that file with openpyxl, cleans and checks its columns, and uses pandas for the summaries. Dash and Plotly provide the interface and charts.

## Run locally

Use Python 3.11 or newer. From the project directory:

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

Open [http://127.0.0.1:8050](http://127.0.0.1:8050). Set `PORT` if that port is already in use.

The app regenerates `demo/synthetic_records.xlsx` when it starts. That directory is ignored by Git. The input is created entirely in code; there is no connection to a private database or fallback to local company files. No external services or credentials are needed.

## How the numbers work

Each spreadsheet row represents one spending item in a batch. Costs are summed across those rows. Batch weight is counted once per batch, so cost per kilogram is total selected cost divided by total selected batch weight. Monthly comparisons use the same calculation.

The Pareto chart orders spending items by cost. Its highlighted set includes the item that first takes cumulative spending to at least 80%. It describes cost concentration; it does not measure savings. All currency values are illustrative USD.

The exported workbook includes Summary, Monthly, Batches, Material costs, and Detail sheets. It contains numeric snapshots of the selected data rather than live formulas.

## Check the calculations

```bash
python -m unittest discover -s tests -v
```

The checks cover weighted aggregation, filters, the Pareto threshold, Excel input validation, and reconciliation of exported records with the current selection.

## Files

- `demo_data.py` — deterministic fictional Excel records.
- `analysis.py` — parsing, validation, summaries and Excel export.
- `app.py` — the dashboard and filter/download callbacks.
- `assets/style.css` — dashboard styling.
- `tests/test_analysis.py` — calculation and export checks.

Built with Python, pandas, openpyxl, Dash and Plotly. Codex assisted with this public demo's implementation and debugging.
