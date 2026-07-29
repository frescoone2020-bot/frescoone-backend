# Finance Analyzer

A local-first Python pipeline that turns raw e-commerce exports (marketplace orders, dropship/reseller invoices, bank statements, payroll) into:

- an accrual **Net Profit dashboard** (with a separate, clearly-labeled cash view),
- a local **product-costing database** (import price, currency, invisible-costing fees, RSP/DSP) with a full add/search/edit/bulk-edit UI, and
- a **restock plan** based on recent sales velocity and current stock.

Everything runs on your own machine. There's no server, no external database, and no cloud upload by default — the dashboard is a single static HTML file.

## Setup

```bash
pip install pandas openpyxl pdfplumber
```

## Running it

```bash
python finance_analyzer.py
```

This reads whatever data files it finds in the project folder (see each script's module docstring for the exact filename patterns it looks for) and writes:

- `Executive_Financial_Report.md` — narrative report
- `dashboard.html` — interactive local dashboard (open it in a browser, or run `Start_Dashboard.bat` on Windows to serve + open it automatically)
- `history/<month>.json` — one snapshot per run, so trend charts accumulate real history over time

## Scripts

| File | Purpose |
|---|---|
| `finance_analyzer.py` | Main pipeline — loads all data sources, computes Net Profit, generates the report and dashboard (including the ISKU Manager tab). |
| `migrate_isku_database.py` | Builds/rebuilds the local product-costing database (`isku_database.json`) from a spreadsheet export. Re-runnable. |
| `restock_planner.py` / `restock_pipeline.py` / `restock_strategy.py` | Restock quantity/budget planning on top of the same cost and sales data. |
| `Start_Dashboard.bat` | Windows convenience launcher — serves the generated dashboard on `localhost` and opens it. |

## What's excluded from this repo, and why

This repo is the **tool only**. Everything listed below is real business data or proprietary operational context, and is excluded via `.gitignore` so it never gets committed:

- **Generated outputs** — `dashboard.html`, `Executive_Financial_Report.md`, and the restock report files. These have real figures baked directly into them at generation time, not placeholders.
- **Source data** — every `.xlsx`/`.csv`/`.zip` in the project root, plus the data subfolders (`dropship-reseller-file/`, `past-6months-payroll-history/`, the SiteGiant export folders, etc.) and `history/`. Bank statements, payroll, supplier invoices, and marketplace order exports all live here.
- **The costing database** — `isku_database.json` and `cost_overrides.json` (real supplier pricing).
- **Internal ops docs** — files that name the operating company, its brands, or its founder, and describe their specific internal workflow rather than generic tool usage.

If you're adapting this for your own business: drop your own exports into the project root following the naming patterns each script's docstring describes, run `finance_analyzer.py`, and everything above will populate locally — still gitignored, still private to your machine.
