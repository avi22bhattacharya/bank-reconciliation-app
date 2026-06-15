# Bank Reconciliation Dashboard

Streamlit app that reconciles a property's monthly Bank Statement against its
General Ledger using the pre-existing Yardi and Manage America matching
engines, with prior-month unreconciled items persisted in SQLite.

## Run it

```bash
.env/bin/streamlit run app.py
```

Then open http://localhost:8501 (or the URL Streamlit prints).

## Pages

1. **Run** (`app.py`) — upload the GL workbook (bank statement can be in the
   same workbook or uploaded separately), pick **Yardi** or **Manage America**
   (Yardi also requires the Deposit Register), confirm the auto-detected
   property / period / sheets, and run. Prior unreconciled items are pulled
   from the database; on a property's first run an optional **Prior Open
   Items** upload seeds the history. The output workbook (Summary, Bank
   Statement, GL, Un-Reconcile transactions) is downloadable at the end.
2. **Manual Matching** — select any combination of leftover bank and GL rows;
   balanced selections (bank total = GL debit−credit) can be confirmed
   directly, unbalanced ones need an explicit override. Confirming updates the
   database and regenerates the output workbook.
3. **Matched Transactions** — browse match groups (filter by property, period,
   manual-only), export CSV, and see the run history.

## How it works

```
app.py                      Streamlit run page
pages/                      Manual Matching, Matched Transactions
core/
  db.py                     SQLite schema, content hashing, queries
  ingest.py                 .xls→.xlsx conversion (no LibreOffice needed),
                            sheet + property/period auto-detection
  prior_items.py            DB↔engine adapters, Open-Items bootstrap parsers
  pipeline_yardi.py         mh_recon → reconcile_ph → write_output_ph
  pipeline_ma.py            reconcile_full → build_excel
  persist.py                post-run persistence (dedup, transitions,
                            supersede on same-month re-runs)
  regenerate.py             manual-match overlay → output workbook rebuild
engines/
  yardi/, ma/               refactored copies of the original scripts
                            (matching logic unchanged; hardcoded paths became
                            parameters; the MA pickle-reload crash was fixed)
data/bank_rec.db            SQLite database
runs/<timestamp>_<prop>/    per-run working files + output workbook
```

Original scripts in `deposit-register-trial/` and `Manage America Bank recon/`
are untouched and serve as golden references — the refactored engines
reproduce their outputs cell-for-cell (verified against
`PH_Davie_Bank_Recon_Output.xlsx` / `recon_ph_results.json`).

### Database model

All properties share one database. Transactions are keyed by a content hash
(property, first-seen period, occurrence index, and the row's fields), so
re-running a month or carrying an unreconciled item forward never duplicates
rows. Items that were unreconciled and match in a later month flip to
`matched` automatically; `internal` marks bank-to-bank sweeps and contra
entries. Re-running a month supersedes the earlier run for that property +
period (the app warns when that would discard manual matches).

### Notes / limitations

- Several Manage America matching rules are Bay City-specific by design
  (e.g. "BAY CITY SETTLEMENT" descriptions) — that logic was deliberately
  left untouched; other MA properties will rely more on the generic
  check-number / amount rules.
- Manual matches require at least one bank row and one GL row.
- Re-running an older month after later months were already run is not
  supported (run months in order; re-running the latest month is fine).
