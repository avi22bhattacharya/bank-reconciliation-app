"""Yardi pipeline orchestrator: mh_recon → reconcile_ph → write_output_ph,
with prior unreconciled items injected from the database."""

from __future__ import annotations

from pathlib import Path

from core import ingest
from core.db import PropertyMeta
from core.prior_items import (PriorItems, append_prev_bank_rows,
                              patch_deposit_numbers, write_yardi_prev_workbook)
from engines.yardi import mh_recon, reconcile_ph, write_output_ph

PREV_LABEL = "Prior Unrec GL"   # must contain "Unrec" (write_output_ph contract)


def run_pipeline(workdir, *, gl_path, gl_sheet, bank_path, bank_sheet,
                 deposit_register_path, prop: PropertyMeta, period: str,
                 prior: PriorItems, gl_header_row: int = 5) -> dict:
    """Returns {"results", "results_path", "output_path", "stats"}."""
    workdir = Path(workdir)
    p_label = ingest.period_label(period)            # e.g. "Mar 2026"

    # Bank statement must be .xlsx for openpyxl; prior bank rows are appended
    # to the converted copy so the engine treats them as ordinary rows.
    bank_xlsx = ingest.convert_to_xlsx(bank_path, workdir / "bank_converted.xlsx")
    if bank_xlsx == Path(bank_path):                 # already .xlsx → work on a copy
        import shutil
        bank_xlsx = workdir / "bank_converted.xlsx"
        shutil.copyfile(bank_path, bank_xlsx)
    prev_bank_first_row = append_prev_bank_rows(bank_xlsx, bank_sheet, prior)

    # map injected prior items to their existing DB rows (engine ids embed
    # sheet rows: BANK-<row> for bank, PREV-GL-<row> with data from row 2)
    prior_refs = {}
    for j, b in enumerate(prior.bank):
        if b.get("txn_hash"):
            prior_refs[f"BANK-{prev_bank_first_row + j}"] = b["txn_hash"]
    for i, g in enumerate(prior.gl):
        if g.get("txn_hash"):
            prior_refs[f"PREV-GL-{2 + i}"] = g["txn_hash"]

    # Prior GL items go through mh_recon as a second source sheet so the
    # enriched workbook has the exact shape reconcile_ph expects.
    prev_wb = workdir / "prior_unrec_gl.xlsx"
    write_yardi_prev_workbook(prior, prev_wb, prop.property_code)

    enriched = workdir / "GL_with_Deposit_Numbers.xlsx"
    mh_stats = mh_recon.run(
        deposit_register={"file": str(deposit_register_path),
                          "sheet": "Deposit_Register", "header_row": None},
        gl_sources=[
            {"label": "GL", "file": str(gl_path), "sheet": gl_sheet,
             "header_row": gl_header_row,
             "person_col_idx": 4, "date_col_idx": 2, "remarks_col_idx": 10},
            {"label": "Un-Reconcile GL", "file": str(prev_wb),
             "sheet": "Un-Reconcile GL", "header_row": 0,
             "person_col_idx": 4, "date_col_idx": 2, "remarks_col_idx": 9},
        ],
        output_path=str(enriched),
    )
    # restore deposit numbers persisted when the prior items were current
    patch_deposit_numbers(enriched, prior)

    results_path = workdir / "recon_results.json"
    results = reconcile_ph.run(
        bank_file=str(bank_xlsx), bank_sheet=bank_sheet,
        gl_file=str(enriched), gl_sheet="GL", prev_sheet="Un-Reconcile GL",
        current_bank_label=f"{p_label} Bank", current_gl_label=f"{p_label} GL",
        prev_gl_label=PREV_LABEL, json_out=str(results_path),
    )

    output_path = workdir / f"{period} {prop.property_name} Bank Rec Output.xlsx"
    write_output_ph.run(results, str(output_path), prop)

    stats = {
        "pct_bank": results["pct_bank"], "pct_gl": results["pct_gl"],
        "total_real_bank": results["total_real_bank"], "total_gl": results["total_gl"],
        "n_matched_bank": results["n_matched_bank"], "n_matched_gl": results["n_matched_gl"],
        "gl_opening_balance": results["gl_opening_balance"],
        "mh_recon": [list(s) for s in mh_stats],
    }
    return {"results": results, "results_path": str(results_path),
            "output_path": str(output_path), "stats": stats,
            "prior_refs": prior_refs}
