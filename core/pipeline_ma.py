"""Manage America pipeline orchestrator: reconcile_full → build_excel,
with prior unreconciled items injected from the database."""

from __future__ import annotations

from datetime import date
from pathlib import Path

from core import ingest
from core.db import PropertyMeta
from core.prior_items import PriorItems, to_ma_prev
from engines.ma import build_excel, reconcile_full

PREV_LABEL = "Prior (prev)"     # must contain "prev" (extra rule A preference)


def run_pipeline(workdir, *, gl_path, gl_sheet, bank_path, bank_sheet,
                 prop: PropertyMeta, period: str, prior: PriorItems,
                 bank_ending_balance: float | None = None) -> dict:
    """Returns {"results", "results_path", "output_path", "stats"}."""
    workdir = Path(workdir)
    p_label = ingest.period_label(period)            # e.g. "Apr 2026"
    y, m = (int(x) for x in period.split("-"))
    period_start = date(y, m, 1)

    gl_xlsx = ingest.convert_to_xlsx(gl_path, workdir / "gl_converted.xlsx")
    if str(bank_path) == str(gl_path):
        bank_xlsx = gl_xlsx
    else:
        bank_xlsx = ingest.convert_to_xlsx(bank_path, workdir / "bank_converted.xlsx")

    prev_bank, prev_gl = to_ma_prev(prior, PREV_LABEL)

    # injected prior items occupy indices 0..n-1 of all_bank/all_gl; map them
    # to their existing DB rows so persistence updates in place
    prior_refs = {}
    for i, b in enumerate(prior.bank):
        if b.get("txn_hash"):
            prior_refs[f"B{i}"] = b["txn_hash"]
    for i, g in enumerate(prior.gl):
        if g.get("txn_hash"):
            prior_refs[f"G{i}"] = g["txn_hash"]

    results_path = workdir / "recon_final.pkl"
    results = reconcile_full.run(
        workbook_path=str(gl_xlsx), bank_sheet=bank_sheet, gl_sheet=gl_sheet,
        prev_bank=prev_bank, prev_gl=prev_gl, current_label=p_label,
        bank_workbook_path=str(bank_xlsx), pkl_out=str(results_path),
    )

    beginning_balance = reconcile_full.read_gl_opening_balance(str(gl_xlsx), gl_sheet)

    output_path = workdir / f"{period} {prop.property_name} Reconciliation.xlsx"
    build_excel.run(
        results, str(output_path),
        property_name=prop.property_name,
        property_title=prop.prop_display or prop.property_name,
        property_code=prop.property_code,
        gl_account_label=prop.gl_account_id,
        wfb_account=prop.account_info,
        period_start=period_start, beginning_balance=beginning_balance,
        bank_ending_balance=bank_ending_balance,
    )

    stats = dict(results["stats"])
    stats["beginning_balance"] = beginning_balance
    stats["bank_ending_balance"] = bank_ending_balance
    return {"results": results, "results_path": str(results_path),
            "output_path": str(output_path), "stats": stats,
            "prior_refs": prior_refs}
