"""Rebuild a run's output workbook with manual matches overlaid.

The formatted workbook is never edited in place — a manual match changes row
codes, removes rows from several unreconciled sections, and shifts every
Summary subtotal. Instead the run's saved engine results are reloaded, the
manual matches (mapped back to engine records via engine_ref) are applied as
if the engine had made them, and the original writer regenerates the file.
"""

from __future__ import annotations

import base64
import json
import pickle
import tempfile
from datetime import date
from pathlib import Path

from core import db
from engines.ma import build_excel
from engines.yardi import write_output_ph

MANUAL_RULE = "Manual match"


def _manual_match_members(con, run_id: int) -> list[dict]:
    groups = []
    for m in con.execute(
            "SELECT match_id FROM matches WHERE run_id = ? AND is_manual = 1", (run_id,)):
        mid = m["match_id"]
        bank = [r["engine_ref"] for r in con.execute(
            "SELECT engine_ref FROM bank_txns WHERE match_id = ?", (mid,))]
        gl = [r["engine_ref"] for r in con.execute(
            "SELECT engine_ref FROM gl_txns WHERE match_id = ?", (mid,))]
        groups.append({"match_id": mid, "bank": bank, "gl": gl})
    return groups


def _overlay_yardi(results: dict, groups: list[dict]):
    bank_by_id = {b["id"]: b for b in results["all_bank"]}
    gl_by_id = {g["id"]: g for g in results["all_gl"]}
    for grp in groups:
        for bref in grp["bank"]:
            b = bank_by_id.get(bref)
            if b is None:
                continue
            b["matched"] = True
            b["match_rule"] = MANUAL_RULE
            b["match_ids"] = sorted(set((b.get("match_ids") or []) + grp["gl"]))
        for gref in grp["gl"]:
            g = gl_by_id.get(gref)
            if g is None:
                continue
            g["matched"] = True
            g["match_rule"] = MANUAL_RULE
            g["match_ids"] = sorted(set((g.get("match_ids") or []) + grp["bank"]))
    real_bank = [b for b in results["all_bank"]
                 if b.get("match_rule") != "INTERNAL – Stagecoach Sweep"]
    matched_bank = [b for b in real_bank if b.get("matched")]
    matched_gl = [g for g in results["all_gl"] if g.get("matched")]
    results["total_real_bank"] = len(real_bank)
    results["total_gl"] = len(results["all_gl"])
    results["n_matched_bank"] = len(matched_bank)
    results["n_matched_gl"] = len(matched_gl)
    results["pct_bank"] = round(len(matched_bank) / len(real_bank) * 100, 1) if real_bank else 0
    results["pct_gl"] = round(len(matched_gl) / len(results["all_gl"]) * 100, 1) if results["all_gl"] else 0
    return results


def _overlay_ma(results: dict, groups: list[dict]):
    def bank_idx(ref):
        return int(ref[1:]) if ref and ref.startswith("B") else None

    def gl_idx(ref):
        return int(ref[1:]) if ref and ref.startswith("G") else None

    matched_bank = results["matched_bank"]
    matched_gl = results["matched_gl"]
    for grp in groups:
        b_idxs = [i for i in (bank_idx(r) for r in grp["bank"]) if i is not None]
        g_idxs = [i for i in (gl_idx(r) for r in grp["gl"]) if i is not None]
        if not b_idxs:
            continue
        matched_bank[b_idxs[0]] = {"gl_indices": g_idxs, "rule": MANUAL_RULE}
        for bi in b_idxs[1:]:
            matched_bank[bi] = {"gl_indices": [], "rule": MANUAL_RULE}
        matched_gl.update(g_idxs)

    all_bank, all_gl = results["all_bank"], results["all_gl"]
    unmatched_bank_idx = [bi for bi in range(len(all_bank))
                          if bi not in matched_bank
                          and "STAGECOACH SWEEP" not in all_bank[bi]["description"].upper()]
    unmatched_gl_idx = [gi for gi in range(len(all_gl)) if gi not in matched_gl]
    total_bank_regular = sum(1 for b in all_bank
                             if "STAGECOACH SWEEP" not in b["description"].upper())
    results["unmatched_bank_idx"] = unmatched_bank_idx
    results["unmatched_gl_idx"] = unmatched_gl_idx
    results["stats"].update({
        "total_bank_regular": total_bank_regular,
        "total_matched_bank": total_bank_regular - len(unmatched_bank_idx),
        "total_gl_all": len(all_gl),
        "total_matched_gl": len(all_gl) - len(unmatched_gl_idx),
        "pct_bank": (100.0 * (total_bank_regular - len(unmatched_bank_idx)) / total_bank_regular
                     if total_bank_regular else 0),
        "pct_gl": (100.0 * (len(all_gl) - len(unmatched_gl_idx)) / len(all_gl)
                   if all_gl else 0),
    })
    return results


def rebuild_output(con, run_id: int) -> tuple[bytes, str]:
    """Regenerate the run's output workbook including its manual matches.
    Returns (xlsx_bytes, filename)."""
    run = con.execute("SELECT * FROM runs WHERE run_id = ?", (run_id,)).fetchone()
    if run is None:
        raise ValueError(f"run {run_id} not found")
    prop = db.get_property(con, run["property_code"])
    groups = _manual_match_members(con, run_id)
    stats = db.run_stats(run)

    stored = run["results_data"] or ""
    if not stored:
        raise ValueError("No results_data stored for this run — re-run the reconciliation.")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        tmp_path = tmp.name

    try:
        if run["gl_type"] == "yardi":
            results = json.loads(stored)
            results = _overlay_yardi(results, groups)
            write_output_ph.run(results, tmp_path, prop,
                                bank_ending_balance=stats.get("bank_ending_balance"))
        else:
            results = pickle.loads(base64.b64decode(stored))
            results = _overlay_ma(results, groups)
            y, m = (int(x) for x in run["period"].split("-"))
            build_excel.run(
                results, tmp_path,
                property_name=prop.property_name,
                property_title=prop.prop_display or prop.property_name,
                property_code=prop.property_code,
                gl_account_label=prop.gl_account_id,
                wfb_account=prop.account_info,
                period_start=date(y, m, 1),
                beginning_balance=stats.get("beginning_balance", 0.0),
                bank_ending_balance=stats.get("bank_ending_balance"))

        xlsx_bytes = Path(tmp_path).read_bytes()
        filename = Path(run["output_path"]).name if run["output_path"] else f"reconciliation_{run_id}.xlsx"
        return xlsx_bytes, filename
    finally:
        Path(tmp_path).unlink(missing_ok=True)
