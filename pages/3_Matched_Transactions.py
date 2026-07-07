"""View reconciled (matched) transactions and run history."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core import db, persist
from core import regenerate
from core.auth import require_login

st.set_page_config(page_title="Matched Transactions", page_icon="✅", layout="wide")
require_login()
st.title("✅ Matched Transactions")


con = db.get_or_connect()

props = db.list_properties(con)
if not props:
    st.info("No properties yet — run a reconciliation first.")
    st.stop()

f1, f2, f3 = st.columns([2, 1, 1])
with f1:
    codes = [p.property_code for p in props]
    prop_code = st.selectbox("Property", codes,
                             format_func=lambda c: f"{c} — {next(p.property_name for p in props if p.property_code == c)}")
with f2:
    periods = [r["source_period"] for r in con.execute(
        """SELECT DISTINCT source_period FROM gl_txns WHERE property_code = ?
           UNION SELECT DISTINCT source_period FROM bank_txns WHERE property_code = ?
           ORDER BY source_period DESC""", (prop_code, prop_code))]
    period = st.selectbox("Period (first seen)", ["All"] + periods)
with f3:
    manual_only = st.checkbox("Manual matches only")

# ── Reconciliation report download ────────────────────────────────────────────
_latest_run = db.latest_run(con, prop_code)
if _latest_run:
    _run_id   = _latest_run["run_id"]
    _cache_key = f"report_{prop_code}_{_run_id}"

    _dl_col, _ = st.columns([2, 4])
    with _dl_col:
        if st.button("📊 Generate Reconciliation Report", key="gen_report_btn",
                     help="Rebuild the summary Excel for the latest run, "
                          "reflecting any manual matches made since."):
            with st.spinner("Building report…"):
                try:
                    _bytes, _fname = regenerate.rebuild_output(con, _run_id)
                    st.session_state["report_bytes"]    = _bytes
                    st.session_state["report_filename"] = _fname
                    st.session_state["report_key"]      = _cache_key
                except Exception as _exc:
                    st.error(f"Could not generate report: {_exc}")

        if st.session_state.get("report_key") == _cache_key and \
                "report_bytes" in st.session_state:
            st.download_button(
                "⬇️ Download Report",
                data=st.session_state["report_bytes"],
                file_name=st.session_state["report_filename"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_report_btn",
            )
else:
    st.caption("No completed run found for this property.")

rows = db.matched_view(con, prop_code, None if period == "All" else period, manual_only)
if not rows:
    st.info("No matched transactions for this selection.")
else:
    df = pd.DataFrame([{
        "Match #": r["match_id"], "Rule": r["match_rule"],
        "Manual": "✓" if r["is_manual"] else "",
        "Side": r["side"].upper(), "Date": r["date"],
        "Amount": r["amount_cents"] / 100.0, "Ref": r["ref"] or "",
        "Description": r["description"] or "", "Remarks": r["remarks"] or "",
        "Period": r["source_period"],
    } for r in rows])
    n_matches = df["Match #"].nunique()
    st.caption(f"{n_matches} match group(s), {len(df)} transaction rows. "
               "Rows sharing a Match # were reconciled together.")
    st.dataframe(df, hide_index=True, height=480,
                 column_config={"Amount": st.column_config.NumberColumn(format="%.2f")})
    import io as _io
    _buf = _io.BytesIO()
    df.to_excel(_buf, index=False, engine="openpyxl")
    st.download_button("⬇️ Export Excel", data=_buf.getvalue(),
                       file_name=f"matched_{prop_code}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.subheader("Run history")
runs = con.execute("""SELECT run_id, period, gl_type, run_at, status, prior_source
                      FROM runs WHERE property_code = ? ORDER BY run_id DESC""",
                   (prop_code,)).fetchall()
if not runs:
    st.caption("No runs recorded.")
else:
    rdf = pd.DataFrame([dict(r) for r in runs])
    st.dataframe(rdf, hide_index=True)

    with st.expander("Cancel a run"):
        st.warning(
            "Cancelling a run permanently removes all transactions and matches "
            "it introduced. Prior-period carryover rows matched by the run are "
            "reverted to unmatched. This cannot be undone.")

        run_labels = {
            f"#{r['run_id']} — {r['period']}  [{r['status']}]  {r['run_at'][:16]}": r
            for r in runs
        }
        chosen_label = st.selectbox("Run to cancel", list(run_labels.keys()),
                                    key="cancel_run_select")
        chosen = run_labels[chosen_label]

        n_manual = persist.manual_matches_at_risk(con, prop_code, chosen["period"]) \
            if chosen["status"] != "superseded" else 0
        if n_manual:
            st.error(f"This run has **{n_manual} manual match(es)** that will also be deleted.")

        confirmed = st.checkbox("I understand this is permanent", key="cancel_run_confirm")
        if st.button("Cancel run", type="primary", disabled=not confirmed,
                     key="cancel_run_btn"):
            try:
                summary = persist.cancel_run(con, chosen["run_id"])
                st.success(
                    f"Run #{chosen['run_id']} removed — "
                    f"{summary['matches_deleted']} match(es), "
                    f"{summary['bank_deleted']} bank txn(s), "
                    f"{summary['gl_deleted']} GL txn(s) deleted.")
                # Clear any stale session state referencing this run
                for key in ("last_run_id", "last_output_bytes", "last_output_filename",
                            "last_run_sig", "regen_bytes", "regen_filename",
                            "report_bytes", "report_filename", "report_key"):
                    st.session_state.pop(key, None)
                st.rerun()
            except Exception as e:
                st.error(f"Failed to cancel run: {e}")

    with st.expander("Clear all property data"):
        st.warning(
            "This permanently removes **all** runs, matches, and transactions "
            "for this property. Unlike cancelling a single run, this wipes the "
            "entire history and cannot be undone.")
        delete_prop_cfg = st.checkbox(
            "Also delete the property configuration",
            key="clear_prop_delete_config",
            help="If checked, the property will be removed entirely and must be "
                 "re-configured before running a new reconciliation.")
        confirmed_clear = st.checkbox(
            "I understand this is permanent", key="clear_prop_confirm")
        if st.button("Clear all data", type="primary",
                     disabled=not confirmed_clear, key="clear_prop_btn"):
            try:
                summary = persist.clear_property_data(
                    con, prop_code, delete_property=delete_prop_cfg)
                st.success(
                    f"Cleared {summary['runs_deleted']} run(s), "
                    f"{summary['matches_deleted']} match(es), "
                    f"{summary['bank_deleted']} bank txn(s), "
                    f"{summary['gl_deleted']} GL txn(s).")
                for key in ("last_run_id", "last_output_bytes", "last_output_filename",
                            "last_run_sig", "regen_bytes", "regen_filename",
                            "report_bytes", "report_filename", "report_key"):
                    st.session_state.pop(key, None)
                st.rerun()
            except Exception as e:
                st.error(f"Failed to clear property data: {e}")
