"""Manual matching: pair leftover unreconciled bank and GL transactions."""

from __future__ import annotations

import contextlib
import io

import pandas as pd
import streamlit as st

from core import db, regenerate
from core.auth import require_login

st.set_page_config(page_title="Manual Matching", page_icon="🔗", layout="wide")
require_login()
st.title("🔗 Manual Matching")


con = db.get_or_connect()

props = db.list_properties(con)
if not props:
    st.info("No properties yet — run a reconciliation first.")
    st.stop()

codes = [p.property_code for p in props]
default_idx = codes.index(st.session_state["last_property"]) \
    if st.session_state.get("last_property") in codes else 0
prop_code = st.selectbox("Property", codes, index=default_idx,
                         format_func=lambda c: f"{c} — {next(p.property_name for p in props if p.property_code == c)}")

latest = db.latest_run(con, prop_code)
if latest is None:
    st.warning("No completed run for this property yet. Manual matches can only be "
               "recorded against a run (the output workbook is regenerated from it).")
    st.stop()
st.caption(f"Latest run: #{latest['run_id']} · period {latest['period']} · {latest['run_at']}")

bank_rows = db.unmatched_bank(con, prop_code)
gl_rows   = db.unmatched_gl(con, prop_code)

if not bank_rows and not gl_rows:
    st.success("Nothing left to match — all transactions are reconciled. 🎉")
    st.stop()

bank_df = pd.DataFrame([{
    "Date": r["date"], "Amount": r["amount_cents"] / 100.0,
    "Check #": r["check_number"] or "", "Description": r["description"],
    "Since": r["source_period"], "_hash": r["txn_hash"],
} for r in bank_rows])

gl_df = pd.DataFrame([{
    "Date": r["date"], "Control": r["control"], "Reference": r["reference"],
    "Description": r["description"], "Remarks": r["remarks"],
    "Deposit #": r["deposit_number"] or "",
    "Amount": (r["debit_cents"] - r["credit_cents"]) / 100.0,
    "Since": r["source_period"], "_hash": r["txn_hash"],
} for r in gl_rows])

# ── Search / filter controls ──────────────────────────────────────────────────
with st.expander("🔍 Filter GL rows", expanded=False):
    fa, fb, fc = st.columns(3)
    with fa:
        f_ref  = st.text_input("Reference contains", key="f_ref").strip()
    with fb:
        f_desc = st.text_input("Description / Remarks contains", key="f_desc").strip()
    with fc:
        f_date = st.text_input("Date contains (e.g. 2026-04)", key="f_date").strip()

gl_display = gl_df.copy()
if f_ref:
    gl_display = gl_display[gl_display["Reference"].str.contains(f_ref, case=False, na=False)]
if f_desc:
    mask = (gl_display["Description"].str.contains(f_desc, case=False, na=False) |
            gl_display["Remarks"].str.contains(f_desc, case=False, na=False))
    gl_display = gl_display[mask]
if f_date:
    gl_display = gl_display[gl_display["Date"].astype(str).str.contains(f_date, na=False)]

# ── Tables ────────────────────────────────────────────────────────────────────
# Table key suffix increments after each confirmed match so Streamlit discards
# stale selection state and re-renders the tables from fresh data.
_tbl_ver = st.session_state.get("match_table_ver", 0)

c1, c2 = st.columns(2)
with c1:
    st.subheader(f"Unreconciled bank ({len(bank_df)})")
    bank_sel = st.dataframe(
        bank_df.drop(columns=["_hash"]), hide_index=True, height=420,
        on_select="rerun", selection_mode="multi-row",
        key=f"bank_table_{_tbl_ver}",
        column_config={"Amount": st.column_config.NumberColumn(format="%.2f")})
with c2:
    filtered_label = f" — {len(gl_display)} shown" if (f_ref or f_desc or f_date) else ""
    st.subheader(f"Unreconciled GL ({len(gl_df)}){filtered_label}")
    gl_sel = st.dataframe(
        gl_display.drop(columns=["_hash"]), hide_index=True, height=420,
        on_select="rerun", selection_mode="multi-row",
        key=f"gl_table_{_tbl_ver}",
        column_config={"Amount": st.column_config.NumberColumn(format="%.2f")})

# Clamp selection indices to current dataframe bounds to prevent IndexError
# after a rerun where matched rows are removed and the table is shorter.
raw_bank_idx = bank_sel.selection.rows if bank_sel and bank_sel.selection else []
raw_gl_idx   = gl_sel.selection.rows   if gl_sel   and gl_sel.selection   else []
bank_idx = [i for i in raw_bank_idx if i < len(bank_df)]
gl_idx   = [i for i in raw_gl_idx   if i < len(gl_display)]

bank_total = round(bank_df.iloc[bank_idx]["Amount"].sum(), 2) if bank_idx else 0.0
gl_total   = round(gl_display.iloc[gl_idx]["Amount"].sum(), 2) if gl_idx else 0.0
delta = round(bank_total - gl_total, 2)

t1, t2, t3 = st.columns(3)
t1.metric(f"Bank selected ({len(bank_idx)})", f"{bank_total:,.2f}")
t2.metric(f"GL selected ({len(gl_idx)})", f"{gl_total:,.2f}")
t3.metric("Difference", f"{delta:,.2f}")

if not bank_idx or not gl_idx:
    st.info("Select at least one bank row and one GL row to record a match.")
    st.stop()

balanced = abs(delta) < 0.005
force = False
if balanced:
    st.success("Amounts balance.")
else:
    st.warning(f"Amounts do **not** balance (difference {delta:,.2f}). "
               "Matching anyway will leave a variance in the output workbook.")
    force = st.checkbox("Match anyway despite the difference")

if st.button("Confirm match", type="primary", disabled=not (balanced or force),
             use_container_width=True):
    bank_hashes = bank_df.iloc[bank_idx]["_hash"].tolist()
    gl_hashes   = gl_display.iloc[gl_idx]["_hash"].tolist()
    match_id = db.create_manual_match(con, prop_code, bank_hashes, gl_hashes,
                                      latest["run_id"])
    # Bump table key so matched rows are removed from the UI on rerun
    st.session_state["match_table_ver"] = _tbl_ver + 1
    log = io.StringIO()
    try:
        with contextlib.redirect_stdout(log):
            regen_bytes, regen_filename = regenerate.rebuild_output(con, latest["run_id"])
        st.session_state["regen_bytes"]    = regen_bytes
        st.session_state["regen_filename"] = regen_filename
        st.success(f"Match #{match_id} recorded ({len(bank_hashes)} bank ↔ "
                   f"{len(gl_hashes)} GL). Output workbook regenerated.")
    except Exception as e:
        st.error(f"Match #{match_id} saved, but regenerating the workbook failed: {e}")
    st.rerun()

if st.session_state.get("regen_bytes"):
    st.download_button("⬇️ Download updated workbook",
                       data=st.session_state["regen_bytes"],
                       file_name=st.session_state.get("regen_filename", "reconciliation.xlsx"),
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       use_container_width=True)
