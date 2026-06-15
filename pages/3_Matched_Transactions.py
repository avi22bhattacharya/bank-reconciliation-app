"""View reconciled (matched) transactions and run history."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core import db
from core.auth import require_login

st.set_page_config(page_title="Matched Transactions", page_icon="✅", layout="wide")
require_login()
st.title("✅ Matched Transactions")


def get_con():
    if "db_con" not in st.session_state:
        st.session_state["db_con"] = db.connect()
    return st.session_state["db_con"]


con = get_con()

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
runs = con.execute("""SELECT run_id, period, gl_type, run_at, status, prior_source, output_path
                      FROM runs WHERE property_code = ? ORDER BY run_id DESC""",
                   (prop_code,)).fetchall()
if runs:
    rdf = pd.DataFrame([dict(r) for r in runs])
    st.dataframe(rdf, hide_index=True)
else:
    st.caption("No runs recorded.")
