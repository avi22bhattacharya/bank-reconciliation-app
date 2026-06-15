"""Bank Reconciliation Dashboard — Run page.

Launch:  .env/bin/streamlit run app.py
"""

from __future__ import annotations

import contextlib
import hashlib
import io
import tempfile
import traceback
import uuid
from datetime import datetime
from pathlib import Path

import streamlit as st

from core import db, ingest, persist, prior_items, storage
from core import pipeline_ma, pipeline_yardi
from core.auth import require_login

ROOT = Path(__file__).resolve().parent
# Staging dir for uploaded files within a session (ephemeral local disk is fine)
STAGING = Path(tempfile.gettempdir()) / "bank_rec_staging"
STAGING.mkdir(exist_ok=True)

st.set_page_config(page_title="Bank Reconciliation", page_icon="🏦", layout="wide")
require_login()
st.title("🏦 Bank Reconciliation — Run")


def get_con():
    if "db_con" not in st.session_state:
        st.session_state["db_con"] = db.connect()
    return st.session_state["db_con"]


con = get_con()


def stage_upload(uploaded) -> Path:
    """Write an uploaded file to the session staging dir; return its path."""
    digest = hashlib.sha1(uploaded.getbuffer()).hexdigest()[:12]
    d = STAGING / digest
    d.mkdir(parents=True, exist_ok=True)
    p = d / uploaded.name
    if not p.exists():
        p.write_bytes(uploaded.getbuffer())
    return p


@st.cache_data(show_spinner=False)
def cached_detection(path_str: str):
    names = ingest.sheet_names(path_str)
    gl_sheet = ingest.detect_gl_sheet(path_str)
    bank_sheet = ingest.detect_bank_sheet(path_str)
    meta = ingest.detect_gl_metadata(path_str, gl_sheet) if gl_sheet else None
    return names, gl_sheet, bank_sheet, meta


def prev_period(period: str) -> str:
    y, m = (int(x) for x in period.split("-"))
    y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return f"{y:04d}-{m:02d}"


# ── Step 1: inputs ───────────────────────────────────────────────────────────
st.subheader("1 · Inputs")

gl_type_label = st.radio("GL type", ["Yardi", "Manage America"], horizontal=True)
gl_type = "yardi" if gl_type_label == "Yardi" else "ma"

c1, c2 = st.columns(2)
with c1:
    gl_up = st.file_uploader("General Ledger workbook (.xls / .xlsx)",
                             type=["xls", "xlsx"], key="gl_up")
    same_wb = st.checkbox("Bank statement is in the same workbook", value=True)
    bank_up = None
    if not same_wb:
        bank_up = st.file_uploader("Bank statement workbook (.xls / .xlsx)",
                                   type=["xls", "xlsx"], key="bank_up")
with c2:
    dr_up = None
    if gl_type == "yardi":
        dr_up = st.file_uploader("Deposit Register (.xlsx) — required for Yardi",
                                 type=["xls", "xlsx"], key="dr_up")
    boot_up = st.file_uploader(
        "Prior Open Items (optional)", type=["xlsx", "xls"], key="boot_up",
        help="Used only when this property has no history in the database yet "
             "(first run). Afterwards prior unreconciled items come from the DB.")

if gl_up is None:
    st.info("Upload the GL workbook to begin.")
    st.stop()

gl_path = stage_upload(gl_up)
bank_path = stage_upload(bank_up) if bank_up else gl_path

# ── Detection + confirmation ─────────────────────────────────────────────────
st.subheader("2 · Detected details")

try:
    gl_names, gl_sheet_guess, _, meta = cached_detection(str(gl_path))
    bank_names, _, bank_sheet_guess, _ = cached_detection(str(bank_path))
except Exception as e:
    st.error(f"Could not read workbook: {e}")
    st.stop()

if meta and meta.gl_type_guess and meta.gl_type_guess != gl_type:
    other = "Yardi" if meta.gl_type_guess == "yardi" else "Manage America"
    st.warning(f"The GL metadata looks like a **{other}** export, but "
               f"**{gl_type_label}** is selected. Double-check before running.")

d1, d2, d3, d4 = st.columns(4)
with d1:
    property_code = st.text_input("Property code", value=meta.property_code if meta else "")
sub_codes = property_code.split("^") if "^" in property_code else []
is_compound = len(sub_codes) > 1
with d2:
    default_name = ""
    if meta:
        default_name = meta.property_name
    existing = db.get_property(con, property_code) if property_code else None
    sub_existing = [db.get_property(con, c) for c in sub_codes] if is_compound else []
    if not default_name:
        if existing:
            default_name = existing.property_name
        elif sub_existing:
            first_ex = next((e for e in sub_existing if e), None)
            if first_ex:
                default_name = first_ex.property_name
    property_name = st.text_input("Property name", value=default_name)
with d3:
    period = st.text_input("Period (YYYY-MM)", value=meta.period if meta else "")
with d4:
    st.text_input("GL account #", value=meta.gl_account_number if meta else "",
                  disabled=True)

s1, s2 = st.columns(2)
with s1:
    gl_sheet = st.selectbox("GL sheet", gl_names,
                            index=gl_names.index(gl_sheet_guess) if gl_sheet_guess in gl_names else 0)
with s2:
    bank_sheet = st.selectbox("Bank statement sheet", bank_names,
                              index=bank_names.index(bank_sheet_guess) if bank_sheet_guess in bank_names else 0)

_expander_open = (existing is None) if not is_compound else any(e is None for e in sub_existing)
with st.expander("Property settings (used in the output workbook)", expanded=_expander_open):
    if is_compound:
        sub_settings = []
        for tab, code, ex in zip(st.tabs(sub_codes), sub_codes, sub_existing):
            with tab:
                pm1, pm2 = st.columns(2)
                with pm1:
                    ai = st.text_input(
                        "Bank account label", value=ex.account_info if ex else "",
                        placeholder="e.g. WFB Account number 4573468451",
                        key=f"account_info_{code}")
                    gi = st.text_input(
                        "GL account label",
                        value=ex.gl_account_id if ex
                        else (f"{meta.gl_account_number} {code}" if meta and meta.gl_account_number else ""),
                        placeholder="e.g. 11100113 PH Davie - Ops - WFB   (Bank Account)",
                        key=f"gl_account_id_{code}")
                with pm2:
                    plb = st.text_input(
                        "Bank-row property label (Yardi output)",
                        value=ex.prop_label_bank if ex else code,
                        key=f"prop_label_bank_{code}")
                    pdt = st.text_input(
                        "Property display title",
                        value=ex.prop_display if ex else code,
                        placeholder="e.g. BAY CITY MHC, LLC",
                        key=f"prop_display_{code}")
                sub_settings.append({"code": code, "account_info": ai,
                                     "gl_account_id": gi, "prop_label_bank": plb,
                                     "prop_display": pdt})
        account_info = sub_settings[0]["account_info"]
        gl_account_id = sub_settings[0]["gl_account_id"]
        prop_label_bank = sub_settings[0]["prop_label_bank"]
        prop_display = sub_settings[0]["prop_display"]
    else:
        sub_settings = []
        pm1, pm2 = st.columns(2)
        with pm1:
            account_info = st.text_input(
                "Bank account label", value=existing.account_info if existing else "",
                placeholder="e.g. WFB Account number 4573468451")
            gl_account_id = st.text_input(
                "GL account label",
                value=existing.gl_account_id if existing
                else (f"{meta.gl_account_number} {property_name}" if meta and meta.gl_account_number else ""),
                placeholder="e.g. 11100113 PH Davie - Ops - WFB   (Bank Account)")
        with pm2:
            prop_label_bank = st.text_input(
                "Bank-row property label (Yardi output)",
                value=existing.prop_label_bank if existing else property_name)
            prop_display = st.text_input(
                "Property display title",
                value=existing.prop_display if existing else property_name,
                placeholder="e.g. BAY CITY MHC, LLC")

# ── Run ──────────────────────────────────────────────────────────────────────
st.subheader("3 · Run reconciliation")

problems = []
if not property_code:
    problems.append("Property code is required (not detected — enter it manually).")
if not period or len(period) != 7:
    problems.append("Period must be YYYY-MM.")
if gl_type == "yardi" and dr_up is None:
    problems.append("Deposit Register is required for Yardi.")
for p in problems:
    st.error(p)

n_at_risk = (persist.manual_matches_at_risk(con, property_code, period)
             if property_code and period and len(period) == 7 else 0)
if n_at_risk:
    st.warning(f"Re-running {period} for {property_code} will discard "
               f"**{n_at_risk} manual match(es)** recorded for that month.")

has_history = bool(property_code) and con.execute(
    "SELECT COUNT(*) AS count FROM gl_txns WHERE property_code = ?",
    (property_code,)).fetchone()["count"] > 0
if property_code:
    if has_history:
        st.caption(f"Prior unreconciled items for **{property_code}** will be loaded from the database.")
    elif boot_up is not None:
        st.caption("No history for this property — the uploaded Prior Open Items file will seed the database.")
    else:
        st.caption("No history for this property and no Prior Open Items file — the run starts with no carryover items.")

run_clicked = st.button("Run reconciliation", type="primary",
                        disabled=bool(problems), use_container_width=True)

run_sig = hashlib.sha1("|".join([
    str(gl_path), str(bank_path), gl_type, property_code, period,
    gl_sheet or "", bank_sheet or "",
    stage_upload(dr_up).name if dr_up else "",
]).encode()).hexdigest()

if run_clicked:
    if st.session_state.get("last_run_sig") == run_sig:
        st.info("This exact run already completed below. Change an input to run again, "
                "or click again to re-run anyway.")
        st.session_state["last_run_sig"] = None
        st.stop()

    workdir = Path(tempfile.mkdtemp(prefix="bank_rec_"))

    prop = db.PropertyMeta(
        property_code=property_code, property_name=property_name, gl_type=gl_type,
        account_info=account_info, gl_account_id=gl_account_id,
        prop_label_bank=prop_label_bank, prop_display=prop_display)
    db.upsert_property(con, prop)
    for s in sub_settings:
        db.upsert_property(con, db.PropertyMeta(
            property_code=s["code"], property_name=property_name, gl_type=gl_type,
            account_info=s["account_info"], gl_account_id=s["gl_account_id"],
            prop_label_bank=s["prop_label_bank"], prop_display=s["prop_display"]))

    log = io.StringIO()
    try:
        with st.status("Running reconciliation…", expanded=False) as status:
            status.update(label="Loading prior unreconciled items…")
            if not has_history and boot_up is not None:
                boot_path = stage_upload(boot_up)
                boot = prior_items.parse_bootstrap(boot_path)
                n_seeded = prior_items.seed_db(con, property_code, boot, prev_period(period))
                log.write(f"Bootstrap: seeded {n_seeded} prior items from {boot_up.name}\n")
                prior_source = "bootstrap"
            else:
                prior_source = "db" if has_history else "none"
            prior = prior_items.fetch(con, property_code)
            log.write(f"Prior items: {len(prior.bank)} bank, {len(prior.gl)} GL\n")

            status.update(label="Running matching engine…")
            with contextlib.redirect_stdout(log):
                if gl_type == "yardi":
                    out = pipeline_yardi.run_pipeline(
                        workdir, gl_path=gl_path, gl_sheet=gl_sheet,
                        bank_path=bank_path, bank_sheet=bank_sheet,
                        deposit_register_path=stage_upload(dr_up),
                        prop=prop, period=period, prior=prior)
                else:
                    out = pipeline_ma.run_pipeline(
                        workdir, gl_path=gl_path, gl_sheet=gl_sheet,
                        bank_path=bank_path, bank_sheet=bank_sheet,
                        prop=prop, period=period, prior=prior)

            status.update(label="Saving results…")
            run_uuid = uuid.uuid4().hex
            output_filename = Path(out["output_path"]).name
            ext = "json" if gl_type == "yardi" else "pkl"

            results_key = storage.upload(
                out["results_path"],
                f"{property_code}/{period}/{run_uuid}/results.{ext}")
            output_key = storage.upload(
                out["output_path"],
                f"{property_code}/{period}/{run_uuid}/{output_filename}")

            run_id = persist.save_run(
                con, property_code=property_code, gl_type=gl_type, period=period,
                results=out["results"], prior_source=prior_source,
                prior_refs=out["prior_refs"], workdir="",
                results_path=results_key, output_path=output_key,
                stats=out["stats"])
            status.update(label="Reconciliation complete", state="complete")

        st.session_state["last_run_sig"] = run_sig
        st.session_state["last_run_id"] = run_id
        st.session_state["last_property"] = property_code
    except Exception:
        st.error("Run failed — details below.")
        st.code(log.getvalue() + "\n" + traceback.format_exc())
        st.stop()

    st.session_state["last_log"] = log.getvalue()

# ── Results panel ────────────────────────────────────────────────────────────
run_id = st.session_state.get("last_run_id")
if run_id:
    run = con.execute("SELECT * FROM runs WHERE run_id = ?", (run_id,)).fetchone()
    if run and run["status"] == "complete":
        st.divider()
        st.subheader("Results")
        stats = db.run_stats(run)
        m1, m2, m3, m4 = st.columns(4)
        if run["gl_type"] == "yardi":
            m1.metric("Bank matched", f"{stats.get('n_matched_bank','—')}/{stats.get('total_real_bank','—')}")
            m2.metric("Bank match rate", f"{stats.get('pct_bank','—')}%")
            m3.metric("GL matched", f"{stats.get('n_matched_gl','—')}/{stats.get('total_gl','—')}")
            m4.metric("GL match rate", f"{stats.get('pct_gl','—')}%")
        else:
            m1.metric("Bank matched", f"{stats.get('total_matched_bank','—')}/{stats.get('total_bank_regular','—')}")
            m2.metric("Bank match rate", f"{stats.get('pct_bank',0):.1f}%")
            m3.metric("GL matched", f"{stats.get('total_matched_gl','—')}/{stats.get('total_gl_all','—')}")
            m4.metric("GL match rate", f"{stats.get('pct_gl',0):.1f}%")

        try:
            output_data = storage.read_bytes(run["output_path"])
            st.download_button("⬇️ Download reconciliation workbook",
                               data=output_data,
                               file_name=Path(run["output_path"]).name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        except Exception:
            st.warning("Output file not available for download.")

        ub = len(db.unmatched_bank(con, run["property_code"]))
        ug = len(db.unmatched_gl(con, run["property_code"]))
        st.caption(f"{ub} bank and {ug} GL transactions remain unreconciled — "
                   f"use **Manual Matching** in the sidebar to resolve them.")
        if st.session_state.get("last_log"):
            with st.expander("Engine log"):
                st.code(st.session_state["last_log"])
