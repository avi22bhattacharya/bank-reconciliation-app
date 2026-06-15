"""Headless end-to-end test of the Streamlit app via streamlit.testing.AppTest.

File uploads are injected by patching st.file_uploader (AppTest cannot set
uploader values); everything else — detection, property form, run button,
pipeline, persistence, results panel — is the real app code path.
"""
import io
import sys
from pathlib import Path

import streamlit as st
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
TESTDB = ROOT / "runs" / "test_app_e2e.db"
TESTDB.unlink(missing_ok=True)

from core import db
db.DB_PATH = TESTDB  # the app's db.connect() default


class FakeUpload(io.BytesIO):
    def __init__(self, path: Path):
        super().__init__(path.read_bytes())
        self.name = path.name


UPLOADS = {}


def fake_uploader(label, *a, key=None, **k):
    return UPLOADS.get(key)


st.file_uploader = fake_uploader
# also patch the alias AppTest scripts resolve through module attribute access
import streamlit as _st_mod
_st_mod.file_uploader = fake_uploader


def run_app(timeout=900):
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=timeout)
    return at


def values(widgets):
    return [w.value for w in widgets]


# ───────────────────────── MA flow ─────────────────────────
MA = ROOT / "Manage America Bank recon"
UPLOADS.clear()
UPLOADS["gl_up"] = FakeUpload(MA / "03 2026 Bank Recon Data_Bay City_161bct_MA.xls")
UPLOADS["boot_up"] = FakeUpload(MA / "02 2026 Open Items_Bay City_161bct_MA.xlsx")

at = run_app()
at.run()
assert not at.exception, at.exception
# choose Manage America
at.radio[0].set_value("Manage America").run()
assert not at.exception, at.exception

# detection populated?
ti = {t.label: t.value for t in at.text_input}
assert ti["Property code"] == "161bct", ti
assert ti["Period (YYYY-MM)"] == "2026-03", ti
print("MA detection OK:", ti["Property code"], ti["Period (YYYY-MM)"])
sb = {s.label: s.value for s in at.selectbox}
assert sb["Bank statement sheet"] == "DepositAccount_1680_040126_6000", sb
assert sb["GL sheet"] == "GL", sb

# fill property settings
for t in at.text_input:
    if t.label == "Bank account label":
        t.set_value("WFB Account number 4076381680")
    elif t.label == "GL account label":
        t.set_value("11100161 BAY CITY MHC, LLC")
    elif t.label == "Property display title":
        t.set_value("BAY CITY MHC, LLC")
at.run()

# run it
btn = next(b for b in at.button if b.label == "Run reconciliation")
assert not btn.disabled, "run button disabled: " + str([e.value for e in at.error])
btn.click().run()
assert not at.exception, at.exception
assert not at.error, [e.value for e in at.error]

con = db.connect(TESTDB)
run = con.execute("SELECT * FROM runs ORDER BY run_id DESC LIMIT 1").fetchone()
assert run and run["status"] == "complete" and run["gl_type"] == "ma", dict(run or {})
assert Path(run["output_path"]).exists()
mets = {m.label: m.value for m in at.metric}
print("MA run OK — metrics:", mets)
assert "Bank matched" in mets
n_match = con.execute("SELECT COUNT(*) FROM matches").fetchone()[0]
print(f"MA: matches persisted = {n_match}")
assert n_match > 0

# ───────────────────────── Yardi flow ─────────────────────────
Y = ROOT / "deposit-register-trial"
UPLOADS.clear()
UPLOADS["gl_up"] = FakeUpload(Y / "03 2026 Bank Recon Data_PH Davie_113phv^114epr_MH.xls")
UPLOADS["dr_up"] = FakeUpload(Y / "Deposit_Register_May_PH davie MH.xlsx")
UPLOADS["boot_up"] = FakeUpload(Y / "02 2026_Open Items_PH Davie_113phv^114epr_MH.xlsx")

at = run_app()
at.run()  # default radio = Yardi
assert not at.exception, at.exception
ti = {t.label: t.value for t in at.text_input}
assert ti["Property code"] == "113phv^114epr", ti
assert ti["Period (YYYY-MM)"] == "2026-03", ti
print("Yardi detection OK:", ti["Property code"], ti["Period (YYYY-MM)"])

for t in at.text_input:
    if t.label == "Bank account label":
        t.set_value("WFB Account number 4573468451")
    elif t.label == "GL account label":
        t.set_value("11100113 PH Davie - Ops - WFB   (Bank Account)")
    elif t.label == "Property name":
        t.set_value("PH Davie")
    elif t.label == "Bank-row property label (Yardi output)":
        t.set_value("Palm Haven^East Pine Ridge")
    elif t.label == "Property display title":
        t.set_value("Property: East Pine Ridge(114epr), Palm Haven (113phv)")
at.run()
btn = next(b for b in at.button if b.label == "Run reconciliation")
assert not btn.disabled, "run button disabled: " + str([e.value for e in at.error])
btn.click().run()
assert not at.exception, at.exception
assert not at.error, [e.value for e in at.error]

run = con.execute("SELECT * FROM runs WHERE gl_type='yardi' ORDER BY run_id DESC LIMIT 1").fetchone()
assert run and run["status"] == "complete", dict(run or {})
assert Path(run["output_path"]).exists()
mets = {m.label: m.value for m in at.metric}
print("Yardi run OK — metrics:", mets)
assert mets.get("Bank match rate") == "91.3%", mets   # matches the golden stats

# ───────────────────────── pages render ─────────────────────────
for page in ["pages/2_Manual_Matching.py", "pages/3_Matched_Transactions.py"]:
    atp = AppTest.from_file(str(ROOT / page), default_timeout=120)
    atp.run()
    assert not atp.exception, (page, atp.exception)
    print(f"{page} renders OK "
          f"(dataframes: {len(atp.dataframe) if hasattr(atp, 'dataframe') else '?'})")

print("\nALL APP E2E CHECKS PASSED")
