"""Post-run persistence: write reconciliation results to PostgreSQL / SQLite.

Within one transaction per run:
  1. Supersede earlier runs of the same (property, period): revert their
     matches, delete the period's transaction rows, mark runs superseded.
  2. Batch-INSERT every engine record by content hash (ON CONFLICT DO NOTHING).
  3. Batch-UPDATE statuses: 'internal', 'matched', 'unmatched'.
  4. INSERT match groups (individual — needs RETURNING match_id).
  5. Batch-UPDATE match_id back onto member rows.

Steps 2, 3, 5 each make ONE round-trip to the database regardless of row
count, using psycopg2.extras.execute_values on PostgreSQL.  This replaces
the previous per-row approach that made ~8 000 round-trips for a property
with 4 000 transactions — causing apparent hangs on remote Supabase.
"""

from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime

from core import db
from core.prior_items import clean_check
from engines.ma.build_excel import assign_codes as ma_assign_codes

_YARDI_INTERNAL_BANK = {"INTERNAL – Stagecoach Sweep", "Contra - Bank"}

_BANK_INSERT_SQL = """
    INSERT INTO bank_txns
      (txn_hash, property_code, date, amount_cents, check_number, description,
       status, source_period, first_seen_run_id)
    VALUES (?,?,?,?,?,?,?,?,?)
    ON CONFLICT (txn_hash) DO NOTHING
"""

_GL_INSERT_SQL = """
    INSERT INTO gl_txns
      (txn_hash, property_code, property_label, date, control, reference,
       description, remarks, debit_cents, credit_cents, deposit_number,
       tenant_code, status, source_period, first_seen_run_id)
    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT (txn_hash) DO NOTHING
"""


def manual_matches_at_risk(con, property_code: str, period: str) -> int:
    """Manual matches that a re-run of this (property, period) would discard."""
    row = con.execute("""
        SELECT COUNT(*) AS count FROM matches m
        JOIN runs r ON r.run_id = m.run_id
        WHERE m.is_manual = 1 AND r.property_code = ? AND r.period = ?
          AND r.status != 'superseded'
    """, (property_code, period)).fetchone()
    return row["count"]


def _supersede(con, property_code: str, period: str):
    run_ids = [r["run_id"] for r in con.execute(
        "SELECT run_id FROM runs WHERE property_code = ? AND period = ? AND status != 'superseded'",
        (property_code, period))]
    if not run_ids:
        return
    ph = ",".join("?" * len(run_ids))
    match_ids = [r["match_id"] for r in con.execute(
        f"SELECT match_id FROM matches WHERE run_id IN ({ph})", run_ids)]
    if match_ids:
        mph = ",".join("?" * len(match_ids))
        for table in ("bank_txns", "gl_txns"):
            con.execute(f"""UPDATE {table} SET status='unmatched', match_id=NULL,
                            matched_run_id=NULL WHERE match_id IN ({mph})""", match_ids)
        con.execute(f"DELETE FROM matches WHERE match_id IN ({mph})", match_ids)
    con.execute("DELETE FROM bank_txns WHERE property_code = ? AND source_period = ?",
                (property_code, period))
    con.execute("DELETE FROM gl_txns WHERE property_code = ? AND source_period = ?",
                (property_code, period))
    con.execute(f"""UPDATE bank_txns SET status='unmatched', match_id=NULL, matched_run_id=NULL
                    WHERE matched_run_id IN ({ph})""", run_ids)
    con.execute(f"""UPDATE gl_txns SET status='unmatched', match_id=NULL, matched_run_id=NULL
                    WHERE matched_run_id IN ({ph})""", run_ids)
    con.execute(f"UPDATE runs SET status='superseded' WHERE run_id IN ({ph})", run_ids)


# ── engine-specific result walkers ───────────────────────────────────────────

def _persist_yardi(con, property_code, period, run_id, results, prior_refs):
    occ_b, occ_g = defaultdict(int), defaultdict(int)
    bank_hash_by_id: dict[str, str] = {}
    gl_hash_by_id:   dict[str, str] = {}

    bank_inserts: list[tuple] = []   # rows for batch INSERT
    bank_states:  list[tuple] = []   # (status, engine_ref, matched_run, hash)
    gl_inserts:   list[tuple] = []
    gl_states:    list[tuple] = []

    for rec in results["all_bank"]:
        rule = rec.get("match_rule") or ""
        if rule in _YARDI_INTERNAL_BANK:
            status = "internal"
        elif rec.get("matched"):
            status = "matched"
        else:
            status = "unmatched"
        matched_run = run_id if status in ("matched", "internal") else None
        engine_ref  = rec["id"]

        known = prior_refs.get(rec["id"])
        if known:
            h = known
        else:
            cents = db.to_cents(rec["amount"])
            chk   = clean_check(rec.get("check_number"))
            key   = db.bank_content_key(property_code, rec.get("date"), cents, chk,
                                        rec.get("description"))
            h = db.bank_hash(property_code, period, occ_b[key], rec.get("date"),
                             cents, chk, rec.get("description"))
            occ_b[key] += 1
            bank_inserts.append((
                h, property_code, db.iso_date(rec.get("date")), cents, chk,
                (rec.get("description") or "").strip(), "unmatched", period, run_id,
            ))
        bank_hash_by_id[rec["id"]] = h
        bank_states.append((status, engine_ref, matched_run, h))

    for rec in results["all_gl"]:
        rule = rec.get("match_rule") or ""
        if rule == "Contra - GL":
            status = "internal"
        elif rec.get("matched"):
            status = "matched"
        else:
            status = "unmatched"
        matched_run = run_id if status in ("matched", "internal") else None
        engine_ref  = rec["id"]

        known = prior_refs.get(rec["id"])
        if known:
            h = known
        else:
            dc  = db.to_cents(rec.get("debit"))
            cc  = db.to_cents(rec.get("credit"))
            key = db.gl_content_key(property_code, rec.get("date"), rec.get("control"),
                                    rec.get("ref"), dc, cc, rec.get("desc"),
                                    rec.get("remarks"))
            h = db.gl_hash(property_code, period, occ_g[key], rec.get("date"),
                           rec.get("control"), rec.get("ref"), dc, cc,
                           rec.get("desc"), rec.get("remarks"))
            occ_g[key] += 1
            gl_inserts.append((
                h, property_code, (rec.get("property") or "").strip(),
                db.iso_date(rec.get("date")), (rec.get("control") or "").strip(),
                (rec.get("ref") or "").strip(), (rec.get("desc") or "").strip(),
                (rec.get("remarks") or "").strip(), dc, cc,
                rec.get("deposit_num") or "", "",
                "unmatched", period, run_id,
            ))
        gl_hash_by_id[rec["id"]] = h
        gl_states.append((status, engine_ref, matched_run, h))

    # ── Batch INSERT + state UPDATE (2 round-trips each) ─────────────────────
    con.bulk_insert(_BANK_INSERT_SQL, bank_inserts)
    con.bulk_insert(_GL_INSERT_SQL,   gl_inserts)
    con.bulk_set_state("bank_txns", bank_states)
    con.bulk_set_state("gl_txns",   gl_states)

    # ── Match groups (union-find → connected components) ─────────────────────
    bank_by_id = {b["id"]: b for b in results["all_bank"]}
    gl_by_id   = {g["id"]: g for g in results["all_gl"]}

    parent: dict = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        parent[find(a)] = find(b)

    for b in results["all_bank"]:
        for gid in (b.get("match_ids") or []):
            union(("b", b["id"]), ("g", gid))
    for g in results["all_gl"]:
        for bid in (g.get("match_ids") or []):
            union(("g", g["id"]), ("b", bid))

    comps: dict = defaultdict(lambda: {"b": [], "g": []})
    for node in list(parent):
        comps[find(node)][node[0]].append(node[1])

    bank_mid_rows: list[tuple] = []   # (match_id, txn_hash)
    gl_mid_rows:   list[tuple] = []

    now = datetime.now().isoformat(timespec="seconds")
    for comp in comps.values():
        if not comp["b"] and not comp["g"]:
            continue
        rule = ""
        for bid in comp["b"]:
            rule = bank_by_id[bid].get("match_rule") or rule
        bank_cents = sum(db.to_cents(bank_by_id[i]["amount"]) for i in comp["b"])
        gl_cents   = sum(db.to_cents(gl_by_id[i].get("debit")   or 0)
                       - db.to_cents(gl_by_id[i].get("credit") or 0)
                       for i in comp["g"])
        cur = con.execute("""
            INSERT INTO matches (property_code, run_id, match_rule, is_manual,
                                 matched_at, bank_total_cents, gl_total_cents)
            VALUES (?,?,?,0,?,?,?) RETURNING match_id
        """, (property_code, run_id, rule, now, bank_cents, gl_cents))
        mid = cur.lastrowid
        for bid in comp["b"]:
            bank_mid_rows.append((mid, bank_hash_by_id[bid]))
        for gid in comp["g"]:
            gl_mid_rows.append((mid, gl_hash_by_id[gid]))

    # ── Batch match_id UPDATE (1 round-trip each) ─────────────────────────────
    con.bulk_set_match_id("bank_txns", bank_mid_rows)
    con.bulk_set_match_id("gl_txns",   gl_mid_rows)


def _persist_ma(con, property_code, period, run_id, results, prior_refs):
    bank_code, gl_code, _ = ma_assign_codes(results)
    occ_b, occ_g = defaultdict(int), defaultdict(int)

    bank_hashes: list[str] = []
    gl_hashes:   list[str] = []
    bank_inserts: list[tuple] = []
    bank_states:  list[tuple] = []
    gl_inserts:   list[tuple] = []
    gl_states:    list[tuple] = []

    for bi, rec in enumerate(results["all_bank"]):
        code   = abs(bank_code[bi])
        status = {1: "unmatched", 2: "matched", 3: "matched", 4: "internal"}[code]
        matched_run = run_id if status in ("matched", "internal") else None
        engine_ref  = f"B{bi}"

        known = prior_refs.get(f"B{bi}")
        if known:
            h = known
        else:
            cents = db.to_cents(rec["amount"])
            chk   = clean_check(rec.get("check_number"))
            key   = db.bank_content_key(property_code, rec.get("date"), cents, chk,
                                        rec.get("description"))
            h = db.bank_hash(property_code, period, occ_b[key], rec.get("date"),
                             cents, chk, rec.get("description"))
            occ_b[key] += 1
            bank_inserts.append((
                h, property_code, db.iso_date(rec.get("date")), cents, chk,
                (rec.get("description") or "").strip(), "unmatched", period, run_id,
            ))
        bank_hashes.append(h)
        bank_states.append((status, engine_ref, matched_run, h))

    for gi, rec in enumerate(results["all_gl"]):
        code   = abs(gl_code[gi])
        status = {1: "unmatched", 2: "matched", 3: "matched", 4: "internal"}[code]
        matched_run = run_id if status in ("matched", "internal") else None
        engine_ref  = f"G{gi}"

        known = prior_refs.get(f"G{gi}")
        if known:
            h = known
        else:
            dc  = db.to_cents(rec.get("debit"))
            cc  = db.to_cents(rec.get("credit"))
            key = db.gl_content_key(property_code, rec.get("date"), rec.get("control"),
                                    rec.get("reference"), dc, cc, rec.get("person_desc"),
                                    rec.get("remarks"))
            h = db.gl_hash(property_code, period, occ_g[key], rec.get("date"),
                           rec.get("control"), rec.get("reference"), dc, cc,
                           rec.get("person_desc"), rec.get("remarks"))
            occ_g[key] += 1
            gl_inserts.append((
                h, property_code, (rec.get("property_label") or "").strip(),
                db.iso_date(rec.get("date")), (rec.get("control") or "").strip(),
                (rec.get("reference") or "").strip(), (rec.get("person_desc") or "").strip(),
                (rec.get("remarks") or "").strip(), dc, cc, "", "",
                "unmatched", period, run_id,
            ))
        gl_hashes.append(h)
        gl_states.append((status, engine_ref, matched_run, h))

    con.bulk_insert(_BANK_INSERT_SQL, bank_inserts)
    con.bulk_insert(_GL_INSERT_SQL,   gl_inserts)
    con.bulk_set_state("bank_txns", bank_states)
    con.bulk_set_state("gl_txns",   gl_states)

    bank_mid_rows: list[tuple] = []
    gl_mid_rows:   list[tuple] = []

    now = datetime.now().isoformat(timespec="seconds")
    for bi, info in results["matched_bank"].items():
        gl_idxs = info.get("gl_indices") or []
        if not gl_idxs:
            continue
        b          = results["all_bank"][bi]
        bank_cents = db.to_cents(b["amount"])
        gl_cents   = sum(db.to_cents(results["all_gl"][gi].get("debit")   or 0)
                       - db.to_cents(results["all_gl"][gi].get("credit") or 0)
                       for gi in gl_idxs)
        cur = con.execute("""
            INSERT INTO matches (property_code, run_id, match_rule, is_manual,
                                 matched_at, bank_total_cents, gl_total_cents)
            VALUES (?,?,?,0,?,?,?) RETURNING match_id
        """, (property_code, run_id, info["rule"], now, bank_cents, gl_cents))
        mid = cur.lastrowid
        bank_mid_rows.append((mid, bank_hashes[bi]))
        for gi in gl_idxs:
            gl_mid_rows.append((mid, gl_hashes[gi]))

    con.bulk_set_match_id("bank_txns", bank_mid_rows)
    con.bulk_set_match_id("gl_txns",   gl_mid_rows)


def save_run(con, *, property_code: str, gl_type: str, period: str,
             results: dict, prior_source: str, prior_refs: dict | None = None,
             workdir: str = "", results_path: str = "", output_path: str = "",
             stats: dict | None = None, results_data: str = "") -> int:
    """Persist one reconciliation run atomically. Returns run_id."""
    prior_refs = prior_refs or {}
    try:
        _supersede(con, property_code, period)
        cur = con.execute("""
            INSERT INTO runs (property_code, gl_type, period, run_at, status,
                              prior_source, workdir, results_path, output_path,
                              stats_json, results_data)
            VALUES (?,?,?,?,?,?,?,?,?,?,?) RETURNING run_id
        """, (property_code, gl_type, period,
              datetime.now().isoformat(timespec="seconds"), "complete", prior_source,
              str(workdir), str(results_path), str(output_path),
              json.dumps(stats or {}), results_data))
        run_id = cur.lastrowid

        if gl_type == "yardi":
            _persist_yardi(con, property_code, period, run_id, results, prior_refs)
        else:
            _persist_ma(con, property_code, period, run_id, results, prior_refs)

        con.commit()
        return run_id
    except Exception:
        con.rollback()
        raise


def cancel_run(con, run_id: int) -> dict:
    """Undo a reconciliation run and remove all its effects from the DB."""
    run = con.execute("SELECT * FROM runs WHERE run_id = ?",
                      (run_id,)).fetchone()
    if run is None:
        raise ValueError(f"Run {run_id} not found")

    summary = {"matches_deleted": 0, "bank_deleted": 0, "gl_deleted": 0}
    try:
        if run["status"] != "superseded":
            property_code = run["property_code"]
            period        = run["period"]

            match_ids = [r["match_id"] for r in con.execute(
                "SELECT match_id FROM matches WHERE run_id = ?", (run_id,))]
            summary["matches_deleted"] = len(match_ids)

            if match_ids:
                mph = ",".join("?" * len(match_ids))
                for table in ("bank_txns", "gl_txns"):
                    con.execute(
                        f"UPDATE {table} SET status='unmatched', match_id=NULL,"
                        f" matched_run_id=NULL WHERE match_id IN ({mph})", match_ids)
                con.execute(f"DELETE FROM matches WHERE match_id IN ({mph})", match_ids)

            con.execute("""UPDATE bank_txns SET status='unmatched', match_id=NULL,
                           matched_run_id=NULL WHERE matched_run_id = ?""", (run_id,))
            con.execute("""UPDATE gl_txns SET status='unmatched', match_id=NULL,
                           matched_run_id=NULL WHERE matched_run_id = ?""", (run_id,))

            cur = con.execute(
                "DELETE FROM bank_txns WHERE property_code = ? AND source_period = ?",
                (property_code, period))
            summary["bank_deleted"] = cur.rowcount
            cur = con.execute(
                "DELETE FROM gl_txns WHERE property_code = ? AND source_period = ?",
                (property_code, period))
            summary["gl_deleted"] = cur.rowcount

        con.execute("DELETE FROM runs WHERE run_id = ?", (run_id,))
        con.commit()
        return summary
    except Exception:
        con.rollback()
        raise


def clear_property_data(con, property_code: str, delete_property: bool = False) -> dict:
    """Remove ALL stored data for a property (all runs, matches, and transactions).

    Optionally also removes the property configuration row.
    Returns a summary dict of how many rows were deleted.
    """
    summary = {"runs_deleted": 0, "matches_deleted": 0,
               "bank_deleted": 0, "gl_deleted": 0}
    try:
        match_rows = con.execute(
            "SELECT match_id FROM matches WHERE property_code = ?",
            (property_code,)).fetchall()
        match_ids = [r["match_id"] for r in match_rows]
        summary["matches_deleted"] = len(match_ids)

        if match_ids:
            mph = ",".join("?" * len(match_ids))
            for table in ("bank_txns", "gl_txns"):
                con.execute(
                    f"UPDATE {table} SET status='unmatched', match_id=NULL,"
                    f" matched_run_id=NULL WHERE match_id IN ({mph})", match_ids)
            con.execute(f"DELETE FROM matches WHERE match_id IN ({mph})", match_ids)

        cur = con.execute("DELETE FROM bank_txns WHERE property_code = ?", (property_code,))
        summary["bank_deleted"] = cur.rowcount
        cur = con.execute("DELETE FROM gl_txns WHERE property_code = ?", (property_code,))
        summary["gl_deleted"] = cur.rowcount
        cur = con.execute("DELETE FROM runs WHERE property_code = ?", (property_code,))
        summary["runs_deleted"] = cur.rowcount

        if delete_property:
            con.execute("DELETE FROM properties WHERE property_code = ?", (property_code,))

        con.commit()
        return summary
    except Exception:
        con.rollback()
        raise
