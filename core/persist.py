"""Post-run persistence: write reconciliation results to SQLite.

Within one transaction per run:
  1. Supersede earlier runs of the same (property, period): revert their
     matches (older-period member rows go back to 'unmatched'), delete the
     period's transaction rows, mark the runs superseded. Manual matches
     created off those runs are discarded too (the UI warns first).
  2. Upsert every engine record by content hash (INSERT OR IGNORE — a
     still-unreconciled carryover lands on its existing row).
  3. Apply statuses: 'internal' (stagecoach/contra), 'matched', 'unmatched'.
     A prior-period row that matched this month flips to 'matched', which is
     exactly the transition out of the unreconciled set.
  4. Insert match groups and point member rows at them.
"""

from __future__ import annotations

import json
from collections import defaultdict
from datetime import datetime

from core import db
from core.prior_items import clean_check
from engines.ma.build_excel import assign_codes as ma_assign_codes

_YARDI_INTERNAL_BANK = {"INTERNAL – Stagecoach Sweep", "Contra - Bank"}


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
    # rows first seen in this period disappear with the superseded run
    con.execute("DELETE FROM bank_txns WHERE property_code = ? AND source_period = ?",
                (property_code, period))
    con.execute("DELETE FROM gl_txns WHERE property_code = ? AND source_period = ?",
                (property_code, period))
    # any remaining rows matched by those runs (older-period carryovers)
    con.execute(f"""UPDATE bank_txns SET status='unmatched', match_id=NULL, matched_run_id=NULL
                    WHERE matched_run_id IN ({ph})""", run_ids)
    con.execute(f"""UPDATE gl_txns SET status='unmatched', match_id=NULL, matched_run_id=NULL
                    WHERE matched_run_id IN ({ph})""", run_ids)
    con.execute(f"UPDATE runs SET status='superseded' WHERE run_id IN ({ph})", run_ids)


def _upsert_bank(con, property_code, period, run_id, rec, occ, known_hash=None) -> str:
    """Insert a current-month bank record (occurrence-indexed hash), or just
    return the carried hash for a prior-month record (its row already exists)."""
    if known_hash:
        return known_hash
    cents = db.to_cents(rec["amount"])
    chk = clean_check(rec.get("check_number"))
    key = db.bank_content_key(property_code, rec.get("date"), cents, chk,
                              rec.get("description"))
    h = db.bank_hash(property_code, period, occ[key], rec.get("date"), cents, chk,
                     rec.get("description"))
    occ[key] += 1
    con.execute("""
        INSERT INTO bank_txns
          (txn_hash, property_code, date, amount_cents, check_number, description,
           status, source_period, first_seen_run_id)
        VALUES (?,?,?,?,?,?,'unmatched',?,?)
        ON CONFLICT (txn_hash) DO NOTHING
    """, (h, property_code, db.iso_date(rec.get("date")), cents, chk,
          (rec.get("description") or "").strip(), period, run_id))
    return h


def _upsert_gl(con, property_code, period, run_id, rec, occ, *, desc_key, ref_key,
               known_hash=None) -> str:
    if known_hash:
        return known_hash
    dc, cc = db.to_cents(rec.get("debit")), db.to_cents(rec.get("credit"))
    key = db.gl_content_key(property_code, rec.get("date"), rec.get("control"),
                            rec.get(ref_key), dc, cc, rec.get(desc_key),
                            rec.get("remarks"))
    h = db.gl_hash(property_code, period, occ[key], rec.get("date"), rec.get("control"),
                   rec.get(ref_key), dc, cc, rec.get(desc_key), rec.get("remarks"))
    occ[key] += 1
    con.execute("""
        INSERT INTO gl_txns
          (txn_hash, property_code, property_label, date, control, reference,
           description, remarks, debit_cents, credit_cents, deposit_number,
           tenant_code, status, source_period, first_seen_run_id)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,'unmatched',?,?)
        ON CONFLICT (txn_hash) DO NOTHING
    """, (h, property_code, (rec.get("property") or rec.get("property_label") or "").strip(),
          db.iso_date(rec.get("date")), (rec.get("control") or "").strip(),
          (rec.get(ref_key) or "").strip(), (rec.get(desc_key) or "").strip(),
          (rec.get("remarks") or "").strip(), dc, cc,
          rec.get("deposit_num") or rec.get("deposit_number") or "", "",
          period, run_id))
    return h


def _set_state(con, table, h, status, engine_ref, run_id):
    # matched_run_id records the run that resolved the row — for 'internal'
    # too, so superseding that run reverts prior-period rows it resolved
    # (e.g. last month's closing sweep paired with this month's return).
    matched_run = run_id if status in ("matched", "internal") else None
    con.execute(f"""UPDATE {table} SET status=?, engine_ref=?,
                    matched_run_id=COALESCE(?, matched_run_id)
                    WHERE txn_hash=?""", (status, engine_ref, matched_run, h))


def _insert_match(con, property_code, run_id, rule, bank_hashes, gl_hashes,
                  bank_cents, gl_cents) -> int:
    cur = con.execute("""
        INSERT INTO matches (property_code, run_id, match_rule, is_manual,
                             matched_at, bank_total_cents, gl_total_cents)
        VALUES (?,?,?,0,?,?,?) RETURNING match_id
    """, (property_code, run_id, rule,
          datetime.now().isoformat(timespec="seconds"), bank_cents, gl_cents))
    mid = cur.lastrowid
    for h in bank_hashes:
        con.execute("UPDATE bank_txns SET match_id=? WHERE txn_hash=?", (mid, h))
    for h in gl_hashes:
        con.execute("UPDATE gl_txns SET match_id=? WHERE txn_hash=?", (mid, h))
    return mid


# ── engine-specific result walkers ───────────────────────────────────────────

def _persist_yardi(con, property_code, period, run_id, results, prior_refs):
    bank_hash_by_id, gl_hash_by_id = {}, {}
    occ_b, occ_g = defaultdict(int), defaultdict(int)
    for rec in results["all_bank"]:
        h = _upsert_bank(con, property_code, period, run_id, rec, occ_b,
                         known_hash=prior_refs.get(rec["id"]))
        bank_hash_by_id[rec["id"]] = h
        rule = rec.get("match_rule") or ""
        if rule in _YARDI_INTERNAL_BANK:
            status = "internal"
        elif rec.get("matched"):
            status = "matched"
        else:
            status = "unmatched"
        _set_state(con, "bank_txns", h, status, rec["id"], run_id)

    for rec in results["all_gl"]:
        h = _upsert_gl(con, property_code, period, run_id, rec, occ_g,
                       desc_key="desc", ref_key="ref",
                       known_hash=prior_refs.get(rec["id"]))
        gl_hash_by_id[rec["id"]] = h
        rule = rec.get("match_rule") or ""
        if rule == "Contra - GL":
            status = "internal"
        elif rec.get("matched"):
            status = "matched"
        else:
            status = "unmatched"
        _set_state(con, "gl_txns", h, status, rec["id"], run_id)

    # match groups = connected components of the bank↔GL match_ids graph
    # (covers 1 bank → N GL and the P9 N bank → 1 GL case)
    parent = {}

    def find(x):
        parent.setdefault(x, x)
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        parent[find(a)] = find(b)

    bank_by_id = {b["id"]: b for b in results["all_bank"]}
    gl_by_id = {g["id"]: g for g in results["all_gl"]}
    for b in results["all_bank"]:
        for gid in (b.get("match_ids") or []):
            union(("b", b["id"]), ("g", gid))
    for g in results["all_gl"]:
        for bid in (g.get("match_ids") or []):
            union(("g", g["id"]), ("b", bid))

    comps = defaultdict(lambda: {"b": [], "g": []})
    for node in list(parent):
        comps[find(node)][node[0]].append(node[1])

    for comp in comps.values():
        if not comp["b"] and not comp["g"]:
            continue
        rule = ""
        for bid in comp["b"]:
            rule = bank_by_id[bid].get("match_rule") or rule
        bank_cents = sum(db.to_cents(bank_by_id[i]["amount"]) for i in comp["b"])
        gl_cents = sum(db.to_cents(gl_by_id[i].get("debit") or 0)
                       - db.to_cents(gl_by_id[i].get("credit") or 0) for i in comp["g"])
        _insert_match(con, property_code, run_id, rule,
                      [bank_hash_by_id[i] for i in comp["b"]],
                      [gl_hash_by_id[i] for i in comp["g"]],
                      bank_cents, gl_cents)


def _persist_ma(con, property_code, period, run_id, results, prior_refs):
    bank_code, gl_code, _ = ma_assign_codes(results)
    bank_hashes, gl_hashes = [], []
    occ_b, occ_g = defaultdict(int), defaultdict(int)

    for bi, rec in enumerate(results["all_bank"]):
        h = _upsert_bank(con, property_code, period, run_id, rec, occ_b,
                         known_hash=prior_refs.get(f"B{bi}"))
        bank_hashes.append(h)
        code = abs(bank_code[bi])
        status = {1: "unmatched", 2: "matched", 3: "matched", 4: "internal"}[code]
        _set_state(con, "bank_txns", h, status, f"B{bi}", run_id)

    for gi, rec in enumerate(results["all_gl"]):
        h = _upsert_gl(con, property_code, period, run_id, rec, occ_g,
                       desc_key="person_desc", ref_key="reference",
                       known_hash=prior_refs.get(f"G{gi}"))
        gl_hashes.append(h)
        code = abs(gl_code[gi])
        status = {1: "unmatched", 2: "matched", 3: "matched", 4: "internal"}[code]
        _set_state(con, "gl_txns", h, status, f"G{gi}", run_id)

    for bi, info in results["matched_bank"].items():
        gl_idxs = info.get("gl_indices") or []
        if not gl_idxs:
            continue  # stagecoach internal marker
        b = results["all_bank"][bi]
        bank_cents = db.to_cents(b["amount"])
        gl_cents = sum(db.to_cents(results["all_gl"][gi].get("debit") or 0)
                       - db.to_cents(results["all_gl"][gi].get("credit") or 0)
                       for gi in gl_idxs)
        _insert_match(con, property_code, run_id, info["rule"],
                      [bank_hashes[bi]], [gl_hashes[gi] for gi in gl_idxs],
                      bank_cents, gl_cents)


def save_run(con, *, property_code: str, gl_type: str, period: str,
             results: dict, prior_source: str, prior_refs: dict | None = None,
             workdir: str = "", results_path: str = "", output_path: str = "",
             stats: dict | None = None) -> int:
    """Persist one reconciliation run atomically. Returns run_id.

    prior_refs maps engine record keys to existing txn_hashes for injected
    prior items (Yardi: "BANK-<row>"/"PREV-GL-<row>" ids; MA: "B<i>"/"G<i>"
    indices) — built by the pipeline, sourced from PriorItems.fetch().
    """
    prior_refs = prior_refs or {}
    try:
        _supersede(con, property_code, period)
        cur = con.execute("""
            INSERT INTO runs (property_code, gl_type, period, run_at, status,
                              prior_source, workdir, results_path, output_path, stats_json)
            VALUES (?,?,?,?,?,?,?,?,?,?) RETURNING run_id
        """, (property_code, gl_type, period,
              datetime.now().isoformat(timespec="seconds"), "complete", prior_source,
              str(workdir), str(results_path), str(output_path),
              json.dumps(stats or {})))
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
