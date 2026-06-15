#!/usr/bin/env python3
"""
Clear data from bank_rec.db.

Usage:
    python clear_db.py              # interactive menu
    python clear_db.py --all        # wipe everything (no prompt per-table)
    python clear_db.py --txns       # clear bank_txns + gl_txns only
    python clear_db.py --runs       # clear runs + matches + txns (keep properties)
"""

import argparse
import sqlite3
import sys
from pathlib import Path

DB_PATH = Path(__file__).resolve().parent / "data" / "bank_rec.db"

# Delete order matters — children before parents (FK constraints)
TABLES_ALL = ["bank_txns", "gl_txns", "matches", "runs", "properties"]
TABLES_RUNS = ["bank_txns", "gl_txns", "matches", "runs"]
TABLES_TXNS = ["bank_txns", "gl_txns"]


def row_counts(con: sqlite3.Connection) -> dict[str, int]:
    return {
        t: con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
        for t in TABLES_ALL
    }


def print_counts(counts: dict[str, int], header: str = "Current row counts:") -> None:
    print(f"\n{header}")
    for table, n in counts.items():
        print(f"  {table:<20} {n:>6} rows")
    print()


def confirm(prompt: str) -> bool:
    answer = input(f"{prompt} [y/N] ").strip().lower()
    return answer == "y"


def clear_tables(con: sqlite3.Connection, tables: list[str]) -> None:
    con.execute("PRAGMA foreign_keys = OFF")
    for table in tables:
        before = con.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
        con.execute(f"DELETE FROM {table}")
        print(f"  Cleared {table:<20} ({before} rows removed)")
    con.execute("PRAGMA foreign_keys = ON")
    con.commit()


def interactive(con: sqlite3.Connection) -> None:
    print_counts(row_counts(con))
    print("What would you like to clear?")
    print("  1. Transactions only      (bank_txns, gl_txns)")
    print("  2. Runs + transactions    (runs, matches, bank_txns, gl_txns)")
    print("  3. Everything             (all tables including properties)")
    print("  4. Cancel")
    choice = input("\nChoice [1-4]: ").strip()

    if choice == "1":
        if confirm("Clear all bank and GL transactions?"):
            clear_tables(con, TABLES_TXNS)
    elif choice == "2":
        if confirm("Clear runs, matches, and all transactions?"):
            clear_tables(con, TABLES_RUNS)
    elif choice == "3":
        if confirm("Wipe the entire database (including property configs)?"):
            clear_tables(con, TABLES_ALL)
    else:
        print("Cancelled.")
        return

    print_counts(row_counts(con), header="Row counts after clear:")


def main() -> None:
    parser = argparse.ArgumentParser(description="Clear bank_rec.db data")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--all",  action="store_true", help="Wipe all tables")
    group.add_argument("--runs", action="store_true", help="Clear runs, matches, and txns")
    group.add_argument("--txns", action="store_true", help="Clear bank_txns and gl_txns only")
    args = parser.parse_args()

    if not DB_PATH.exists():
        print(f"Database not found: {DB_PATH}")
        sys.exit(1)

    con = sqlite3.connect(DB_PATH)

    if args.all:
        print_counts(row_counts(con))
        if confirm("Wipe the entire database (including property configs)?"):
            clear_tables(con, TABLES_ALL)
            print_counts(row_counts(con), header="Row counts after clear:")
    elif args.runs:
        print_counts(row_counts(con))
        if confirm("Clear runs, matches, and all transactions?"):
            clear_tables(con, TABLES_RUNS)
            print_counts(row_counts(con), header="Row counts after clear:")
    elif args.txns:
        print_counts(row_counts(con))
        if confirm("Clear all bank and GL transactions?"):
            clear_tables(con, TABLES_TXNS)
            print_counts(row_counts(con), header="Row counts after clear:")
    else:
        interactive(con)

    con.close()


if __name__ == "__main__":
    main()
