#!/usr/bin/env python3
"""
Clear data from the PostgreSQL (Supabase) database.

Reads the connection URL from .streamlit/secrets.toml.

Usage:
    python clear_db.py              # interactive menu
    python clear_db.py --all        # wipe everything (no prompt per-table)
    python clear_db.py --txns       # clear bank_txns + gl_txns only
    python clear_db.py --runs       # clear runs + matches + txns (keep properties)
"""

import argparse
import sys
import tomllib
from pathlib import Path

import psycopg2

SECRETS_PATH = Path(__file__).resolve().parent / ".streamlit" / "secrets.toml"

# Delete order matters — children before parents (FK constraints)
TABLES_ALL  = ["bank_txns", "gl_txns", "matches", "runs", "properties"]
TABLES_RUNS = ["bank_txns", "gl_txns", "matches", "runs"]
TABLES_TXNS = ["bank_txns", "gl_txns"]


def get_url() -> str:
    if not SECRETS_PATH.exists():
        print(f"Secrets file not found: {SECRETS_PATH}")
        sys.exit(1)
    with open(SECRETS_PATH, "rb") as f:
        secrets = tomllib.load(f)
    try:
        return secrets["postgres"]["url"]
    except KeyError:
        print("No [postgres] url found in secrets.toml")
        sys.exit(1)


def row_counts(cur) -> dict[str, int]:
    counts = {}
    for t in TABLES_ALL:
        cur.execute(f"SELECT COUNT(*) FROM {t}")
        counts[t] = cur.fetchone()[0]
    return counts


def print_counts(counts: dict[str, int], header: str = "Current row counts:") -> None:
    print(f"\n{header}")
    for table, n in counts.items():
        print(f"  {table:<20} {n:>6} rows")
    print()


def confirm(prompt: str) -> bool:
    answer = input(f"{prompt} [y/N] ").strip().lower()
    return answer == "y"


def clear_tables(con, tables: list[str]) -> None:
    with con.cursor() as cur:
        for table in tables:
            cur.execute(f"SELECT COUNT(*) FROM {table}")
            before = cur.fetchone()[0]
            cur.execute(f"DELETE FROM {table}")
            print(f"  Cleared {table:<20} ({before} rows removed)")
    con.commit()


def interactive(con) -> None:
    with con.cursor() as cur:
        print_counts(row_counts(cur))
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

    with con.cursor() as cur:
        print_counts(row_counts(cur), header="Row counts after clear:")


def main() -> None:
    parser = argparse.ArgumentParser(description="Clear the Supabase PostgreSQL database")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--all",  action="store_true", help="Wipe all tables")
    group.add_argument("--runs", action="store_true", help="Clear runs, matches, and txns")
    group.add_argument("--txns", action="store_true", help="Clear bank_txns and gl_txns only")
    args = parser.parse_args()

    url = get_url()
    print(f"Connecting to PostgreSQL…")
    con = psycopg2.connect(url)

    if args.all:
        with con.cursor() as cur:
            print_counts(row_counts(cur))
        if confirm("Wipe the entire database (including property configs)?"):
            clear_tables(con, TABLES_ALL)
            with con.cursor() as cur:
                print_counts(row_counts(cur), header="Row counts after clear:")
    elif args.runs:
        with con.cursor() as cur:
            print_counts(row_counts(cur))
        if confirm("Clear runs, matches, and all transactions?"):
            clear_tables(con, TABLES_RUNS)
            with con.cursor() as cur:
                print_counts(row_counts(cur), header="Row counts after clear:")
    elif args.txns:
        with con.cursor() as cur:
            print_counts(row_counts(cur))
        if confirm("Clear all bank and GL transactions?"):
            clear_tables(con, TABLES_TXNS)
            with con.cursor() as cur:
                print_counts(row_counts(cur), header="Row counts after clear:")
    else:
        interactive(con)

    con.close()


if __name__ == "__main__":
    main()
